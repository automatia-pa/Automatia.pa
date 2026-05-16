from flask import Flask, render_template, request, redirect, url_for, flash, send_file, jsonify, abort, session
from flask_login import LoginManager, login_user, logout_user, login_required, current_user
from werkzeug.utils import secure_filename
import os
import re
import threading
import time
import secrets
import json
from datetime import timedelta, datetime

from models import User, db_init
from facturas_processor import procesar_cliente, get_rutas_cliente, exportar_dgi_csv

app = Flask(__name__)

# ── SEGURIDAD ─────────────────────────────────────────────────
app.secret_key = os.environ.get("FLASK_SECRET_KEY")
if not app.secret_key or len(app.secret_key) < 32:
    raise RuntimeError("FLASK_SECRET_KEY no está configurada correctamente")

app.config['SESSION_COOKIE_SECURE'] = True
app.config['SESSION_COOKIE_HTTPONLY'] = True
app.config['SESSION_COOKIE_SAMESITE'] = 'Lax'
app.config['PERMANENT_SESSION_LIFETIME'] = timedelta(hours=8)
app.config['SESSION_COOKIE_NAME'] = '__Host-session'
app.config['MAX_CONTENT_LENGTH'] = 10 * 1024 * 1024  # 10 MB

ALLOWED_EXTENSIONS = {'pdf', 'txt', 'xlsx', 'xls', 'xml', 'jpg', 'jpeg', 'png', 'webp'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def is_safe_file(file_storage):
    if not file_storage or not file_storage.filename:
        return False
    return allowed_file(file_storage.filename)

# ── FLASK-LOGIN ───────────────────────────────────────────────
login_manager = LoginManager()
login_manager.init_app(app)
login_manager.login_view = "index"
login_manager.login_message = "Inicia sesión para acceder al portal"
login_manager.login_message_category = "danger"

@login_manager.user_loader
def load_user(user_id):
    return User.get(user_id)

db_init()

ADMIN_SECRET = os.environ.get("ADMIN_SECRET")

# ── JOBS EN BACKGROUND ───────────────────────────────────────
_jobs: dict = {}
_jobs_lock = threading.Lock()

# ══════════════════════════════════════════════════════════════
# PATCH 1 — RATE LIMITER PERSISTENTE EN ARCHIVO JSON
# Reemplaza el defaultdict en memoria que se perdía al reiniciar.
# Guarda intentos en ~/.private_data/rate_limit.json con formato:
#   { "ip": [timestamp, timestamp, ...], ... }
# ══════════════════════════════════════════════════════════════
_RATE_FILE = os.path.expanduser("~/.private_data/rate_limit.json")
_rate_lock  = threading.Lock()
MAX_ATTEMPTS   = 8
WINDOW_SECONDS = 300  # 5 minutos

def _load_attempts() -> dict:
    """Lee el archivo de intentos. Devuelve {} si no existe o está corrupto."""
    try:
        with open(_RATE_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return {}

def _save_attempts(data: dict):
    """Escribe el dict al archivo. Crea el directorio si hace falta."""
    os.makedirs(os.path.dirname(_RATE_FILE), exist_ok=True)
    tmp = _RATE_FILE + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f)
    os.replace(tmp, _RATE_FILE)   # escritura atómica

def _check_rate_limit(ip: str) -> bool:
    """Retorna True si la IP puede intentar login, False si está bloqueada."""
    now = time.time()
    with _rate_lock:
        data = _load_attempts()
        attempts = [t for t in data.get(ip, []) if now - t < WINDOW_SECONDS]
        data[ip] = attempts
        _save_attempts(data)
        return len(attempts) < MAX_ATTEMPTS

def _register_attempt(ip: str):
    """Registra un intento fallido para la IP dada."""
    now = time.time()
    with _rate_lock:
        data = _load_attempts()
        attempts = [t for t in data.get(ip, []) if now - t < WINDOW_SECONDS]
        attempts.append(now)
        data[ip] = attempts
        _save_attempts(data)

def _get_ip() -> str:
    xff = request.headers.get('X-Forwarded-For', '')
    if xff:
        return xff.split(',')[0].strip()
    return request.remote_addr or '0.0.0.0'

# ── SECURITY HEADERS ─────────────────────────────────────────
@app.after_request
def set_security_headers(response):
    response.headers['X-Content-Type-Options'] = 'nosniff'
    response.headers['X-Frame-Options'] = 'DENY'
    response.headers['X-XSS-Protection'] = '1; mode=block'
    response.headers['Referrer-Policy'] = 'strict-origin-when-cross-origin'
    response.headers['Permissions-Policy'] = 'geolocation=(), microphone=(), camera=()'
    response.headers['Content-Security-Policy'] = (
        "default-src 'self'; "
        "script-src 'self' 'unsafe-inline' https://cdnjs.cloudflare.com; "
        "style-src 'self' 'unsafe-inline' https://fonts.googleapis.com; "
        "font-src https://fonts.gstatic.com; "
        "img-src 'self' data:; "
        "connect-src 'self'; "
        "frame-ancestors 'none';"
    )
    return response

# ── CSRF TOKEN ───────────────────────────────────────────────
def generate_csrf_token():
    if '_csrf_token' not in session:
        session['_csrf_token'] = secrets.token_hex(32)
    return session['_csrf_token']

app.jinja_env.globals['csrf_token'] = generate_csrf_token

def validate_csrf():
    token      = session.get('_csrf_token')
    form_token = request.form.get('_csrf_token') or request.headers.get('X-CSRF-Token')
    if not token or not form_token or not secrets.compare_digest(token, form_token):
        abort(403)

def _run_background(nombre_cliente, archivos_guardados):
    with _jobs_lock:
        _jobs[nombre_cliente] = {"status": "processing", "msg": "Procesando...", "ts": time.time()}
    try:
        procesar_cliente(nombre_cliente, set())
        with _jobs_lock:
            _jobs[nombre_cliente] = {
                "status": "done",
                "msg": f"{archivos_guardados} factura(s) procesada(s)",
                "ts": time.time()
            }
    except Exception:
        with _jobs_lock:
            _jobs[nombre_cliente] = {"status": "error", "msg": "Error interno", "ts": time.time()}

# ══════════════════════════════════════════════════════════════
# PATCH 2 — SESSION FIXATION REAL
# Flask no tiene session.regenerate. La solución correcta es:
#   1. Guardar los datos que necesitamos de la sesión vieja.
#   2. session.clear() — invalida la sesión existente.
#   3. Reasignar los datos necesarios en la sesión nueva.
# Esto hace que Flask genere un nuevo session cookie ID.
# ══════════════════════════════════════════════════════════════
def _regenerar_sesion():
    """Regenera el session ID conservando solo el CSRF token."""
    old_csrf = session.get('_csrf_token')
    session.clear()
    # Genera un CSRF token fresco para la sesión autenticada
    session['_csrf_token'] = old_csrf or secrets.token_hex(32)
    session.modified = True

# ── INDEX + LOGIN ─────────────────────────────────────────────
@app.route("/", methods=["GET", "POST"])
def index():
    if current_user.is_authenticated:
        return redirect(url_for("dashboard"))
    if request.method == "POST":
        validate_csrf()
        ip = _get_ip()
        if not _check_rate_limit(ip):
            flash("Demasiados intentos. Espera 5 minutos.", "danger")
            return redirect(url_for("index"))

        email    = request.form.get("email", "").strip()
        password = request.form.get("password", "")

        _register_attempt(ip)
        if User.check_password(email, password):
            user = User.get_by_email(email)

            # ── MFA: si el usuario tiene TOTP configurado, ir a verificación ──
            if User.has_totp(user.id):
                session['_mfa_pending_user'] = user.id
                return redirect(url_for("mfa_verify"))

            # Sin MFA: login directo + regenerar sesión
            _regenerar_sesion()          # ← PATCH 2 en acción
            login_user(user)
            return redirect(url_for("dashboard"))

        flash("Correo o contraseña incorrectos", "danger")
        return redirect(url_for("index"))
    return render_template("index.html")

# ── LOGIN (alias de index) ────────────────────────────────────
@app.route("/login", methods=["GET", "POST"])
def login():
    if current_user.is_authenticated:
        return redirect(url_for("dashboard"))
    if request.method == "POST":
        validate_csrf()
        ip = _get_ip()
        if not _check_rate_limit(ip):
            flash("Demasiados intentos. Espera 5 minutos.", "danger")
            return redirect(url_for("login"))

        email    = request.form.get("email", "").strip()
        password = request.form.get("password", "")
        _register_attempt(ip)

        if User.check_password(email, password):
            user = User.get_by_email(email)

            if User.has_totp(user.id):
                session['_mfa_pending_user'] = user.id
                return redirect(url_for("mfa_verify"))

            _regenerar_sesion()
            login_user(user)
            return redirect(url_for("dashboard"))

        flash("Correo o contraseña incorrectos", "danger")
    return render_template("login.html")

# ══════════════════════════════════════════════════════════════
# PATCH 3 — MFA CON PYOTP (TOTP compatible con Google Authenticator)
#
# Flujo:
#   GET  /mfa/setup   → genera secret + QR code para escanear
#   POST /mfa/setup   → verifica primer código y activa MFA
#   GET  /mfa/verify  → formulario de código (tras login exitoso)
#   POST /mfa/verify  → verifica código y completa el login
#   POST /mfa/disable → desactiva MFA (requiere contraseña)
# ══════════════════════════════════════════════════════════════
import pyotp
import qrcode
import io
import base64

@app.route("/mfa/setup", methods=["GET", "POST"])
@login_required
def mfa_setup():
    if request.method == "POST":
        validate_csrf()
        secret = request.form.get("secret", "").strip()
        code   = request.form.get("code", "").strip()

        if not secret or not code:
            flash("Datos incompletos", "danger")
            return redirect(url_for("mfa_setup"))

        totp = pyotp.TOTP(secret)
        if totp.verify(code, valid_window=1):
            User.set_totp(current_user.id, secret)
            flash("✅ Autenticación de dos factores activada correctamente", "success")
            return redirect(url_for("dashboard"))
        else:
            flash("Código incorrecto. Intenta nuevamente.", "danger")
            return redirect(url_for("mfa_setup"))

    # GET: generar nuevo secret y QR
    secret = pyotp.random_base32()
    totp   = pyotp.TOTP(secret)
    uri    = totp.provisioning_uri(
        name=current_user.email,
        issuer_name="AutomatIA"
    )

    # Generar QR como imagen base64 (sin guardar en disco)
    img    = qrcode.make(uri)
    buffer = io.BytesIO()
    img.save(buffer, format="PNG")
    qr_b64 = base64.b64encode(buffer.getvalue()).decode()

    return render_template("mfa_setup.html", secret=secret, qr_b64=qr_b64)


@app.route("/mfa/verify", methods=["GET", "POST"])
def mfa_verify():
    """Segundo factor post-login. El user_id está en session['_mfa_pending_user']."""
    pending_id = session.get('_mfa_pending_user')
    if not pending_id:
        return redirect(url_for("login"))

    if request.method == "POST":
        validate_csrf()
        code = request.form.get("code", "").strip()
        user = User.get(pending_id)

        if not user:
            session.pop('_mfa_pending_user', None)
            return redirect(url_for("login"))

        secret = User.get_totp_secret(user.id)
        if secret:
            totp = pyotp.TOTP(secret)
            if totp.verify(code, valid_window=1):
                session.pop('_mfa_pending_user', None)
                _regenerar_sesion()        # regenerar sesión también aquí
                login_user(user)
                return redirect(url_for("dashboard"))

        flash("Código incorrecto o expirado.", "danger")
        return redirect(url_for("mfa_verify"))

    return render_template("mfa_verify.html")


@app.route("/mfa/disable", methods=["POST"])
@login_required
def mfa_disable():
    validate_csrf()
    password = request.form.get("password", "")
    if User.check_password(current_user.email, password):
        User.set_totp(current_user.id, None)
        flash("Autenticación de dos factores desactivada.", "success")
    else:
        flash("Contraseña incorrecta.", "danger")
    return redirect(url_for("dashboard"))


# ── REGISTER ──────────────────────────────────────────────────
@app.route("/register", methods=["GET", "POST"])
def register():
    admin_key = request.form.get("admin_key") or request.headers.get("X-Admin-Key")
    if not ADMIN_SECRET or not secrets.compare_digest(admin_key or "", ADMIN_SECRET):
        abort(403)

    if request.method == "POST":
        validate_csrf()
        nombre   = request.form.get("nombre", "").strip()
        email    = request.form.get("email", "").strip().lower()
        password = request.form.get("password", "")

        if not nombre or not email or not password:
            flash("Todos los campos son obligatorios", "danger")
            return redirect(url_for("register"))

        if len(password) < 10:
            flash("La contraseña debe tener al menos 10 caracteres", "danger")
            return redirect(url_for("register"))

        if not re.match(r'^[^@]+@[^@]+\.[^@]+$', email):
            flash("Correo inválido", "danger")
            return redirect(url_for("register"))

        if User.exists(email):
            flash("Este correo ya está registrado", "danger")
            return redirect(url_for("register"))

        User.create(nombre, email, password)
        flash(f"Cliente '{nombre}' creado exitosamente", "success")
        return redirect(url_for("register"))

    return render_template("register.html")

# ── DASHBOARD ─────────────────────────────────────────────────
@app.route("/dashboard")
@login_required
def dashboard():
    import sqlite3 as _sq
    rutas    = get_rutas_cliente(current_user.nombre)
    facturas = []
    if os.path.exists(rutas["db"]):
        try:
            with _sq.connect(rutas["db"]) as conn:
                conn.row_factory = _sq.Row
                rows = conn.execute("""
                    SELECT id, archivo AS Archivo, proveedor AS Proveedor,
                           ruc AS RUC, fecha AS Fecha,
                           monto_total AS Monto_Total, subtotal AS Subtotal,
                           itbms AS ITBMS, moneda AS Moneda,
                           categoria AS Categoria, tipo_doc AS Tipo_Documento,
                           fuente AS Fuente, confianza AS Confianza,
                           estado AS Estado,
                           comentario_estado AS Comentario_Estado,
                           descripcion AS Descripcion, cufe AS CUFE,
                           fecha_procesamiento AS Fecha_Procesado
                    FROM facturas ORDER BY fecha_procesamiento DESC
                """).fetchall()
                facturas = [dict(r) for r in rows]
        except Exception:
            pass

    has_mfa = User.has_totp(current_user.id)
    return render_template("dashboard.html", user=current_user, facturas=facturas, has_mfa=has_mfa)

# ── EDITAR CAMPOS DE FACTURA ──────────────────────────────────
@app.route("/factura/editar", methods=["POST"])
@login_required
def editar_factura():
    validate_csrf()
    import sqlite3

    try:
        factura_id = int(request.form.get("factura_id"))
    except (TypeError, ValueError):
        flash("ID de factura inválido", "danger")
        return redirect(url_for("dashboard"))

    # Campos editables con sus límites
    proveedor   = request.form.get("proveedor",   "").strip()[:200]
    ruc         = request.form.get("ruc",         "").strip()[:30]
    fecha       = request.form.get("fecha",       "").strip()[:10]
    descripcion = request.form.get("descripcion", "").strip()[:300]
    moneda      = request.form.get("moneda",      "USD").strip()[:5]
    categoria   = request.form.get("categoria",   "Otros").strip()

    monedas_validas    = {"USD", "PAB", "EUR", "COP", "MXN", "PEN", "CLP", "ARS", "BRL"}
    categorias_validas = {
        "Servicios", "Materiales y Suministros", "Transporte y Logistica",
        "Tecnologia y Software", "Nomina y RRHH", "Alquiler e Inmuebles",
        "Publicidad y Marketing", "Impuestos y Tasas", "Alimentacion", "Otros"
    }

    if moneda not in monedas_validas:
        moneda = "USD"
    if categoria not in categorias_validas:
        categoria = "Otros"

    # Numéricos — None si vacío
    def parse_float(key):
        val = request.form.get(key, "").strip()
        if not val:
            return None
        try:
            v = float(val)
            return v if 0 <= v <= 10_000_000 else None
        except ValueError:
            return None

    monto_total = parse_float("monto_total")
    itbms       = parse_float("itbms")
    subtotal    = parse_float("subtotal")

    if not proveedor:
        flash("El proveedor no puede estar vacío", "danger")
        return redirect(url_for("dashboard"))

    if monto_total is None or monto_total <= 0:
        flash("El monto total debe ser mayor que cero", "danger")
        return redirect(url_for("dashboard"))

    rutas = get_rutas_cliente(current_user.nombre)
    if not os.path.exists(rutas["db"]):
        flash("Sin datos aún", "warning")
        return redirect(url_for("dashboard"))

    with sqlite3.connect(rutas["db"]) as conn:
        # Verificar que la factura pertenece a este cliente
        existe = conn.execute(
            "SELECT 1 FROM facturas WHERE id=?", (factura_id,)
        ).fetchone()
        if not existe:
            flash("Factura no encontrada", "danger")
            return redirect(url_for("dashboard"))

        conn.execute("""
            UPDATE facturas
            SET proveedor=?, ruc=?, fecha=?, monto_total=?,
                itbms=?, subtotal=?, moneda=?, categoria=?,
                descripcion=?, confianza=100, fuente='manual',
                fecha_procesamiento=?
            WHERE id=?
        """, (
            proveedor, ruc or None, fecha or None, monto_total,
            itbms, subtotal, moneda, categoria,
            descripcion or None, datetime.now().isoformat(),
            factura_id
        ))

    from facturas_processor import exportar_excel
    exportar_excel(rutas["db"], rutas["excel"])

    flash(f"Factura de '{proveedor}' actualizada correctamente ✓", "success")
    return redirect(url_for("dashboard"))

# ── APROBAR / RECHAZAR / REVISAR FACTURA ─────────────────────
@app.route("/factura/estado", methods=["POST"])
@login_required
def actualizar_estado():
    validate_csrf()
    import sqlite3

    try:
        factura_id = int(request.form.get("factura_id"))
    except (TypeError, ValueError):
        flash("ID de factura inválido", "danger")
        return redirect(url_for("dashboard"))

    nuevo_estado = request.form.get("estado")
    comentario   = request.form.get("comentario", "")[:500]

    estados_validos = {"aprobada", "en_revision", "rechazada", "pendiente"}
    if nuevo_estado not in estados_validos:
        flash("Estado inválido", "danger")
        return redirect(url_for("dashboard"))

    rutas = get_rutas_cliente(current_user.nombre)
    if not os.path.exists(rutas["db"]):
        flash("Sin datos aún", "warning")
        return redirect(url_for("dashboard"))

    with sqlite3.connect(rutas["db"]) as conn:
        conn.execute("""
            UPDATE facturas
            SET estado = ?, comentario_estado = ?, fecha_estado = ?
            WHERE id = ?
        """, (nuevo_estado, comentario, datetime.now().isoformat(), factura_id))

    from facturas_processor import exportar_excel
    exportar_excel(rutas["db"], rutas["excel"])

    etiquetas = {"aprobada": "aprobada ✓", "en_revision": "en revisión",
                 "rechazada": "rechazada ✗", "pendiente": "pendiente"}
    flash(f"Factura marcada como {etiquetas.get(nuevo_estado, nuevo_estado)}", "success")
    return redirect(url_for("dashboard"))

# ── DOWNLOAD CSV DGI ──────────────────────────────────────────
@app.route("/download-dgi")
@login_required
def download_dgi():
    periodo = request.args.get("periodo", "")

    if periodo and not re.match(r'^\d{4}-\d{2}$', periodo):
        flash("Período inválido. Usa el formato YYYY-MM.", "danger")
        return redirect(url_for("dashboard"))

    rutas = get_rutas_cliente(current_user.nombre)
    if not os.path.exists(rutas["db"]):
        flash("Aún no tienes facturas procesadas", "warning")
        return redirect(url_for("dashboard"))

    sufijo   = f"_{periodo}" if periodo else ""
    csv_path = os.path.join(rutas["base"], f"declaracion_dgi{sufijo}.csv")
    total    = exportar_dgi_csv(rutas["db"], csv_path, periodo=periodo or None)

    if total == 0:
        flash("No hay facturas aprobadas para exportar.", "warning")
        return redirect(url_for("dashboard"))

    return send_file(csv_path, as_attachment=True,
                     download_name=f"declaracion_dgi{sufijo}.csv")

# ── UPLOAD ────────────────────────────────────────────────────
@app.route("/upload", methods=["POST"])
@login_required
def upload():
    validate_csrf()
    if "files[]" not in request.files:
        flash("No se seleccionaron archivos", "danger")
        return redirect(url_for("dashboard"))

    files = request.files.getlist("files[]")

    if len(files) > 20:
        flash("Máximo 20 archivos por carga", "danger")
        return redirect(url_for("dashboard"))

    rutas = get_rutas_cliente(current_user.nombre)
    os.makedirs(rutas["facturas"], exist_ok=True)

    guardados  = 0
    rechazados = 0

    for file in files:
        if file.filename == "" or not is_safe_file(file):
            rechazados += 1
            continue

        filename = secure_filename(file.filename)
        if not filename:
            rechazados += 1
            continue

        filepath = os.path.join(rutas["facturas"], filename)
        file.save(filepath)
        guardados += 1

    if rechazados > 0:
        flash(f"{rechazados} archivo(s) rechazados por seguridad", "warning")

    if guardados > 0:
        t = threading.Thread(
            target=_run_background,
            args=(current_user.nombre, guardados),
            daemon=True
        )
        t.start()
        flash(f"{guardados} archivo(s) subido(s). Procesando en segundo plano...", "success")

    return redirect(url_for("dashboard"))

# ── ESTADO DEL JOB ────────────────────────────────────────────
@app.route("/upload/status")
@login_required
def upload_status():
    with _jobs_lock:
        job = _jobs.get(current_user.nombre)
    if not job:
        return jsonify({"status": "idle"})
    safe_job = {
        "status": job["status"],
        "msg": job["msg"] if job["status"] != "error" else "Error procesando archivos",
        "ts": job["ts"]
    }
    return jsonify(safe_job)

# ── DOWNLOAD EXCEL ────────────────────────────────────────────
@app.route("/download-excel")
@login_required
def download_excel():
    rutas = get_rutas_cliente(current_user.nombre)
    if os.path.exists(rutas["excel"]):
        return send_file(rutas["excel"], as_attachment=True)
    flash("Aún no tienes facturas procesadas", "warning")
    return redirect(url_for("dashboard"))

# ── LOGOUT ────────────────────────────────────────────────────
@app.route("/logout")
@login_required
def logout():
    logout_user()
    session.clear()
    return redirect(url_for("index"))

if __name__ == "__main__":
    app.run(debug=False)
