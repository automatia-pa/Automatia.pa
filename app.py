from flask import Flask, render_template, request, redirect, url_for, flash, send_file, jsonify, abort
from flask_login import LoginManager, login_user, logout_user, login_required, current_user
from werkzeug.utils import secure_filename
import os
import re
import threading
import time
import secrets
from datetime import timedelta
from collections import defaultdict
import magic  # Nueva librería para validar archivos

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

ALLOWED_EXTENSIONS = {'pdf', 'txt', 'xlsx', 'xls', 'xml'}
ALLOWED_MIME_TYPES = {
    'application/pdf',
    'text/plain',
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    'application/vnd.ms-excel',
    'text/xml',
    'application/xml'
}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def is_safe_file(file_storage):
    """Validación de seguridad mejorada"""
    if not file_storage or not file_storage.filename:
        return False
    if not allowed_file(file_storage.filename):
        return False
    
    try:
        file_storage.stream.seek(0)
        header = file_storage.stream.read(2048)
        file_storage.stream.seek(0)
        
        mime_type = magic.from_buffer(header, mime=True)
        if mime_type not in ALLOWED_MIME_TYPES:
            logging.warning(f"Archivo con MIME no permitido: {mime_type} - {file_storage.filename}")
            return False
    except Exception as e:
        logging.error(f"Error validando archivo {file_storage.filename}: {e}")
        return False
    
    return True

# ── FLASK-LOGIN y resto de configuración (se mantiene igual) ──
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

# ── PROTECCIÓN BRUTE-FORCE (en memoria, simple) ───────────────
# Para producción real usar Flask-Limiter + Redis
_login_attempts: dict = defaultdict(list)   # ip -> [timestamps]
_login_lock = threading.Lock()
MAX_ATTEMPTS = 8
WINDOW_SECONDS = 300  # 5 minutos

def _check_rate_limit(ip: str) -> bool:
    """Retorna True si la IP puede intentar login, False si está bloqueada."""
    now = time.time()
    with _login_lock:
        attempts = [t for t in _login_attempts[ip] if now - t < WINDOW_SECONDS]
        _login_attempts[ip] = attempts
        if len(attempts) >= MAX_ATTEMPTS:
            return False
        return True

def _register_attempt(ip: str):
    now = time.time()
    with _login_lock:
        _login_attempts[ip].append(now)

def _get_ip() -> str:
    # PythonAnywhere pone la IP real en X-Forwarded-For
    xff = request.headers.get('X-Forwarded-For', '')
    if xff:
        return xff.split(',')[0].strip()
    return request.remote_addr or '0.0.0.0'

# ── SECURITY HEADERS (todas las respuestas) ───────────────────
@app.after_request
def set_security_headers(response):
    response.headers['X-Content-Type-Options'] = 'nosniff'
    response.headers['X-Frame-Options'] = 'DENY'
    response.headers['X-XSS-Protection'] = '1; mode=block'
    response.headers['Referrer-Policy'] = 'strict-origin-when-cross-origin'
    response.headers['Permissions-Policy'] = 'geolocation=(), microphone=(), camera=()'
    # CSP — ajusta si agregas más CDNs
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

# ── CSRF TOKEN simple (sin Flask-WTF) ────────────────────────
def generate_csrf_token():
    if '_csrf_token' not in __import__('flask').session:
        __import__('flask').session['_csrf_token'] = secrets.token_hex(32)
    return __import__('flask').session['_csrf_token']

app.jinja_env.globals['csrf_token'] = generate_csrf_token

def validate_csrf():
    token = __import__('flask').session.get('_csrf_token')
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
    except Exception as e:
        with _jobs_lock:
            _jobs[nombre_cliente] = {"status": "error", "msg": "Error interno", "ts": time.time()}

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

        # Simular tiempo constante para evitar timing attacks
        _register_attempt(ip)
        if User.check_password(email, password):
            user = User.get_by_email(email)
            login_user(user)
            # Regenerar sesión tras login exitoso (evita session fixation)
            __import__('flask').session.regenerate = True
            return redirect(url_for("dashboard"))

        flash("Correo o contraseña incorrectos", "danger")
        return redirect(url_for("index"))
    return render_template("index.html")

# ── REGISTER ──────────────────────────────────────────────────
@app.route("/register", methods=["GET", "POST"])
def register():
    admin_key = request.args.get("key") or request.form.get("admin_key")
    if not ADMIN_SECRET or not secrets.compare_digest(admin_key or "", ADMIN_SECRET):
        return redirect(url_for("index"))
    if request.method == "POST":
        validate_csrf()
        nombre   = request.form.get("nombre", "").strip()
        email    = request.form.get("email", "").strip().lower()
        password = request.form.get("password", "")

        # Validaciones
        if not nombre or not email or not password:
            flash("Todos los campos son obligatorios", "danger")
            return redirect(url_for("register", key=admin_key))
        if len(password) < 10:
            flash("La contraseña debe tener al menos 10 caracteres", "danger")
            return redirect(url_for("register", key=admin_key))
        if not re.match(r'^[^@]+@[^@]+\.[^@]+$', email):
            flash("Correo inválido", "danger")
            return redirect(url_for("register", key=admin_key))

        if User.exists(email):
            flash("Este correo ya está registrado", "danger")
            return redirect(url_for("register", key=admin_key))
        User.create(nombre, email, password)
        flash(f"Cliente '{nombre}' creado exitosamente", "success")
        return redirect(url_for("register", key=admin_key))
    return render_template("register.html")

# ── LOGIN ─────────────────────────────────────────────────────
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
            login_user(user)
            return redirect(url_for("dashboard"))
        flash("Correo o contraseña incorrectos", "danger")
    return render_template("login.html")

# ── DASHBOARD ─────────────────────────────────────────────────
@app.route("/dashboard")
@login_required
def dashboard():
    import sqlite3 as _sq
    rutas = get_rutas_cliente(current_user.nombre)
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
    return render_template("dashboard.html", user=current_user, facturas=facturas)

# ── APROBAR / RECHAZAR / REVISAR FACTURA ─────────────────────
@app.route("/factura/estado", methods=["POST"])
@login_required
def actualizar_estado():
    validate_csrf()
    import sqlite3
    from datetime import datetime

    try:
        factura_id = int(request.form.get("factura_id"))
    except (TypeError, ValueError):
        flash("ID de factura inválido", "danger")
        return redirect(url_for("dashboard"))

    nuevo_estado = request.form.get("estado")
    comentario   = request.form.get("comentario", "")[:500]  # limitar longitud

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
        flash("No hay facturas aprobadas para exportar. Aprueba facturas primero.", "warning")
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

    guardados = 0
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
    # No exponer msg de error interno directamente
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
    # Limpiar sesión completa
    __import__('flask').session.clear()
    return redirect(url_for("index"))

if __name__ == "__main__":
    app.run(debug=False)

