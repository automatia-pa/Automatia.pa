from flask import Flask, render_template, request, redirect, url_for, flash, send_file, jsonify, abort, session, g
from flask_login import LoginManager, login_user, logout_user, login_required, current_user
from werkzeug.utils import secure_filename
import os
import re
import threading
import time
import secrets
import json
import sqlite3
from datetime import timedelta, datetime

from models import User, db_init
from facturas_processor import (procesar_cliente, get_rutas_cliente,
                                exportar_dgi_csv, validar_ruc_dgi,
                                exportar_formulario103, calcular_itbms_esperado,
                                verificar_itbms)

app = Flask(__name__)

# ── SEGURIDAD ─────────────────────────────────────────────────
app.secret_key = os.environ.get("FLASK_SECRET_KEY")
if not app.secret_key or len(app.secret_key) < 32:
    raise RuntimeError("FLASK_SECRET_KEY no está configurada correctamente")

app.config['SESSION_COOKIE_SECURE'] = True
app.config['SESSION_COOKIE_HTTPONLY'] = True
app.config['SESSION_COOKIE_SAMESITE'] = 'Lax'
app.config['PERMANENT_SESSION_LIFETIME'] = timedelta(hours=8)
# FIX: Eliminado SESSION_COOKIE_NAME = '__Host-session'
# PythonAnywhere usa HTTPS con proxy, pero el prefijo __Host- requiere
# que la cookie NO tenga atributo Domain y esté en la raíz del path.
# Flask-Login no garantiza esto, lo que causaba que la cookie se rechazara
# silenciosamente en algunos browsers → sesiones que no persistían.
app.config['SESSION_COOKIE_NAME'] = 'session'
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

# ══════════════════════════════════════════════════════════════
# FIX 1 — IP REAL EN PYTHONANYWHERE
#
# PythonAnywhere pone la IP real del cliente en X-Forwarded-For,
# pero SIEMPRE desde su propio proxy interno — no hay riesgo de
# spoofing porque las requests externas no llegan directamente a
# tu app WSGI.
#
# El problema original: si X-Forwarded-For venía con múltiples IPs
# (ej: "1.2.3.4, 10.0.0.1"), se tomaba la primera sin validar.
# Un atacante podría inyectar una IP falsa como primer valor si
# hubiera otro proxy delante.
#
# Solución para PythonAnywhere: tomar el ÚLTIMO valor de la cadena
# X-Forwarded-For, que es el que añade el proxy confiable de PA.
# Si no existe el header, usar remote_addr.
# ══════════════════════════════════════════════════════════════
# IPs internas de PythonAnywhere (rango de su proxy WSGI)
# Estas son las únicas que deberían aparecer en remote_addr
_PYTHONANYWHERE_PROXY_NETS = (
    "10.",
    "172.16.", "172.17.", "172.18.", "172.19.",
    "172.20.", "172.21.", "172.22.", "172.23.",
    "172.24.", "172.25.", "172.26.", "172.27.",
    "172.28.", "172.29.", "172.30.", "172.31.",
    "192.168.",
    "127.",
)

def _get_ip() -> str:
    """
    Extrae la IP real del cliente, segura para PythonAnywhere.
    Toma el último valor de X-Forwarded-For (añadido por el proxy de PA)
    solo si remote_addr es una IP interna (de confianza).
    Si remote_addr es pública → usarla directamente (dev/otro entorno).
    """
    remote = request.remote_addr or "0.0.0.0"
    xff = request.headers.get("X-Forwarded-For", "").strip()

    # Solo confiar en XFF si quien nos lo entrega es el proxy interno
    if xff and any(remote.startswith(net) for net in _PYTHONANYWHERE_PROXY_NETS):
        # Tomar el ÚLTIMO valor — es el que añadió el proxy confiable
        ips = [ip.strip() for ip in xff.split(",")]
        return ips[-1] if ips else remote

    return remote


# ══════════════════════════════════════════════════════════════
# RATE LIMITER — PERSISTENTE EN SQLITE
#
# FIX 2: El rate limiter original usaba un archivo JSON con escritura
# completa en cada request, lo que es lento y puede corromperse bajo
# concurrencia. Reemplazado por SQLite con WAL, que es atómico,
# concurrente y mucho más rápido para lecturas/escrituras parciales.
#
# La tabla rate_limit guarda (ip TEXT, ts REAL) — un registro por
# intento. _check_rate_limit limpia los registros viejos y cuenta
# los que quedan en la ventana activa.
# ══════════════════════════════════════════════════════════════
_RATE_DB_PATH = os.path.expanduser("~/.private_data/rate_limit.db")
_rate_lock    = threading.Lock()
MAX_ATTEMPTS   = 8
WINDOW_SECONDS = 300  # 5 minutos

def _init_rate_db():
    os.makedirs(os.path.dirname(_RATE_DB_PATH), exist_ok=True)
    with sqlite3.connect(_RATE_DB_PATH) as conn:
        conn.execute("PRAGMA journal_mode=WAL")
        conn.execute("""
            CREATE TABLE IF NOT EXISTS rate_limit (
                ip   TEXT NOT NULL,
                ts   REAL NOT NULL
            )
        """)
        conn.execute("CREATE INDEX IF NOT EXISTS idx_rate_ip ON rate_limit(ip, ts)")
        conn.commit()

_init_rate_db()

def _check_rate_limit(ip: str) -> bool:
    """Retorna True si la IP puede intentar login, False si está bloqueada."""
    now = time.time()
    cutoff = now - WINDOW_SECONDS
    with _rate_lock:
        with sqlite3.connect(_RATE_DB_PATH) as conn:
            conn.execute("PRAGMA journal_mode=WAL")
            # Limpiar intentos fuera de la ventana
            conn.execute("DELETE FROM rate_limit WHERE ts < ?", (cutoff,))
            count = conn.execute(
                "SELECT COUNT(*) FROM rate_limit WHERE ip=? AND ts >= ?",
                (ip, cutoff)
            ).fetchone()[0]
            conn.commit()
    return count < MAX_ATTEMPTS

def _register_attempt(ip: str):
    """Registra un intento fallido para la IP dada."""
    now = time.time()
    with _rate_lock:
        with sqlite3.connect(_RATE_DB_PATH) as conn:
            conn.execute("PRAGMA journal_mode=WAL")
            conn.execute("INSERT INTO rate_limit (ip, ts) VALUES (?, ?)", (ip, now))
            conn.commit()


# ══════════════════════════════════════════════════════════════
# FIX 3 — JOBS PERSISTENTES EN SQLITE (sin Redis)
#
# El problema: _jobs era un dict en memoria del proceso. En producción
# con múltiples workers (o tras un reinicio del worker en PythonAnywhere),
# el estado del job se perdía y el cliente veía "idle" para siempre.
#
# Solución: tabla jobs_status en la misma SQLite del rate limiter.
# Guardamos status, msg y ts por nombre de cliente.
# Los jobs de más de 24h se limpian automáticamente.
# ══════════════════════════════════════════════════════════════
_JOBS_DB_PATH = _RATE_DB_PATH  # reutilizamos la misma DB
_jobs_lock    = threading.Lock()

def _init_jobs_db():
    with sqlite3.connect(_JOBS_DB_PATH) as conn:
        conn.execute("PRAGMA journal_mode=WAL")
        conn.execute("""
            CREATE TABLE IF NOT EXISTS jobs_status (
                cliente TEXT PRIMARY KEY,
                status  TEXT NOT NULL,
                msg     TEXT NOT NULL,
                ts      REAL NOT NULL
            )
        """)
        conn.commit()

_init_jobs_db()

def _job_set(cliente: str, status: str, msg: str):
    now = time.time()
    with _jobs_lock:
        with sqlite3.connect(_JOBS_DB_PATH) as conn:
            conn.execute("PRAGMA journal_mode=WAL")
            conn.execute("""
                INSERT INTO jobs_status (cliente, status, msg, ts)
                VALUES (?, ?, ?, ?)
                ON CONFLICT(cliente) DO UPDATE SET status=excluded.status,
                    msg=excluded.msg, ts=excluded.ts
            """, (cliente, status, msg, now))
            conn.commit()

def _job_get(cliente: str) -> dict | None:
    cutoff = time.time() - 86400  # limpiar jobs > 24h
    with _jobs_lock:
        with sqlite3.connect(_JOBS_DB_PATH) as conn:
            conn.execute("PRAGMA journal_mode=WAL")
            conn.execute("DELETE FROM jobs_status WHERE ts < ?", (cutoff,))
            row = conn.execute(
                "SELECT status, msg, ts FROM jobs_status WHERE cliente=?", (cliente,)
            ).fetchone()
            conn.commit()
    if row:
        return {"status": row[0], "msg": row[1], "ts": row[2]}
    return None


# ══════════════════════════════════════════════════════════════
# FIX 4 — CSP CON NONCE (eliminar unsafe-inline)
#
# El problema: 'unsafe-inline' en script-src anula la protección
# contra XSS del CSP — cualquier script inyectado en el HTML
# se ejecutaría igual.
#
# Solución: generar un nonce criptográfico por request y pasarlo
# a Jinja2 como variable global. Cada <script> y <style> del
# template debe incluir el atributo nonce="{{ csp_nonce() }}".
#
# ACCIÓN REQUERIDA EN LOS TEMPLATES:
#   Reemplazar: <script>
#   Por:        <script nonce="{{ csp_nonce() }}">
#   Reemplazar: <style>
#   Por:        <style nonce="{{ csp_nonce() }}">
# ══════════════════════════════════════════════════════════════
def _generate_nonce() -> str:
    return secrets.token_urlsafe(16)

@app.before_request
def _set_csp_nonce():
    g.csp_nonce = _generate_nonce()

def _get_csp_nonce() -> str:
    return getattr(g, "csp_nonce", "")

app.jinja_env.globals["csp_nonce"] = _get_csp_nonce

# ── SECURITY HEADERS ─────────────────────────────────────────
@app.after_request
def set_security_headers(response):
    nonce = _get_csp_nonce()
    response.headers['X-Content-Type-Options'] = 'nosniff'
    response.headers['X-Frame-Options'] = 'DENY'
    response.headers['X-XSS-Protection'] = '1; mode=block'
    response.headers['Referrer-Policy'] = 'strict-origin-when-cross-origin'
    response.headers['Permissions-Policy'] = 'geolocation=(), microphone=(), camera=()'
    response.headers['Content-Security-Policy'] = (
        f"default-src 'self'; "
        f"script-src 'self' 'nonce-{nonce}' https://cdnjs.cloudflare.com; "
        f"style-src 'self' 'nonce-{nonce}' https://fonts.googleapis.com; "
        f"font-src https://fonts.gstatic.com; "
        f"img-src 'self' data:; "
        f"connect-src 'self'; "
        f"frame-ancestors 'none';"
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

# ── BACKGROUND JOB (ahora con persistencia en SQLite) ─────────
def _run_background(nombre_cliente, archivos_guardados):
    _job_set(nombre_cliente, "processing", "Procesando...")
    try:
        procesar_cliente(nombre_cliente, set())
        _job_set(nombre_cliente, "done", f"{archivos_guardados} factura(s) procesada(s)")
    except Exception as e:
        import logging, traceback
        logging.error(f"Error en _run_background [{nombre_cliente}]: {e}\n{traceback.format_exc()}")
        _job_set(nombre_cliente, "error", "Error interno")


# ══════════════════════════════════════════════════════════════
# SESSION FIXATION — igual que antes (correcto)
# ══════════════════════════════════════════════════════════════
def _regenerar_sesion():
    """Regenera el session ID conservando solo el CSRF token."""
    old_csrf = session.get('_csrf_token')
    session.clear()
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

            if User.has_totp(user.id):
                session['_mfa_pending_user'] = user.id
                return redirect(url_for("mfa_verify"))

            _regenerar_sesion()
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

# ── MFA ───────────────────────────────────────────────────────
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

    secret = pyotp.random_base32()
    totp   = pyotp.TOTP(secret)
    uri    = totp.provisioning_uri(
        name=current_user.email,
        issuer_name="AutomatIA"
    )

    img    = qrcode.make(uri)
    buffer = io.BytesIO()
    img.save(buffer, format="PNG")
    qr_b64 = base64.b64encode(buffer.getvalue()).decode()

    return render_template("mfa_setup.html", secret=secret, qr_b64=qr_b64)


@app.route("/mfa/verify", methods=["GET", "POST"])
def mfa_verify():
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
                _regenerar_sesion()
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
# FIX: El GET original aplicaba la validación de admin_key ANTES de
# leer request.form, pero en un GET el form está vacío → admin_key
# siempre era None → abort(403) inmediato si intentabas cargar la página.
# Ahora el GET solo valida si ADMIN_SECRET está configurado y retorna
# el form. El POST sí valida la clave antes de crear el usuario.
@app.route("/register", methods=["GET", "POST"])
def register():
    if not ADMIN_SECRET:
        # Si no hay ADMIN_SECRET configurado, registro completamente cerrado
        abort(403)

    if request.method == "POST":
        validate_csrf()
        admin_key = request.form.get("admin_key") or request.headers.get("X-Admin-Key")
        if not secrets.compare_digest(admin_key or "", ADMIN_SECRET):
            abort(403)

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

    # GET: mostrar el form (la clave la ingresan en el form mismo)
    return render_template("register.html")

# ── DASHBOARD ─────────────────────────────────────────────────
@app.route("/dashboard")
@login_required
def dashboard():
    rutas    = get_rutas_cliente(current_user.nombre)
    facturas = []
    if os.path.exists(rutas["db"]):
        try:
            with sqlite3.connect(rutas["db"]) as conn:
                conn.row_factory = sqlite3.Row
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

    # ── Insights accionables ──────────────────────────────────
    insights = {"pendientes_antiguas": 0, "baja_confianza": 0,
                "posibles_duplicados": 0, "aprobadas_hoy": 0}
    hoy = datetime.now().date()
    vistos = {}  # clave: (proveedor_lower, monto_redondeado, fecha) → count

    for f in facturas:
        estado  = f.get("Estado", "pendiente") or "pendiente"
        conf    = f.get("Confianza", 100) or 100
        fp      = f.get("Fecha_Procesado") or ""
        adv     = f.get("Advertencias_Fiscales") or ""  # puede venir None

        # Facturas pendientes sin tocar hace más de 7 días
        if estado == "pendiente" and fp:
            try:
                dias = (hoy - datetime.fromisoformat(fp[:10]).date()).days
                if dias >= 7:
                    insights["pendientes_antiguas"] += 1
            except (ValueError, TypeError):
                pass

        # Confianza baja y aún pendiente
        if conf < 70 and estado == "pendiente":
            insights["baja_confianza"] += 1

        # Aprobadas hoy
        if estado == "aprobada" and fp and fp[:10] == hoy.isoformat():
            insights["aprobadas_hoy"] += 1

        # Posibles duplicados por contenido (mismo proveedor+monto+fecha)
        prov  = (f.get("Proveedor") or "").strip().lower()
        monto = round(float(f.get("Monto_Total") or 0), 0)
        fecha = (f.get("Fecha") or "").strip()
        if prov and monto and fecha:
            key = (prov, monto, fecha)
            vistos[key] = vistos.get(key, 0) + 1

    insights["posibles_duplicados"] = sum(1 for v in vistos.values() if v > 1)

    return render_template("dashboard.html", user=current_user, facturas=facturas,
                           has_mfa=has_mfa, insights=insights)

# ── VER FACTURA ORIGINAL ─────────────────────────────────────
@app.route("/factura/ver/<int:factura_id>")
@login_required
def ver_factura(factura_id):
    """Sirve el archivo original de la factura para verlo en el browser."""
    rutas = get_rutas_cliente(current_user.nombre)
    if not os.path.exists(rutas["db"]):
        abort(404)

    with sqlite3.connect(rutas["db"]) as conn:
        row = conn.execute(
            "SELECT archivo FROM facturas WHERE id=?", (factura_id,)
        ).fetchone()

    if not row:
        abort(404)

    # Seguridad: basename previene path traversal
    archivo = os.path.basename(row[0] or "")
    if not archivo:
        abort(404)

    # Buscar en procesados primero, luego en la carpeta activa
    for carpeta in [rutas["procesados"], rutas["facturas"]]:
        ruta = os.path.join(carpeta, archivo)
        if os.path.exists(ruta):
            return send_file(ruta, as_attachment=False)

    abort(404)


# ── EDITAR CAMPOS DE FACTURA ──────────────────────────────────
@app.route("/factura/editar", methods=["POST"])
@login_required
def editar_factura():
    validate_csrf()

    try:
        factura_id = int(request.form.get("factura_id"))
    except (TypeError, ValueError):
        flash("ID de factura inválido", "danger")
        return redirect(url_for("dashboard"))

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

# ── ESTADO DEL JOB (ahora lee desde SQLite) ──────────────────
@app.route("/upload/status")
@login_required
def upload_status():
    job = _job_get(current_user.nombre)
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

# ══════════════════════════════════════════════════════════════
# VALIDACIÓN DE RUC CONTRA DGI — endpoint AJAX
# Llamado desde el dashboard cuando el usuario escribe un RUC.
# Devuelve JSON: {ok, nombre, error, fuente}
# ══════════════════════════════════════════════════════════════
@app.route("/api/validar-ruc")
@login_required
def api_validar_ruc():
    ruc = request.args.get("ruc", "").strip()
    if not ruc or len(ruc) > 30:
        return jsonify({"ok": False, "nombre": "", "error": "RUC inválido", "fuente": "error"})

    resultado = validar_ruc_dgi(ruc)
    # Nunca devolver datos internos del servidor en el error
    resultado.pop("ts", None)
    return jsonify(resultado)


# ══════════════════════════════════════════════════════════════
# VERIFICACIÓN RÁPIDA DE ITBMS — endpoint AJAX
# Recibe subtotal + itbms + descripcion, responde si el ITBMS es correcto.
# ══════════════════════════════════════════════════════════════
@app.route("/api/verificar-itbms")
@login_required
def api_verificar_itbms():
    try:
        subtotal    = float(request.args.get("subtotal", 0))
        itbms       = float(request.args.get("itbms", 0))
        monto_total = float(request.args.get("monto_total", 0))
        descripcion = request.args.get("descripcion", "")[:200]
    except (ValueError, TypeError):
        return jsonify({"ok": False, "error": "Parámetros inválidos"})

    esperado, tasa, categoria = calcular_itbms_esperado(subtotal, descripcion)
    diferencia = abs(itbms - esperado)
    ok = diferencia <= 0.02

    return jsonify({
        "ok":         ok,
        "esperado":   round(esperado, 2),
        "declarado":  round(itbms, 2),
        "diferencia": round(diferencia, 2),
        "tasa_pct":   int(tasa * 100),
        "categoria":  categoria,
        "mensaje":    (
            f"ITBMS correcto ({int(tasa*100)}%)" if ok else
            f"ITBMS posiblemente incorrecto: esperado B/.{esperado:.2f} "
            f"({int(tasa*100)}% de B/.{subtotal:.2f}), declarado B/.{itbms:.2f}"
        ),
    })


# ══════════════════════════════════════════════════════════════
# FORMULARIO 103 — DESCARGA
# Genera el F103 precompletado con las facturas aprobadas del cliente.
# ══════════════════════════════════════════════════════════════
@app.route("/download-formulario103")
@login_required
def download_formulario103():
    periodo = request.args.get("periodo", "").strip()

    if periodo and not re.match(r'^\d{4}-\d{2}$', periodo):
        flash("Período inválido. Usa el formato YYYY-MM.", "danger")
        return redirect(url_for("dashboard"))

    rutas = get_rutas_cliente(current_user.nombre)
    if not os.path.exists(rutas["db"]):
        flash("Aún no tienes facturas procesadas", "warning")
        return redirect(url_for("dashboard"))

    sufijo   = f"_{periodo}" if periodo else ""
    f103_path = os.path.join(rutas["base"], f"formulario103{sufijo}.xlsx")

    # RUC del cliente: buscamos en sus propias facturas como receptor, o dejamos vacío
    ruc_empresa = ""
    try:
        with sqlite3.connect(rutas["db"]) as conn:
            row = conn.execute(
                "SELECT ruc_receptor FROM facturas WHERE ruc_receptor IS NOT NULL LIMIT 1"
            ).fetchone()
            if row:
                ruc_empresa = row[0] or ""
    except Exception:
        pass

    resultado = exportar_formulario103(
        db_path       = rutas["db"],
        output_path   = f103_path,
        nombre_empresa= current_user.nombre,
        ruc_empresa   = ruc_empresa,
        periodo       = periodo or "",
    )

    if "error" in resultado:
        flash(f"Error generando el Formulario 103: {resultado['error']}", "danger")
        return redirect(url_for("dashboard"))

    if resultado.get("total_proveedores", 0) == 0:
        flash("No hay facturas aprobadas para incluir en el Formulario 103.", "warning")
        return redirect(url_for("dashboard"))

    flash(
        f"Formulario 103 generado: {resultado['total_proveedores']} proveedores, "
        f"Total compras B/.{resultado['total_compras']:,.2f}, "
        f"ITBMS B/.{resultado['total_itbms']:,.2f}",
        "success"
    )
    return send_file(f103_path, as_attachment=True,
                     download_name=f"formulario103{sufijo}.xlsx")


# ── LOGOUT ────────────────────────────────────────────────────
@app.route("/logout")
@login_required
def logout():
    logout_user()
    session.clear()
    return redirect(url_for("index"))

if __name__ == "__main__":
    app.run(debug=False)
