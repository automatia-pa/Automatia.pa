from flask import Flask, render_template, request, redirect, url_for, flash, send_file, jsonify
from flask_login import LoginManager, login_user, logout_user, login_required, current_user
from werkzeug.utils import secure_filename
import os
import re
import threading
import time
from dotenv import load_dotenv

from models import User, db_init
from facturas_processor import procesar_cliente, get_rutas_cliente, exportar_dgi_csv

load_dotenv()

app = Flask(__name__)

app.secret_key = os.environ.get("FLASK_SECRET_KEY")
app.config['SESSION_COOKIE_SECURE'] = True
app.config['SESSION_COOKIE_HTTPONLY'] = True
app.config['SESSION_COOKIE_SAMESITE'] = 'Lax'
if not app.secret_key:
    raise RuntimeError("FLASK_SECRET_KEY no está configurado en .env")

app.config['MAX_CONTENT_LENGTH'] = 10 * 1024 * 1024

ALLOWED_EXTENSIONS = {'pdf', 'txt', 'xlsx', 'xls', 'xml'}

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

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

_jobs: dict = {}
_jobs_lock = threading.Lock()

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
            _jobs[nombre_cliente] = {"status": "error", "msg": str(e), "ts": time.time()}

# ── INDEX + LOGIN ────────────────────────────────────────────
@app.route("/", methods=["GET", "POST"])
def index():
    if current_user.is_authenticated:
        return redirect(url_for("dashboard"))
    if request.method == "POST":
        email    = request.form.get("email", "").strip()
        password = request.form.get("password", "")
        if User.check_password(email, password):
            user = User.get_by_email(email)
            login_user(user)
            return redirect(url_for("dashboard"))
        flash("Correo o contraseña incorrectos", "danger")
        return redirect(url_for("index"))
    return render_template("index.html")

# ── REGISTER ─────────────────────────────────────────────────
@app.route("/register", methods=["GET", "POST"])
def register():
    admin_key = request.args.get("key") or request.form.get("admin_key")
    if not ADMIN_SECRET or admin_key != ADMIN_SECRET:
        return redirect(url_for("index"))
    if request.method == "POST":
        nombre   = request.form.get("nombre", "").strip()
        email    = request.form.get("email", "").strip()
        password = request.form.get("password", "")
        if User.exists(email):
            flash("Este correo ya está registrado", "danger")
            return redirect(url_for("register", key=admin_key))
        User.create(nombre, email, password)
        flash(f"Cliente '{nombre}' creado exitosamente", "success")
        return redirect(url_for("register", key=admin_key))
    return render_template("register.html")

# ── LOGIN ────────────────────────────────────────────────────
@app.route("/login", methods=["GET", "POST"])
def login():
    if current_user.is_authenticated:
        return redirect(url_for("dashboard"))
    if request.method == "POST":
        email    = request.form.get("email", "").strip()
        password = request.form.get("password", "")
        if User.check_password(email, password):
            user = User.get_by_email(email)
            login_user(user)
            return redirect(url_for("dashboard"))
        flash("Correo o contraseña incorrectos", "danger")
    return render_template("login.html")

# ── DASHBOARD ────────────────────────────────────────────────
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
    import sqlite3
    from datetime import datetime

    # Validar factura_id como entero
    try:
        factura_id = int(request.form.get("factura_id"))
    except (TypeError, ValueError):
        flash("ID de factura inválido", "danger")
        return redirect(url_for("dashboard"))

    nuevo_estado = request.form.get("estado")
    comentario   = request.form.get("comentario", "")

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

# ── DOWNLOAD CSV DGI ─────────────────────────────────────────
@app.route("/download-dgi")
@login_required
def download_dgi():
    periodo = request.args.get("periodo", "")

    # Validar formato de periodo
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

# ── UPLOAD ───────────────────────────────────────────────────
@app.route("/upload", methods=["POST"])
@login_required
def upload():
    if "files[]" not in request.files:
        flash("No se seleccionaron archivos", "danger")
        return redirect(url_for("dashboard"))

    files = request.files.getlist("files[]")
    rutas = get_rutas_cliente(current_user.nombre)

    os.makedirs(rutas["facturas"],   exist_ok=True)
    os.makedirs(rutas["procesados"], exist_ok=True)
    os.makedirs(rutas["error"],      exist_ok=True)

    guardados  = 0
    rechazados = 0
    for file in files:
        if file.filename == "":
            continue
        if not allowed_file(file.filename):
            rechazados += 1
            continue
        filename = secure_filename(file.filename)
        filepath = os.path.join(rutas["facturas"], filename)
        file.save(filepath)
        guardados += 1

    if rechazados > 0:
        flash(f"{rechazados} archivo(s) rechazados (solo PDF, TXT, XLSX, XML)", "warning")

    if guardados > 0:
        t = threading.Thread(
            target=_run_background,
            args=(current_user.nombre, guardados),
            daemon=True
        )
        t.start()
        flash(f"{guardados} archivo(s) subido(s). Procesando en segundo plano...", "success")

    return redirect(url_for("dashboard"))

# ── ESTADO DEL JOB ───────────────────────────────────────────
@app.route("/upload/status")
@login_required
def upload_status():
    with _jobs_lock:
        job = _jobs.get(current_user.nombre)
    if not job:
        return jsonify({"status": "idle"})
    return jsonify(job)

# ── DOWNLOAD EXCEL ───────────────────────────────────────────
@app.route("/download-excel")
@login_required
def download_excel():
    rutas = get_rutas_cliente(current_user.nombre)
    if os.path.exists(rutas["excel"]):
        return send_file(rutas["excel"], as_attachment=True)
    flash("Aún no tienes facturas procesadas", "warning")
    return redirect(url_for("dashboard"))

# ── LOGOUT ───────────────────────────────────────────────────
@app.route("/logout")
@login_required
def logout():
    logout_user()
    return redirect(url_for("index"))

if __name__ == "__main__":
    app.run(debug=False)
