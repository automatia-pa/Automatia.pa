from flask import Flask, render_template, request, redirect, url_for, flash, send_file
from flask_login import LoginManager, login_user, logout_user, login_required, current_user
from werkzeug.utils import secure_filename
import os
import pandas as pd
from dotenv import load_dotenv

from models import User, db_init
from facturas_processor import procesar_cliente, get_rutas_cliente

# ── Cargar variables de entorno ──────────────────────────────
load_dotenv()

app = Flask(__name__)

# FIX 1: secret_key viene del .env, nunca hardcodeado
app.secret_key = os.environ.get("FLASK_SECRET_KEY")
if not app.secret_key:
    raise RuntimeError("FLASK_SECRET_KEY no está configurado en .env")

# FIX 2: limite de tamaño de archivos — 10 MB máximo
app.config['MAX_CONTENT_LENGTH'] = 10 * 1024 * 1024

# FIX 3: extensiones permitidas
ALLOWED_EXTENSIONS = {'pdf', 'txt', 'xlsx', 'xls'}

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

# FIX 4: ADMIN_SECRET viene del .env
ADMIN_SECRET = os.environ.get("ADMIN_SECRET")

# ── INDEX + LOGIN ────────────────────────────────────────────
@app.route("/", methods=["GET", "POST"])
def index():
    if current_user.is_authenticated:
        return redirect(url_for("dashboard"))
    if request.method == "POST":
        email    = request.form.get("email")
        password = request.form.get("password")
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
        nombre   = request.form.get("nombre")
        email    = request.form.get("email")
        password = request.form.get("password")
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
        email    = request.form.get("email")
        password = request.form.get("password")
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
    rutas = get_rutas_cliente(current_user.nombre)
    facturas = []
    if os.path.exists(rutas["excel"]):
        try:
            df = pd.read_excel(rutas["excel"])
            facturas = df.to_dict("records")
        except Exception:
            pass
    return render_template("dashboard.html", user=current_user, facturas=facturas)

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

    procesados = 0
    rechazados = 0
    for file in files:
        if file.filename == "":
            continue
        # FIX 5: validar extensión y sanitizar nombre
        if not allowed_file(file.filename):
            rechazados += 1
            continue
        filename = secure_filename(file.filename)
        filepath = os.path.join(rutas["facturas"], filename)
        file.save(filepath)
        procesados += 1

    if rechazados > 0:
        flash(f"{rechazados} archivo(s) rechazados (solo PDF, TXT, XLSX)", "warning")

    if procesados > 0:
        procesar_cliente(current_user.nombre, set())
        flash(f"{procesados} factura(s) procesada(s) correctamente", "success")

    return redirect(url_for("dashboard"))

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

# FIX 6: debug=False — NUNCA True en producción
if __name__ == "__main__":
    app.run(debug=False)