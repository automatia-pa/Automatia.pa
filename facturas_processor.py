import urllib.request
import urllib.error
import json
import os
import urllib.parse
import sqlite3
import hashlib
import logging
import time
import shutil
import xml.etree.ElementTree as ET
from datetime import datetime
from logging.handlers import RotatingFileHandler
import re

import PyPDF2
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

# ─────────────────────────────────────────
# CONFIGURACION - SIN load_dotenv()
# ─────────────────────────────────────────
API_KEY          = os.environ.get("OPENROUTER_API_KEY")
TELEGRAM_TOKEN   = os.environ.get("TELEGRAM_TOKEN")
TELEGRAM_CHAT_ID = os.environ.get("TELEGRAM_CHAT_ID")

CARPETA_CLIENTES = os.path.expanduser("~/.private_data/clientes")

MODELOS = [
    "meta-llama/llama-3.3-70b-instruct:free",
    "nvidia/nemotron-3-super-120b-a12b:free",
    "nvidia/nemotron-3-nano-30b-a3b:free",
    "nvidia/nemotron-nano-12b-2-VL:free",
    "openai/gpt-oss-120b:free",
    "openai/gpt-oss-20b:free",
    "google/gemma-4-26b-a4b-it:free",
    "nousresearch/hermes-3-llama-3.1-405b:free",
    "openrouter/free",
]

CAMPOS_OBLIGATORIOS = ["proveedor", "monto_total", "moneda"]

# ── Logging con rotación (máx 2MB × 3 archivos) ─────────────
_handler = RotatingFileHandler(
    "procesamiento.log", maxBytes=2 * 1024 * 1024, backupCount=3, encoding="utf-8"
)
_handler.setFormatter(logging.Formatter("%(asctime)s | %(levelname)s | %(message)s"))
logging.getLogger().addHandler(_handler)
logging.getLogger().setLevel(logging.INFO)

# ─────────────────────────────────────────
def enviar_telegram(mensaje):
    if not TELEGRAM_TOKEN or not TELEGRAM_CHAT_ID:
        return
    try:
        url = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendMessage"
        params = urllib.parse.urlencode({
            "chat_id": TELEGRAM_CHAT_ID,
            "text": mensaje,
            "parse_mode": "HTML"
        }).encode("utf-8")
        req = urllib.request.Request(url + "?" + params.decode())
        urllib.request.urlopen(req, timeout=10)
    except Exception as e:
        logging.error(f"Telegram error: {e}")

# ─────────────────────────────────────────
def get_rutas_cliente(nombre_cliente):
    """Retorna rutas seguras para el cliente"""
    # Sanitizar nombre
    nombre_seguro = re.sub(r'[^\w\s\-.]', '', nombre_cliente).strip()
    if not nombre_seguro:
        raise ValueError("Nombre de cliente inválido")

    # Ruta base privada
    base_clientes = os.path.expanduser("~/.private_data/clientes")
    base = os.path.join(base_clientes, nombre_seguro)

    # Seguridad anti path traversal
    if not os.path.abspath(base).startswith(os.path.abspath(base_clientes)):
        raise ValueError("Ruta inválida - intento de path traversal")

    return {
        "base":       base,
        "facturas":   os.path.join(base, "facturas"),
        "procesados": os.path.join(base, "facturas", "procesados"),
        "error":      os.path.join(base, "facturas", "error"),
        "db":         os.path.join(base, "facturas.db"),
        "excel":      os.path.join(base, "resultados.xlsx"),
    }

def crear_carpetas_cliente(rutas):
    os.makedirs(rutas["facturas"],   exist_ok=True)
    os.makedirs(rutas["procesados"], exist_ok=True)
    os.makedirs(rutas["error"],      exist_ok=True)

def listar_clientes():
    if not os.path.exists(CARPETA_CLIENTES):
        os.makedirs(CARPETA_CLIENTES)
        return []
    return [
        f for f in os.listdir(CARPETA_CLIENTES)
        if os.path.isdir(os.path.join(CARPETA_CLIENTES, f))
    ]

# ─────────────────────────────────────────
CATEGORIAS_VALIDAS = [
    "Servicios", "Materiales y Suministros", "Transporte y Logistica",
    "Tecnologia y Software", "Nomina y RRHH", "Alquiler e Inmuebles",
    "Publicidad y Marketing", "Impuestos y Tasas", "Alimentacion", "Otros",
]

def init_db(db_path):
    with sqlite3.connect(db_path) as conn:
        conn.execute('''
            CREATE TABLE IF NOT EXISTS facturas (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                archivo TEXT UNIQUE,
                hash TEXT UNIQUE,
                proveedor TEXT,
                ruc TEXT,
                receptor TEXT,
                ruc_receptor TEXT,
                fecha TEXT,
                monto_total REAL,
                subtotal REAL,
                itbms REAL,
                moneda TEXT,
                descripcion TEXT,
                categoria TEXT DEFAULT "Otros",
                confianza INTEGER DEFAULT 70,
                modelo_usado TEXT,
                notas TEXT,
                cufe TEXT,
                tipo_doc TEXT DEFAULT "Factura",
                fuente TEXT DEFAULT "llm",
                estado TEXT DEFAULT "pendiente",
                comentario_estado TEXT,
                fecha_estado TEXT,
                fecha_procesamiento TEXT
            )
        ''')
        nuevas_columnas = [
            ('categoria',         'TEXT DEFAULT "Otros"'),
            ('receptor',          'TEXT'),
            ('ruc_receptor',      'TEXT'),
            ('subtotal',          'REAL'),
            ('itbms',             'REAL'),
            ('cufe',              'TEXT'),
            ('tipo_doc',          'TEXT DEFAULT "Factura"'),
            ('fuente',            'TEXT DEFAULT "llm"'),
            ('estado',            'TEXT DEFAULT "pendiente"'),
            ('comentario_estado', 'TEXT'),
            ('fecha_estado',      'TEXT'),
        ]
        for col, tipo in nuevas_columnas:
            try:
                conn.execute(f'ALTER TABLE facturas ADD COLUMN {col} {tipo}')
                conn.commit()
            except sqlite3.OperationalError:
                pass

# ─────────────────────────────────────────
# EXCEL con openpyxl puro (sin pandas)
# ─────────────────────────────────────────
COLUMNAS_EXCEL = [
    ("archivo",            "Archivo"),
    ("proveedor",          "Proveedor"),
    ("ruc",                "RUC"),
    ("receptor",           "Receptor"),
    ("ruc_receptor",       "RUC Receptor"),
    ("fecha",              "Fecha"),
    ("monto_total",        "Monto Total"),
    ("subtotal",           "Subtotal"),
    ("itbms",              "ITBMS"),
    ("moneda",             "Moneda"),
    ("categoria",          "Categoría"),
    ("tipo_doc",           "Tipo Documento"),
    ("estado",             "Estado"),
    ("comentario_estado",  "Comentario"),
    ("descripcion",        "Descripción"),
    ("cufe",               "CUFE"),
    ("fuente",             "Fuente"),
    ("confianza",          "Confianza %"),
    ("modelo_usado",       "Modelo"),
    ("fecha_procesamiento","Fecha Procesado"),
]

def exportar_excel(db_path, excel_path):
    try:
        with sqlite3.connect(db_path) as conn:
            conn.row_factory = sqlite3.Row
            cols_sql = ", ".join(c[0] for c in COLUMNAS_EXCEL)
            rows = conn.execute(
                f"SELECT {cols_sql} FROM facturas ORDER BY fecha_procesamiento DESC"
            ).fetchall()

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Facturas"

        # Cabecera con estilo
        header_fill = PatternFill("solid", fgColor="00E5C3")
        header_font = Font(bold=True, color="030F0D")
        for col_idx, (_, label) in enumerate(COLUMNAS_EXCEL, 1):
            cell = ws.cell(row=1, column=col_idx, value=label)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center")

        # Datos
        for row_idx, row in enumerate(rows, 2):
            for col_idx, (field, _) in enumerate(COLUMNAS_EXCEL, 1):
                ws.cell(row=row_idx, column=col_idx, value=row[field])

        # Ancho automático
        for col in ws.columns:
            max_len = max((len(str(c.value or "")) for c in col), default=10)
            ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 50)

        wb.save(excel_path)
        return len(rows)
    except Exception as e:
        logging.error(f"Error Excel: {e}")
        return 0


def exportar_dgi_csv(db_path, csv_path, periodo=None):
    """
    CSV formato DGI Panamá — solo facturas aprobadas.
    """
    try:
        with sqlite3.connect(db_path) as conn:
            query = """
                SELECT
                    ruc          AS RUC_Proveedor,
                    proveedor    AS Nombre_Proveedor,
                    tipo_doc     AS Tipo_Documento,
                    archivo      AS Numero_Documento,
                    cufe         AS CUFE,
                    fecha        AS Fecha_Emision,
                    COALESCE(subtotal, monto_total - COALESCE(itbms, 0)) AS Subtotal,
                    COALESCE(itbms, 0) AS ITBMS_7pct,
                    monto_total  AS Total_Factura,
                    moneda       AS Moneda,
                    categoria    AS Categoria
                FROM facturas
                WHERE estado = 'aprobada'
            """
            params = []
            if periodo:
                query += " AND fecha LIKE ?"
                params.append(f"{periodo}%")
            query += " ORDER BY fecha ASC"

            rows = conn.execute(query, params).fetchall()
            if not rows:
                return 0

            cols = ["RUC_Proveedor", "Nombre_Proveedor", "Tipo_Documento",
                    "Numero_Documento", "CUFE", "Fecha_Emision",
                    "Subtotal", "ITBMS_7pct", "Total_Factura", "Moneda", "Categoria"]

        import csv
        with open(csv_path, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f)
            writer.writerow(cols)
            for row in rows:
                # Formatear montos
                formatted = list(row)
                for i, col in enumerate(cols):
                    if col in ("Subtotal", "ITBMS_7pct", "Total_Factura"):
                        try:
                            formatted[i] = f"{float(row[i] or 0):.2f}"
                        except (TypeError, ValueError):
                            formatted[i] = "0.00"
                writer.writerow(formatted)

        return len(rows)
    except Exception as e:
        logging.error(f"Error exportando CSV DGI: {e}")
        return 0

# ─────────────────────────────────────────
def get_file_hash(ruta):
    with open(ruta, "rb") as f:
        return hashlib.md5(f.read()).hexdigest()

def mover_archivo(ruta_origen, carpeta_destino):
    nombre = os.path.basename(ruta_origen)
    for _ in range(6):
        try:
            shutil.move(ruta_origen, os.path.join(carpeta_destino, nombre))
            return True
        except PermissionError:
            time.sleep(1.5)
        except Exception:
            return False
    return False

# ─────────────────────────────────────────
def extraer_datos_xml_etax(ruta_archivo):
    try:
        tree = ET.parse(ruta_archivo)
        root = tree.getroot()
        ns = {
            'cbc': 'urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2',
            'cac': 'urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2',
            'ext': 'urn:oasis:names:specification:ubl:schema:xsd:CommonExtensionComponents-2',
        }

        def find_text(path, default=''):
            el = root.find(path, ns)
            return el.text.strip() if el is not None and el.text else default

        cufe     = find_text('.//cbc:UUID') or find_text('.//cbc:ID')
        proveedor = (find_text('.//cac:AccountingSupplierParty//cbc:RegistrationName') or
                     find_text('.//cac:AccountingSupplierParty//cbc:Name'))
        ruc      = (find_text('.//cac:AccountingSupplierParty//cbc:CompanyID') or
                    find_text('.//cac:AccountingSupplierParty//cbc:ID'))
        receptor = (find_text('.//cac:AccountingCustomerParty//cbc:RegistrationName') or
                    find_text('.//cac:AccountingCustomerParty//cbc:Name'))
        ruc_receptor = find_text('.//cac:AccountingCustomerParty//cbc:CompanyID')
        fecha    = find_text('.//cbc:IssueDate')

        monto_total_str = (find_text('.//cac:LegalMonetaryTotal//cbc:PayableAmount') or
                           find_text('.//cbc:TaxInclusiveAmount'))
        subtotal_str    = (find_text('.//cac:LegalMonetaryTotal//cbc:TaxExclusiveAmount') or
                           find_text('.//cac:LegalMonetaryTotal//cbc:LineExtensionAmount'))
        itbms_str       = find_text('.//cac:TaxTotal//cbc:TaxAmount')
        moneda          = find_text('.//cbc:DocumentCurrencyCode') or 'USD'

        tipo_doc_map = {'01': 'Factura', '02': 'Nota de Crédito', '03': 'Nota de Débito',
                        '04': 'Factura de Importación', '08': 'Factura de Exportación'}
        tipo_cod = find_text('.//cbc:InvoiceTypeCode')
        tipo_doc = tipo_doc_map.get(tipo_cod, f'Documento {tipo_cod}')

        descripcion = (find_text('.//cac:InvoiceLine//cbc:Description') or
                       find_text('.//cac:InvoiceLine//cbc:Name') or
                       f'{tipo_doc} electrónica')

        try:
            monto_total = float(monto_total_str)
        except (ValueError, TypeError):
            monto_total = 0.0
        try:
            itbms = float(itbms_str)
        except (ValueError, TypeError):
            itbms = 0.0
        try:
            subtotal = float(subtotal_str)
        except (ValueError, TypeError):
            subtotal = monto_total - itbms

        if not proveedor or monto_total == 0:
            logging.warning(f"XML e-Tax sin campos mínimos: {ruta_archivo}")
            return None

        return {
            'proveedor':    proveedor,
            'ruc':          ruc,
            'receptor':     receptor,
            'ruc_receptor': ruc_receptor,
            'fecha':        fecha,
            'monto_total':  monto_total,
            'subtotal':     subtotal,
            'itbms':        itbms,
            'moneda':       moneda,
            'descripcion':  descripcion,
            'cufe':         cufe,
            'tipo_doc':     tipo_doc,
            'confianza':    100,
            'categoria':    'Otros',
            'notas':        f'e-Tax 2.0 | CUFE: {cufe}',
            'fuente':       'xml_etax',
        }
    except ET.ParseError as e:
        logging.error(f"XML inválido {ruta_archivo}: {e}")
        return None
    except Exception as e:
        logging.error(f"Error parseando XML e-Tax {ruta_archivo}: {e}")
        return None


def extraer_texto(ruta_archivo):
    ext = os.path.splitext(ruta_archivo)[1].lower()
    try:
        if ext == ".txt":
            with open(ruta_archivo, "r", encoding="utf-8", errors="ignore") as f:
                return f.read()
        elif ext == ".pdf":
            texto = ""
            with open(ruta_archivo, "rb") as f:
                reader = PyPDF2.PdfReader(f)
                for i, page in enumerate(reader.pages):
                    texto += f"\n--- Pagina {i+1} ---\n" + (page.extract_text() or "")
            return texto
        elif ext in [".xlsx", ".xls"]:
            # Sin pandas: openpyxl directo
            wb = openpyxl.load_workbook(ruta_archivo, read_only=True, data_only=True)
            lineas = []
            for ws in wb.worksheets:
                for row in ws.iter_rows(values_only=True):
                    linea = "\t".join(str(c) if c is not None else "" for c in row)
                    if linea.strip():
                        lineas.append(linea)
            wb.close()
            return "\n".join(lineas)
    except Exception as e:
        logging.error(f"Error leyendo {ruta_archivo}: {e}")
    return None

# ─────────────────────────────────────────
def validar_resultado(datos):
    if not isinstance(datos, dict):
        return False
    for campo in CAMPOS_OBLIGATORIOS:
        if not datos.get(campo):
            logging.warning(f"Campo obligatorio faltante: {campo}")
            return False
    try:
        float(datos["monto_total"])
    except (ValueError, TypeError):
        logging.warning(f"monto_total no es número: {datos.get('monto_total')}")
        return False
    return True

def parsear_json_respuesta(content):
    if "<think>" in content:
        fin_think = content.find("</think>")
        if fin_think != -1:
            content = content[fin_think + len("</think>"):].strip()
    if "```" in content:
        partes = content.split("```")
        for parte in partes:
            parte = parte.replace("json", "").strip()
            if "{" in parte:
                content = parte
                break
    inicio = content.find("{")
    fin    = content.rfind("}") + 1
    if inicio == -1 or fin <= inicio:
        return None
    return json.loads(content[inicio:fin])

# ─────────────────────────────────────────
def llamar_ia(texto):
    texto_limpio = texto[:15000].replace('"', "'").replace('\\', '/')

    prompt = (
        "Eres un asistente contable experto en facturas de Panama y Latinoamerica.\n"
        "Analiza el siguiente texto de una factura y extrae los datos.\n\n"
        "INSTRUCCIONES IMPORTANTES:\n"
        "- Responde UNICAMENTE con un JSON valido, sin texto antes ni despues\n"
        "- No uses bloques de codigo ni comillas triples\n"
        "- Si no encuentras un campo, usa null\n"
        "- monto_total debe ser un numero, no texto\n"
        "- fecha debe estar en formato DD/MM/AAAA\n"
        "- moneda: usa USD, PAB, o la que corresponda\n"
        "- categoria debe ser UNA de estas opciones exactas: "
        "Servicios, Materiales y Suministros, Transporte y Logistica, "
        "Tecnologia y Software, Nomina y RRHH, Alquiler e Inmuebles, "
        "Publicidad y Marketing, Impuestos y Tasas, Alimentacion, Otros\n\n"
        "FORMATO EXACTO DE RESPUESTA:\n"
        '{"proveedor": "nombre empresa", "ruc": "numero ruc o null", '
        '"fecha": "DD/MM/AAAA o null", "monto_total": 0.00, '
        '"moneda": "USD", "categoria": "Servicios", '
        '"descripcion": "descripcion breve", '
        '"confianza": 85, "notas": "observaciones o null"}\n\n'
        "TEXTO DE LA FACTURA:\n"
        + texto_limpio
    )

    for i, model in enumerate(MODELOS):
        try:
            nombre = model.split('/')[1].split(':')[0]
            print(f"   Probando {nombre}...", end=" ", flush=True)

            if i > 0:
                time.sleep(0.5)

            data = {
                "model": model,
                "messages": [{"role": "user", "content": prompt}],
                "temperature": 0.0,
                "max_tokens": 1500
            }

            req = urllib.request.Request(
                url="https://openrouter.ai/api/v1/chat/completions",
                data=json.dumps(data).encode("utf-8"),
                headers={
                    "Content-Type": "application/json",
                    "Authorization": f"Bearer {API_KEY}",
                    "HTTP-Referer": "https://automatiapa.pythonanywhere.com",
                    "X-Title": "AutomatIA Facturas"
                },
                method="POST"
            )

            with urllib.request.urlopen(req, timeout=20) as resp:
                result = json.loads(resp.read().decode())

                if "choices" not in result or not result["choices"]:
                    print("respuesta vacia")
                    continue

                content = result["choices"][0]["message"]["content"].strip()
                if not content:
                    print("contenido vacio")
                    continue

                datos = parsear_json_respuesta(content)
                if datos is None:
                    print("JSON invalido")
                    logging.warning(f"{nombre}: no se pudo parsear JSON")
                    continue

                if not validar_resultado(datos):
                    print("campos incompletos")
                    logging.warning(f"{nombre}: campos obligatorios faltantes")
                    continue

                print(f"OK (confianza: {datos.get('confianza', '?')}%)")
                logging.info(f"Factura procesada con {nombre}")
                return datos, nombre

        except urllib.error.HTTPError as e:
            print(f"HTTP {e.code}")
            logging.warning(f"{nombre}: HTTP error {e.code}")
        except urllib.error.URLError as e:
            print("conexion fallida")
            logging.warning(f"{nombre}: URL error {e.reason}")
        except json.JSONDecodeError as e:
            print("JSON error")
            logging.warning(f"{nombre}: JSON decode error {e}")
        except Exception as e:
            print(f"error: {e}")
            logging.warning(f"{nombre}: error inesperado {e}")

    logging.error("Todos los modelos fallaron para esta factura")
    return None, None

# ─────────────────────────────────────────
def guardar_factura(db_path, datos, archivo, modelo):
    categoria = datos.get('categoria', 'Otros')
    if categoria not in CATEGORIAS_VALIDAS:
        categoria = 'Otros'

    try:
        with sqlite3.connect(db_path) as conn:
            conn.execute('''
                INSERT INTO facturas
                (archivo, hash, proveedor, ruc, receptor, ruc_receptor, fecha,
                 monto_total, subtotal, itbms, moneda, descripcion, categoria,
                 confianza, modelo_usado, notas, cufe, tipo_doc, fuente,
                 estado, fecha_procesamiento)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'pendiente', ?)
            ''', (
                archivo,
                datos.get('hash'),
                datos.get('proveedor'),
                datos.get('ruc'),
                datos.get('receptor'),
                datos.get('ruc_receptor'),
                datos.get('fecha'),
                float(datos.get('monto_total', 0)),
                float(datos.get('subtotal', 0)) if datos.get('subtotal') else None,
                float(datos.get('itbms', 0))    if datos.get('itbms')    else None,
                datos.get('moneda', 'USD'),
                datos.get('descripcion'),
                categoria,
                int(datos.get('confianza', 70)),
                modelo,
                datos.get('notas'),
                datos.get('cufe'),
                datos.get('tipo_doc', 'Factura'),
                datos.get('fuente', 'llm'),
                datetime.now().isoformat()
            ))
        return True
    except sqlite3.IntegrityError:
        logging.warning(f"Factura duplicada: {archivo}")
        return False
    except Exception as e:
        logging.error(f"Error guardando factura {archivo}: {e}")
        return False

# ─────────────────────────────────────────
def procesar_cliente(nombre_cliente, procesadas):
    rutas = get_rutas_cliente(nombre_cliente)
    crear_carpetas_cliente(rutas)
    init_db(rutas["db"])

    archivos = [
        f for f in os.listdir(rutas["facturas"])
        if f.lower().endswith(('.pdf', '.txt', '.xlsx', '.xls', '.xml'))
    ]

    clave = lambda f: f"{nombre_cliente}/{f}"

    for archivo in archivos:
        if clave(archivo) in procesadas:
            continue

        ruta = os.path.join(rutas["facturas"], archivo)
        ext  = os.path.splitext(archivo)[1].lower()
        print(f"\n  [{nombre_cliente}] Procesando: {archivo}")
        logging.info(f"Iniciando: {nombre_cliente}/{archivo}")

        file_hash = get_file_hash(ruta)

        with sqlite3.connect(rutas["db"]) as conn:
            existe = conn.execute(
                "SELECT 1 FROM facturas WHERE hash=?", (file_hash,)
            ).fetchone()

        if existe:
            print("   Duplicada, ignorando")
            mover_archivo(ruta, rutas["procesados"])
            procesadas.add(clave(archivo))
            continue

        resultado = None
        modelo    = None

        if ext == ".xml":
            resultado = extraer_datos_xml_etax(ruta)
            if resultado:
                modelo = "e-Tax-XML"
                resultado['hash'] = file_hash
            else:
                print("   XML inválido o sin campos mínimos")
                logging.error(f"XML e-Tax inválido: {archivo}")
                mover_archivo(ruta, rutas["error"])
                procesadas.add(clave(archivo))
                continue
        else:
            texto = extraer_texto(ruta)
            if not texto or len(texto.strip()) < 20:
                print("   No se pudo extraer texto util")
                logging.error(f"Texto insuficiente: {archivo}")
                mover_archivo(ruta, rutas["error"])
                procesadas.add(clave(archivo))
                continue

            resultado, modelo = llamar_ia(texto)
            if resultado:
                resultado['hash'] = file_hash

        if resultado:
            if guardar_factura(rutas["db"], resultado, archivo, modelo):
                fuente_tag = "🧾 e-Tax XML" if ext == ".xml" else f"🤖 {modelo}"
                print(f"   OK: {resultado.get('proveedor')} | {resultado.get('monto_total')} {resultado.get('moneda')} | {fuente_tag}")
                logging.info(f"OK: {archivo} | {resultado.get('proveedor')} | {resultado.get('monto_total')}")
                enviar_telegram(
                    f"{'🧾' if ext == '.xml' else '📄'} Factura procesada\n"
                    f"Cliente: {nombre_cliente}\n"
                    f"Archivo: {archivo}\n"
                    f"Fuente: {fuente_tag}\n"
                    f"Proveedor: {resultado.get('proveedor')}\n"
                    f"Categoría: {resultado.get('categoria', 'Otros')}\n"
                    f"Monto: {resultado.get('monto_total')} {resultado.get('moneda')}\n"
                    f"Confianza: {resultado.get('confianza')}%"
                )
                mover_archivo(ruta, rutas["procesados"])
            else:
                print("   Ya estaba guardada (duplicado por hash)")
                mover_archivo(ruta, rutas["procesados"])
        else:
            print("   Fallo el procesamiento con todos los modelos")
            logging.error(f"FALLO TOTAL: {nombre_cliente}/{archivo}")
            enviar_telegram(
                f"❌ Error procesando factura\n"
                f"Cliente: {nombre_cliente}\n"
                f"Archivo: {archivo}\n"
                f"Todos los modelos fallaron"
            )
            mover_archivo(ruta, rutas["error"])

        procesadas.add(clave(archivo))

    # Exportar Excel UNA SOLA VEZ al terminar el lote (no dentro del loop)
    exportar_excel(rutas["db"], rutas["excel"])

# ─────────────────────────────────────────
def main():
    os.makedirs(CARPETA_CLIENTES, exist_ok=True)
    print("=" * 60)
    print("  SISTEMA MULTI-CLIENTE DE FACTURAS")
    print("=" * 60)
    enviar_telegram("Sistema multi-cliente iniciado")
    procesadas = set()
    while True:
        clientes = listar_clientes()
        if not clientes:
            print(" Sin clientes aun. Crea una carpeta en ./clientes/", end="\r")
        else:
            for cliente in clientes:
                procesar_cliente(cliente, procesadas)
        time.sleep(15)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\nSistema detenido.")
    except Exception as e:
        logging.error(f"Error critico: {e}")
        print(f"Error critico: {e}")
