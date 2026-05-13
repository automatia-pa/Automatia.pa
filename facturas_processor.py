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
from datetime import datetime

import PyPDF2
import pandas as pd
from dotenv import load_dotenv

load_dotenv()

# ─────────────────────────────────────────
# CONFIGURACION
# ─────────────────────────────────────────
API_KEY          = os.environ.get("OPENROUTER_API_KEY")
TELEGRAM_TOKEN   = os.environ.get("TELEGRAM_TOKEN")
TELEGRAM_CHAT_ID = os.environ.get("TELEGRAM_CHAT_ID")

CARPETA_CLIENTES = "./clientes"

# Modelos ordenados del más confiable al menos confiable
MODELOS = [
    "meta-llama/llama-3.3-70b-instruct:free",
    "nvidia/nemotron-3-super-120b-a12b:free",
    "openai/gpt-oss-20b:free",
    "google/gemma-4-26b-a4b-it:free",
    "nousresearch/hermes-3-llama-3.1-405b:free",
    "openrouter/free",   # fallback automático al mejor disponible,
]

# Campos obligatorios que debe tener el JSON para considerarse válido
CAMPOS_OBLIGATORIOS = ["proveedor", "monto_total", "moneda"]

logging.basicConfig(
    filename="procesamiento.log",
    level=logging.INFO,
    format='%(asctime)s | %(levelname)s | %(message)s',
    encoding='utf-8'
)

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
    base = os.path.join(os.path.dirname(os.path.abspath(__file__)), "clientes", nombre_cliente)
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
    "Servicios",
    "Materiales y Suministros",
    "Transporte y Logistica",
    "Tecnologia y Software",
    "Nomina y RRHH",
    "Alquiler e Inmuebles",
    "Publicidad y Marketing",
    "Impuestos y Tasas",
    "Alimentacion",
    "Otros",
]

def init_db(db_path):
    conn = sqlite3.connect(db_path)
    conn.execute('''
        CREATE TABLE IF NOT EXISTS facturas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            archivo TEXT UNIQUE,
            hash TEXT UNIQUE,
            proveedor TEXT,
            ruc TEXT,
            fecha TEXT,
            monto_total REAL,
            moneda TEXT,
            descripcion TEXT,
            categoria TEXT DEFAULT "Otros",
            confianza INTEGER DEFAULT 70,
            modelo_usado TEXT,
            notas TEXT,
            fecha_procesamiento TEXT
        )
    ''')
    # Agregar columna categoria si no existe (para bases de datos ya creadas)
    try:
        conn.execute('ALTER TABLE facturas ADD COLUMN categoria TEXT DEFAULT "Otros"')
        conn.commit()
    except sqlite3.OperationalError:
        pass  # La columna ya existe
    conn.commit()
    conn.close()

# ─────────────────────────────────────────
def exportar_excel(db_path, excel_path):
    try:
        conn = sqlite3.connect(db_path)
        df = pd.read_sql_query("""
            SELECT 
                archivo as Archivo,
                proveedor as Proveedor,
                ruc as RUC,
                fecha as Fecha,
                monto_total as Monto_Total,
                moneda as Moneda,
                categoria as Categoria,
                descripcion as Descripcion,
                confianza as Confianza,
                modelo_usado as Modelo,
                fecha_procesamiento as Fecha_Procesado
            FROM facturas 
            ORDER BY fecha_procesamiento DESC
        """, conn)
        conn.close()
        df.to_excel(excel_path, index=False, engine='openpyxl')
        return len(df)
    except Exception as e:
        logging.error(f"Error Excel: {e}")
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
        except:
            return False
    return False

# ─────────────────────────────────────────
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
            df = pd.read_excel(ruta_archivo)
            return df.to_string(index=False)
    except Exception as e:
        logging.error(f"Error leyendo {ruta_archivo}: {e}")
    return None

# ─────────────────────────────────────────
def validar_resultado(datos):
    """Verifica que el JSON tenga los campos mínimos necesarios"""
    if not isinstance(datos, dict):
        return False
    for campo in CAMPOS_OBLIGATORIOS:
        if not datos.get(campo):
            logging.warning(f"Campo obligatorio faltante o vacio: {campo}")
            return False
    # Verificar que monto_total sea un número válido
    try:
        float(datos["monto_total"])
    except (ValueError, TypeError):
        logging.warning(f"monto_total no es un numero valido: {datos.get('monto_total')}")
        return False
    return True

def parsear_json_respuesta(content):
    # Eliminar bloques <think>...</think> de Qwen3
    if "<think>" in content:
        fin_think = content.find("</think>")
        if fin_think != -1:
            content = content[fin_think + len("</think>"):].strip()

    # Limpiar bloques de código markdown
    if "```" in content:
        partes = content.split("```")
        for parte in partes:
            parte = parte.replace("json", "").strip()
            if "{" in parte:
                content = parte
                break

    inicio = content.find("{")
    fin = content.rfind("}") + 1
    if inicio == -1 or fin <= inicio:
        return None

    content = content[inicio:fin]
    return json.loads(content)

# ─────────────────────────────────────────
def llamar_ia(texto):
    texto_limpio = texto[:15000].replace('"', "'")
    texto_limpio = texto_limpio.replace('\\', '/')

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

            # Pausa entre modelos para evitar rate limiting
            if i > 0:
                time.sleep(2)

            data = {
                "model": model,
                "messages": [{"role": "user", "content": prompt}],
                "temperature": 0.0,
                "max_tokens": 800
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

            with urllib.request.urlopen(req, timeout=45) as resp:
                result = json.loads(resp.read().decode())

                # Verificar que la respuesta tiene el formato esperado
                if "choices" not in result or not result["choices"]:
                    print(f"respuesta vacia")
                    logging.warning(f"{nombre}: respuesta sin choices")
                    continue

                content = result["choices"][0]["message"]["content"].strip()

                if not content:
                    print("contenido vacio")
                    continue

                datos = parsear_json_respuesta(content)

                if datos is None:
                    print("JSON invalido")
                    logging.warning(f"{nombre}: no se pudo parsear JSON. Respuesta: {content[:200]}")
                    continue

                if not validar_resultado(datos):
                    print("campos incompletos")
                    logging.warning(f"{nombre}: campos obligatorios faltantes. Datos: {datos}")
                    continue

                print(f"OK (confianza: {datos.get('confianza', '?')}%)")
                logging.info(f"Factura procesada con {nombre}")
                return datos, nombre

        except urllib.error.HTTPError as e:
            print(f"HTTP {e.code}")
            logging.warning(f"{nombre}: HTTP error {e.code}")
            continue
        except urllib.error.URLError as e:
            print(f"conexion fallida")
            logging.warning(f"{nombre}: URL error {e.reason}")
            continue
        except json.JSONDecodeError as e:
            print(f"JSON error")
            logging.warning(f"{nombre}: JSON decode error {e}")
            continue
        except Exception as e:
            print(f"error: {e}")
            logging.warning(f"{nombre}: error inesperado {e}")
            continue

    logging.error("Todos los modelos fallaron para esta factura")
    return None, None

# ─────────────────────────────────────────
def guardar_factura(db_path, datos, archivo, modelo):
    # Validar que la categoria sea una de las permitidas
    categoria = datos.get('categoria', 'Otros')
    if categoria not in CATEGORIAS_VALIDAS:
        categoria = 'Otros'

    conn = sqlite3.connect(db_path)
    try:
        conn.execute('''
            INSERT INTO facturas 
            (archivo, hash, proveedor, ruc, fecha, monto_total, moneda,
             descripcion, categoria, confianza, modelo_usado, notas, fecha_procesamiento)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (
            archivo,
            datos.get('hash'),
            datos.get('proveedor'),
            datos.get('ruc'),
            datos.get('fecha'),
            float(datos.get('monto_total', 0)),
            datos.get('moneda', 'USD'),
            datos.get('descripcion'),
            categoria,
            int(datos.get('confianza', 70)),
            modelo,
            datos.get('notas'),
            datetime.now().isoformat()
        ))
        conn.commit()
        return True
    except sqlite3.IntegrityError:
        logging.warning(f"Factura duplicada: {archivo}")
        return False
    except Exception as e:
        logging.error(f"Error guardando factura {archivo}: {e}")
        return False
    finally:
        conn.close()

# ─────────────────────────────────────────
def procesar_cliente(nombre_cliente, procesadas):
    rutas = get_rutas_cliente(nombre_cliente)
    crear_carpetas_cliente(rutas)
    init_db(rutas["db"])

    archivos = [
        f for f in os.listdir(rutas["facturas"])
        if f.lower().endswith(('.pdf', '.txt', '.xlsx', '.xls'))
    ]

    clave = lambda f: f"{nombre_cliente}/{f}"

    for archivo in archivos:
        if clave(archivo) in procesadas:
            continue

        ruta = os.path.join(rutas["facturas"], archivo)
        print(f"\n  [{nombre_cliente}] Procesando: {archivo}")
        logging.info(f"Iniciando: {nombre_cliente}/{archivo}")

        file_hash = get_file_hash(ruta)

        conn = sqlite3.connect(rutas["db"])
        existe = conn.execute(
            "SELECT 1 FROM facturas WHERE hash=?", (file_hash,)
        ).fetchone()
        conn.close()

        if existe:
            print("   Duplicada, ignorando")
            mover_archivo(ruta, rutas["procesados"])
            procesadas.add(clave(archivo))
            continue

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
            if guardar_factura(rutas["db"], resultado, archivo, modelo):
                total = exportar_excel(rutas["db"], rutas["excel"])
                print(f"   OK: {resultado.get('proveedor')} | {resultado.get('monto_total')} {resultado.get('moneda')} | Confianza: {resultado.get('confianza')}%")
                logging.info(f"OK: {archivo} | {resultado.get('proveedor')} | {resultado.get('monto_total')} {resultado.get('moneda')}")
                enviar_telegram(
                    f"Factura procesada\n"
                    f"Cliente: {nombre_cliente}\n"
                    f"Archivo: {archivo}\n"
                    f"Proveedor: {resultado.get('proveedor')}\n"
                    f"Categoria: {resultado.get('categoria', 'Otros')}\n"
                    f"Monto: {resultado.get('monto_total')} {resultado.get('moneda')}\n"
                    f"Confianza: {resultado.get('confianza')}%\n"
                    f"Total en DB: {total} facturas"
                )
                mover_archivo(ruta, rutas["procesados"])
            else:
                print("   Ya estaba guardada (duplicado por hash)")
                mover_archivo(ruta, rutas["procesados"])
        else:
            print("   Fallo el procesamiento con todos los modelos")
            logging.error(f"FALLO TOTAL: {nombre_cliente}/{archivo}")
            enviar_telegram(
                f"Error procesando factura\n"
                f"Cliente: {nombre_cliente}\n"
                f"Archivo: {archivo}\n"
                f"Todos los modelos fallaron"
            )
            mover_archivo(ruta, rutas["error"])

        procesadas.add(clave(archivo))

# ─────────────────────────────────────────
def main():
    os.makedirs(CARPETA_CLIENTES, exist_ok=True)

    print("=" * 60)
    print("  SISTEMA MULTI-CLIENTE DE FACTURAS")
    print("=" * 60)
    print(f"  Clientes en: {CARPETA_CLIENTES}")
    print("  Para agregar un cliente, crea una carpeta con su nombre")
    print("  (Ctrl+C para detener)")
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
