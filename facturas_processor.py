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

MODELOS = [
    "nvidia/nemotron-3-super-120b-a12b:free",
    "qwen/qwen3-coder:free",
    "qwen/qwen3-32b:free",
    "meta-llama/llama-3.3-70b-instruct:free",
]

logging.basicConfig(
    filename="procesamiento.log",
    level=logging.INFO,
    format='%(asctime)s | %(levelname)s | %(message)s',
    encoding='utf-8'
)

# ─────────────────────────────────────────
def enviar_telegram(mensaje):
    try:
        url = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendMessage"
        params = urllib.parse.urlencode({
            "chat_id": TELEGRAM_CHAT_ID,
            "text": mensaje,
            "parse_mode": "HTML"
        }).encode("utf-8")
        req = urllib.request.Request(url + "?" + params.decode())
        urllib.request.urlopen(req)
    except Exception as e:
        logging.error(f"Telegram error: {e}")
        print(f" Telegram error: {e}")

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
    """Crea las carpetas de un cliente si no existen"""
    os.makedirs(rutas["facturas"],   exist_ok=True)
    os.makedirs(rutas["procesados"], exist_ok=True)
    os.makedirs(rutas["error"],      exist_ok=True)

def listar_clientes():
    """Devuelve la lista de clientes que existen en la carpeta"""
    if not os.path.exists(CARPETA_CLIENTES):
        os.makedirs(CARPETA_CLIENTES)
        return []
    return [
        f for f in os.listdir(CARPETA_CLIENTES)
        if os.path.isdir(os.path.join(CARPETA_CLIENTES, f))
    ]

# ─────────────────────────────────────────
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
            confianza INTEGER DEFAULT 70,
            modelo_usado TEXT,
            notas TEXT,
            fecha_procesamiento TEXT
        )
    ''')
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
                descripcion as Descripcion,
                confianza as Confianza,
                modelo_usado as Modelo,
                fecha_procesamiento as Fecha_Procesado
            FROM facturas 
            ORDER BY fecha_procesamiento DESC
        """, conn)
        conn.close()

        # Escribir directo sin archivo temporal
        df.to_excel(excel_path, index=False, engine='openpyxl')

        return len(df)
    except Exception as e:
        print(f" Error Excel: {e}")
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
def llamar_ia(texto):
    texto_limpio = texto[:17000].replace('"', "'").replace('\\', '/')
    prompt = f"""
Eres un experto contable. Extrae los datos de esta factura.
Responde UNICAMENTE con un JSON valido, sin texto adicional.

{{
    "proveedor": "...",
    "ruc": "...",
    "fecha": "DD/MM/AAAA",
    "monto_total": numero,
    "moneda": "...",
    "descripcion": "...",
    "confianza": numero entre 0 y 100,
    "notas": "observaciones opcionales"
}}

TEXTO DE LA FACTURA:
{texto_limpio}
"""
    for model in MODELOS:
        try:
            nombre = model.split('/')[1].split(':')[0]
            print(f"   Probando {nombre}...", end=" ")
            data = {
                "model": model,
                "messages": [{"role": "user", "content": prompt}],
                "temperature": 0.0,
                "max_tokens": 900
            }
            req = urllib.request.Request(
                url="https://openrouter.ai/api/v1/chat/completions",
                data=json.dumps(data).encode("utf-8"),
                headers={
                    "Content-Type": "application/json",
                    "Authorization": f"Bearer {API_KEY}"
                },
                method="POST"
            )
            with urllib.request.urlopen(req, timeout=45) as resp:
                result = json.loads(resp.read().decode())
                content = result["choices"][0]["message"]["content"].strip()
                if "```" in content:
                    content = content.split("```")[1].replace("json", "").strip()
                inicio = content.find("{")
                fin = content.rfind("}") + 1
                if inicio != -1 and fin > inicio:
                    content = content[inicio:fin]
                return json.loads(content), nombre
        except Exception as e:
            print(f"fallo: {e}")
            continue
    return None, None

# ─────────────────────────────────────────
def guardar_factura(db_path, datos, archivo, modelo):
    conn = sqlite3.connect(db_path)
    try:
        conn.execute('''
            INSERT INTO facturas 
            (archivo, hash, proveedor, ruc, fecha, monto_total, moneda,
             descripcion, confianza, modelo_usado, notas, fecha_procesamiento)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (
            archivo, datos.get('hash'), datos.get('proveedor'), datos.get('ruc'),
            datos.get('fecha'), datos.get('monto_total'), datos.get('moneda'),
            datos.get('descripcion'), datos.get('confianza', 70), modelo,
            datos.get('notas'), datetime.now().isoformat()
        ))
        conn.commit()
        return True
    except sqlite3.IntegrityError:
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
        if not texto:
            print("   No se pudo extraer texto")
            mover_archivo(ruta, rutas["error"])
            procesadas.add(clave(archivo))
            continue

        resultado, modelo = llamar_ia(texto)

        if resultado:
            resultado['hash'] = file_hash
            if guardar_factura(rutas["db"], resultado, archivo, modelo):
                total = exportar_excel(rutas["db"], rutas["excel"])
                print(f"   OK: {resultado.get('proveedor')} | {resultado.get('monto_total')} {resultado.get('moneda')} | Confianza: {resultado.get('confianza')}%")
                enviar_telegram(
                    f"Factura procesada\n"
                    f"Cliente: {nombre_cliente}\n"
                    f"Archivo: {archivo}\n"
                    f"Proveedor: {resultado.get('proveedor')}\n"
                    f"Monto: {resultado.get('monto_total')} {resultado.get('moneda')}\n"
                    f"Total en DB: {total} facturas"
                )
                mover_archivo(ruta, rutas["procesados"])
        else:
            print("   Fallo el procesamiento")
            enviar_telegram(f"Error procesando\nCliente: {nombre_cliente}\nArchivo: {archivo}")
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
        print(f"Error critico: {e}")