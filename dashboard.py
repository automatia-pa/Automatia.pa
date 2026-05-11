import streamlit as st
import sqlite3
import pandas as pd
import os
import io

st.set_page_config(page_title="Facturas IA", layout="wide")
st.title("Dashboard - Sistema de Facturas")

CARPETA_CLIENTES = "./clientes"

def cargar_datos():
    todos = []
    if not os.path.exists(CARPETA_CLIENTES):
        return pd.DataFrame()
    
    for cliente in os.listdir(CARPETA_CLIENTES):
        db_path = os.path.join(CARPETA_CLIENTES, cliente, "facturas.db")
        if not os.path.exists(db_path):
            continue
        try:
            conn = sqlite3.connect(db_path)
            df = pd.read_sql_query("""
                SELECT archivo, proveedor, ruc, fecha, monto_total, 
                       moneda, descripcion, confianza, fecha_procesamiento
                FROM facturas
            """, conn)
            conn.close()
            df["cliente"] = cliente
            todos.append(df)
        except:
            continue
    
    if todos:
        return pd.concat(todos, ignore_index=True)
    return pd.DataFrame()

df = cargar_datos()

if df.empty:
    st.warning("No hay facturas procesadas aun.")
    st.stop()

# ── FILTROS ──────────────────────────────
st.sidebar.title("Filtros")

clientes = ["Todos"] + sorted(df["cliente"].unique().tolist())
cliente_sel = st.sidebar.selectbox("Cliente", clientes)

if cliente_sel != "Todos":
    df = df[df["cliente"] == cliente_sel]

proveedores = ["Todos"] + sorted(df["proveedor"].dropna().unique().tolist())
proveedor_sel = st.sidebar.selectbox("Proveedor", proveedores)

if proveedor_sel != "Todos":
    df = df[df["proveedor"] == proveedor_sel]

# ── METRICAS ─────────────────────────────
st.subheader("Resumen")
col1, col2, col3, col4 = st.columns(4)
col1.metric("Total Facturas", len(df))
col2.metric("Monto Total", f"${df['monto_total'].sum():,.2f}")
col3.metric("Confianza Promedio", f"{df['confianza'].mean():.1f}%")
col4.metric("Proveedores", df["proveedor"].nunique())

# ── TABLA ────────────────────────────────
st.subheader("Facturas")
st.dataframe(df.sort_values("fecha_procesamiento", ascending=False), width="stretch")

# ── DESCARGA EXCEL ───────────────────────
buffer = io.BytesIO()
df.sort_values("fecha_procesamiento", ascending=False).to_excel(buffer, index=False)
buffer.seek(0)

st.download_button(
    label="Descargar Excel",
    data=buffer,
    file_name=f"facturas_{cliente_sel}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

# ── GRAFICO ──────────────────────────────
st.subheader("Monto por Proveedor")
chart = df.groupby("proveedor")["monto_total"].sum().sort_values(ascending=False)
st.bar_chart(chart)

# ── GRAFICO POR CLIENTE ───────────────────
if cliente_sel == "Todos":
    st.subheader("Monto por Cliente")
    chart2 = df.groupby("cliente")["monto_total"].sum().sort_values(ascending=False)
    st.bar_chart(chart2)