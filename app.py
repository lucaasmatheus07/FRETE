import io
import sqlite3
import openpyxl
import pandas as pd
import streamlit as st
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# Configuração da página
st.set_page_config(
    page_title="Controle de Registro de Frete - Ellenco",
    page_icon="🚚",
    layout="wide"
)

DB_NAME = "fretes_ellenco.db"

# --- FUNÇÕES DE BANCO DE DADOS (PERSISTÊNCIA) ---
def init_db():
    with sqlite3.connect(DB_NAME) as conn:
        cursor = conn.cursor()
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS fretes (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                data TEXT,
                fretista TEXT,
                placa TEXT,
                obra TEXT,
                fornecedor TEXT,
                nf TEXT,
                num_vale TEXT,
                qtd REAL,
                km REAL,
                km_total REAL,
                pedagio REAL,
                valor_frete REAL,
                tipo_desconto TEXT,
                desconto REAL,
                adicional_diesel REAL,
                total_geral REAL
            )
        """)
        # Migração simples caso a tabela já exista sem a coluna adicional_diesel
        try:
            cursor.execute("ALTER TABLE fretes ADD COLUMN adicional_diesel REAL DEFAULT 0.0")
        except sqlite3.OperationalError:
            pass
        conn.commit()

def carregar_fretes():
    with sqlite3.connect(DB_NAME) as conn:
        query = """
            SELECT 
                data as Data, 
                fretista as Fretista, 
                placa as Placa, 
                obra as Obra, 
                fornecedor as Fornecedor, 
                nf as NF, 
                num_vale as 'Nº Vale', 
                qtd as 'Qtd (Ton)', 
                km as KM, 
                km_total as 'KM Total', 
                pedagio as 'Pedágio (R$)', 
                valor_frete as 'Valor Frete (R$)', 
                tipo_desconto as 'Tipo Desconto', 
                desconto as 'Desconto (R$)', 
                adicional_diesel as 'Adicional Diesel (R$)',
                total_geral as 'Total Geral (R$)' 
            FROM fretes
        """
        df = pd.read_sql_query(query, conn)
    return df.to_dict(orient="records")

def salvar_novo_frete(item):
    with sqlite3.connect(DB_NAME) as conn:
        cursor = conn.cursor()
        cursor.execute("""
            INSERT INTO fretes (data, fretista, placa, obra, fornecedor, nf, num_vale, qtd, km, km_total, pedagio, valor_frete, tipo_desconto, desconto, adicional_diesel, total_geral)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            item["Data"], item["Fretista"], item["Placa"], item["Obra"], item["Fornecedor"],
            item["NF"], item["Nº Vale"], item["Qtd (Ton)"], item["KM"], item["KM Total"],
            item["Pedágio (R$)"], item["Valor Frete (R$)"], item["Tipo Desconto"],
            item["Desconto (R$)"], item["Adicional Diesel (R$)"], item["Total Geral (R$)"]
        ))
        conn.commit()

def atualizar_banco_inteiro(df_edited):
    with sqlite3.connect(DB_NAME) as conn:
        cursor = conn.cursor()
        cursor.execute("DELETE FROM fretes")
        for _, row in df_edited.iterrows():
            cursor.execute("""
                INSERT INTO fretes (data, fretista, placa, obra, fornecedor, nf, num_vale, qtd, km, km_total, pedagio, valor_frete, tipo_desconto, desconto, adicional_diesel, total_geral)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                str(row.get("Data", "")),
                str(row.get("Fretista", "")),
                str(row.get("Placa", "")),
                str(row.get("Obra", "")),
                str(row.get("Fornecedor", "")),
                str(row.get("NF", "")),
                str(row.get("Nº Vale", "")),
                float(row.get("Qtd (Ton)", 0) or 0),
                float(row.get("KM", 0) or 0),
                float(row.get("KM Total", 0) or 0),
                float(row.get("Pedágio (R$)", 0) or 0),
                float(row.get("Valor Frete (R$)", 0) or 0),
                str(row.get("Tipo Desconto", "Nenhum")),
                float(row.get("Desconto (R$)", 0) or 0),
                float(row.get("Adicional Diesel (R$)", 0) or 0),
                float(row.get("Total Geral (R$)", 0) or 0)
            ))
        conn.commit()

def limpar_banco():
    with sqlite3.connect(DB_NAME) as conn:
        cursor = conn.cursor()
        cursor.execute("DELETE FROM fretes")
        conn.commit()

# Inicializa o BD
init_db()

# --- ESTILIZAÇÃO CSS CUSTOMIZADA ---
st.markdown("""
    <style>
        .stApp {
            background-color: #f4f6f9;
        }
        .main-header {
            background: linear-gradient(135deg, #00234B 0%, #001530 100%);
            padding: 20px 24px;
            border-radius: 10px;
            color: white;
            border-bottom: 5px solid #D9232A;
            margin-bottom: 25px;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
        }
        .main-header h1 {
            color: #FFFFFF !important;
            font-weight: 800;
            margin: 0;
            padding: 0;
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            font-size: 28px;
        }
        .main-header p {
            color: #E0E6ED;
            margin-top: 5px;
            margin-bottom: 0;
        }
        [data-testid="stForm"] {
            background-color: #FFFFFF;
            border-radius: 10px;
            border: 1px solid #E2E8F0;
            box-shadow: 0 4px 12px rgba(0,0,0,0.05);
            padding: 20px;
        }
        h5 {
            color: #00234B !important;
            border-left: 4px solid #D9232A;
            padding-left: 8px;
            margin-bottom: 12px !important;
            font-weight: 700 !important;
        }
        div.stButton > button, div[data-testid="stFormSubmitButton"] > button {
            background-color: #FFFFFF !important;
            color: #00234B !important;
            border: 2px solid #00234B !important;
            font-weight: bold !important;
            font-size: 16px !important;
            border-radius: 6px !important;
            padding: 12
