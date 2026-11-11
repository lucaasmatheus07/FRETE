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
        try:
            cursor.execute("ALTER TABLE fretes ADD COLUMN adicional_diesel REAL DEFAULT 0.0")
        except sqlite3.OperationalError:
            pass
        conn.commit()

def carregar_fretes():
    with sqlite3.connect(DB_NAME) as conn:
        query = """
            SELECT 
                fretista as Fretista, 
                data as Data, 
                num_vale as 'Nº Vale', 
                nf as 'Nota Fiscal', 
                km as KM, 
                obra as Obra, 
                qtd as 'Qtd (Ton)', 
                placa as Placa, 
                fornecedor as Fornecedor, 
                valor_frete as 'Valor Frete (R$)', 
                tipo_desconto as 'Tipo Desconto', 
                desconto as 'Desconto (R$)', 
                pedagio as 'Pedágio (R$)', 
                adicional_diesel as 'Adicional Diesel (R$)',
                km_total as 'KM Total',
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
            item["Nota Fiscal"], item["Nº Vale"], item["Qtd (Ton)"], item["KM"], item["KM Total"],
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
                str(row.get("Nota Fiscal", "")),
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
            padding: 12px 24px !important;
            transition: all 0.3s ease !important;
            box-shadow: 0 2px 4px rgba(0, 0, 0, 0.05) !important;
        }
        div.stButton > button:hover, div[data-testid="stFormSubmitButton"] > button:hover {
            background-color: #00234B !important;
            color: #FFFFFF !important;
            border-color: #00234B !important;
            transform: translateY(-1px);
        }
        [data-testid="stMetric"] {
            background-color: #FFFFFF;
            padding: 15px;
            border-radius: 8px;
            border-top: 4px solid #00234B;
            box-shadow: 0 2px 6px rgba(0,0,0,0.05);
        }
        [data-testid="stMetricValue"] {
            color: #00234B !important;
            font-weight: bold !important;
        }
        [data-testid="stMetricLabel"] {
            color: #4A5568 !important;
        }
    </style>
""", unsafe_allow_html=True)

# --- CABEÇALHO ---
col_logo, col_title = st.columns([1, 5])
with col_logo:
    try:
        st.image("ellenco.png", width=190)
    except Exception:
        st.write("🚚 **ELLENCO**")

with col_title:
    st.markdown("""
        <div class="main-header" style="margin-bottom: 0px; padding: 15px 20px;">
            <h1>Controle de Registro de Frete</h1>
            <p>Registro, gestão operacional e consolidação automática de relatórios em Excel.</p>
        </div>
    """, unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

lista_fretes = carregar_fretes()

# --- 1. FORMULÁRIO ---
with st.form("frete_form", clear_on_submit=True):
    st.subheader("➕ Adicionar Novo Frete")
    
    with st.container(border=True):
        st.markdown("##### 📌 Dados do Transporte")
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            data_frete_obj = st.date_input("📅 Data do Frete")
            data_frete = data_frete_obj.strftime("%d/%m/%Y")
        with c2:
            fretista = st.text_input("👤 Nome do Fretista", value="Carlos Silva")
        with c3:
            placa_veiculo = st.text_input("🚛 Placa do Veículo", value="ABC-1234")
        with c4:
            num_vale = st.text_input("🎟️ Nº do Vale", value="V-001")

    with st.container(border=True):
        st.markdown("##### 🏗️ Origem, Destino e Documentação")
        c1, c2, c3 = st.columns(3)
        with c1:
            obra = st.selectbox("🏗️ Obra / Destino", ["013.0413-2", "013.0427-2", "013.0434-2", "013.0437-2", "013.0438-2"])
        with c2:
            fornecedor = st.text_input("🏭 Fornecedor", value="Ellenco")
        with c3:
            nota_fiscal = st.text_input("📄 Nota Fiscal (NF)", value="12345")

    with st.container(border=True):
        st.markdown("##### ⚖️ Carga, Distância e Custos")
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            quantidade = st.number_input("⚖️ Quantidade (Ton)", min_value=0.0, value=0.0, step=0.1, format="%.2f")
        with c2:
            km = st.number_input("KM (Quilometragem)", min_value=0.0, value=0.0, step=1.0, format="%.1f")
        with c3:
            valor_frete = st.number_input("💵
