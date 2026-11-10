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
        # Ordem personalizada das colunas no SELECT
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
            obra = st.selectbox("🏗️ Obra / Destino", ["013.0427-2", "013.0434-2", "013.0437-2", "013.0438-2"])
        with c2:
            fornecedor = st.text_input("🏭 Fornecedor", value="Ellenco")
        with c3:
            nota_fiscal = st.text_input("📄 Nota Fiscal (NF)", value="12345")

    with st.container(border=True):
        st.markdown("##### ⚖️ Carga, Distância e Custos")
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            quantidade = st.number_input("⚖️ Quantidade (Ton)", min_value=0.0, value=2.5, step=0.1, format="%.2f")
        with c2:
            km = st.number_input("KM (Quilometragem)", min_value=0.0, value=45.0, step=1.0, format="%.1f")
        with c3:
            valor_frete = st.number_input("💵 Valor/Ton/KM (R$)", min_value=0.0, value=4.00, step=0.50, format="%.2f")
        with c4:
            pedagio = st.number_input("🛣️ Pedágio (R$)", min_value=0.0, value=12.50, step=1.00, format="%.2f")

    with st.container(border=True):
        st.markdown("##### 🏷️ Descontos e Adicionais")
        c1, c2, c3 = st.columns(3)
        with c1:
            tipo_desconto = st.selectbox("🏷️ Tipo de Desconto", ["Nenhum", "Diesel", "Marmitex"])
        with c2:
            valor_desconto = st.number_input("➖ Valor do Desconto (R$)", min_value=0.0, value=0.00, step=5.00, format="%.2f")
        with c3:
            adicional_diesel = st.number_input("⛽ Adicional Diesel (R$)", min_value=0.0, value=0.00, step=5.00, format="%.2f")

    km_total_preview = quantidade * km
    subtotal_preview = (km_total_preview * valor_frete) + pedagio
    total_geral_preview = subtotal_preview - valor_desconto + adicional_diesel

    st.markdown("---")
    
    m1, m2, m3, m4 = st.columns(4)
    m1.metric(label="Subtotal", value=f"R$ {subtotal_preview:,.2f}")
    m2.metric(label="Descontos", value=f"R$ {valor_desconto:,.2f}")
    m3.metric(label="Adicional Diesel", value=f"R$ {adicional_diesel:,.2f}")
    m4.metric(label="Total Geral Final", value=f"R$ {total_geral_preview:,.2f}")

    st.markdown("---")
    submitted_add = st.form_submit_button("➕ Adicionar Este Frete à Lista", use_container_width=True)

if submitted_add:
    novo_frete = {
        "Data": data_frete,
        "Fretista": fretista,
        "Placa": placa_veiculo,
        "Obra": obra,
        "Fornecedor": fornecedor,
        "Nota Fiscal": nota_fiscal,
        "Nº Vale": num_vale,
        "Qtd (Ton)": float(quantidade),
        "KM": float(km),
        "KM Total": float(km_total_preview),
        "Pedágio (R$)": float(pedagio),
        "Valor Frete (R$)": float(valor_frete),
        "Tipo Desconto": tipo_desconto,
        "Desconto (R$)": float(valor_desconto),
        "Adicional Diesel (R$)": float(adicional_diesel),
        "Total Geral (R$)": float(total_geral_preview)
    }
    salvar_novo_frete(novo_frete)
    st.success("Frete salvo com sucesso no banco de dados!")
    st.rerun()

# --- 2. TABELA INTERATIVA (CRUD) ---
if len(lista_fretes) > 0:
    st.markdown("---")
    st.subheader(f"📋 Controle de Registro de Frete ({len(lista_fretes)} registros salvos)")
    st.markdown("💡 *Você pode editar qualquer célula diretamente na tabela ou apagar linhas selecionando-as.*")

    df_current = pd.DataFrame(lista_fretes)

    df_edited = st.data_editor(
        df_current,
        num_rows="dynamic",
        use_container_width=True,
        key="crud_editor"
    )

    col_save, col_btn1, col_btn2 = st.columns([1, 1, 1])

    with col_save:
        if st.button("💾 Salvar Edições da Tabela", use_container_width=True):
            for idx, row in df_edited.iterrows():
                try:
                    q = float(row["Qtd (Ton)"])
                    k = float(row["KM"])
                    vf = float(row["Valor Frete (R$)"])
                    ped = float(row["Pedágio (R$)"])
                    desc = float(row["Desconto (R$)"])
                    adic = float(row["Adicional Diesel (R$)"])
                    
                    km_tot = q * k
                    df_edited.loc[idx, "KM Total"] = km_tot
                    # Ordem das colunas: Qtd (Ton) = Col G (7), KM = Col E (5), Valor Frete = Col J (10), Pedágio = Col M (13), Desconto = Col L (12), Adicional Diesel = Col N (14)
                    # Fórmula estruturada: ((Qtd * KM) * Valor_Frete) + Pedagio - Desconto + Adicional_Diesel
                    df_edited.loc[idx, "Total Geral (R$)"] = ((km_tot * vf) + ped) - desc + adic
                except Exception:
                    pass
            atualizar_banco_inteiro(df_edited)
            st.success("Alterações salvas com sucesso!")
            st.rerun()

    with col_btn1:
        confirmar_limpeza = st.checkbox("⚠️ Confirmar exclusão total")
        if st.button("🗑️ Limpar Todos os Registros", use_container_width=True):
            if confirmar_limpeza:
                limpar_banco()
                st.success("Todos os registros foram apagados!")
                st.rerun()
            else:
                st.warning("Marque a caixa de confirmação acima para apagar os registros.")

    # --- 3. GERAÇÃO DO EXCEL CONSOLIDADO ---
    try:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Controle de Registro"
        ws.views.sheetView[0].showGridLines = True

        ELLENCO_NAVY = "00234B"
        TOTAL_BG = "E6EEF4"
        ZEBRA_BG = "F8FAFC"
        BORDER_GRAY = "D9D9D9"

        font_title = Font(name="Segoe UI", size=16, bold=True, color=ELLENCO_NAVY)
        font_header = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
        font_regular = Font(name="Segoe UI", size=11)
        font_bold = Font(name="Segoe UI", size=11, bold=True, color=ELLENCO_NAVY)

        fill_header = PatternFill(start_color=ELLENCO_NAVY, end_color=ELLENCO_NAVY, fill_type="solid")
        fill_zebra = PatternFill(start_color=ZEBRA_BG, end_color=ZEBRA_BG, fill_type="solid")
        fill_total = PatternFill(start_color=TOTAL_BG, end_color=TOTAL_BG, fill_type="solid")

        thin_border = Border(
            left=Side(style='thin', color=BORDER_GRAY),
            right=Side(style='thin', color=BORDER_GRAY),
            top=Side(style='thin', color=BORDER_GRAY),
            bottom=Side(style='thin', color=BORDER_GRAY)
        )

        align_center = Alignment(horizontal="center", vertical="center")
        align_left = Alignment(horizontal="left", vertical="center")
        align_right = Alignment(horizontal="right", vertical="center")

        ws.merge_cells("A1:O1")
        ws["A1"] = "ELLENCO - CONTROLE DE REGISTRO DE FRETE"
        ws["A1"].font = font_title

        # Ordem solicitada das colunas no Excel:
        # 1: Fretista, 2: Data, 3: Nº Vale, 4: Nota Fiscal, 5: KM, 6: Obra, 7: Qtde (ton), 8: Placa, 9: Fornecedor, 
        # 10: Valor do frete, 11: Tipo desconto, 12: Desconto, 13: Pedagio, 14: Adicional diesel, 15: Total geral
        headers = [
            "Fretista", "Data", "Nº Vale", "Nota Fiscal", "KM", "Obra", "Qtde (ton)", 
            "Placa", "Fornecedor", "Valor do frete (R$)", "Tipo Desconto", "Desconto (R$)", 
            "Pedágio (R$)", "Adicional Diesel (R$)", "Total Geral (R$)"
        ]

        for col_idx, header in enumerate(headers, 1):
            cell = ws.cell(row=3, column=col_idx, value=header)
            cell.font = font_header
            cell.fill = fill_header
            cell.alignment = align_center
            cell.border = thin_border

        start_row = 4
        for i, item in enumerate(lista_fretes):
            r = start_row + i
            
            # Mapeamento estrito seguindo a ordem das colunas:
            # A (1): Fretista
            ws.cell(row=r, column=1, value=str(item.get("Fretista", ""))).alignment = align_left
            # B (2): Data
            ws.cell(row=r, column=2, value=str(item.get("Data", ""))).alignment = align_center
            # C (3): Nº Vale
            ws.cell(row=r, column=3, value=str(item.get("Nº Vale", ""))).alignment = align_center
            # D (4): Nota Fiscal
            ws.cell(row=r, column=4, value=str(item.get("Nota Fiscal", ""))).alignment = align_center
            # E (5): KM
            ws.cell(row=r, column=5, value=item.get("KM", 0)).alignment = align_right
            # F (6): Obra
            ws.cell(row=r, column=6, value=str(item.get("Obra", ""))).alignment = align_center
            # G (7): Qtde (ton)
            ws.cell(row=r, column=7, value=item.get("Qtd (Ton)", 0)).alignment = align_right
            # H (8): Placa
            ws.cell(row=r, column=8, value=str(item.get("Placa", ""))).alignment = align_center
            # I (9): Fornecedor
            ws.cell(row=r, column=9, value=str(item.get("Fornecedor", ""))).alignment = align_left
            # J (10): Valor do frete
            ws.cell(row=r, column=10, value=item.get("Valor Frete (R$)", 0)).alignment = align_right
            # K (11): Tipo Desconto
            ws.cell(row=r, column=11, value=str(item.get("Tipo Desconto", "Nenhum"))).alignment = align_center
            # L (12): Desconto
            ws.cell(row=r, column=12, value=item.get("Desconto (R$)", 0)).alignment = align_right
            # M (13): Pedágio
            ws.cell(row=r, column=13, value=item.get("Pedágio (R$)", 0)).alignment = align_right
            # N (14): Adicional Diesel
            ws.cell(row=r, column=14, value=item.get("Adicional Diesel (R$)", 0)).alignment = align_right
            # O (15): Total Geral -> Fórmula: ((Qtde * KM) * Valor_Frete) + Pedagio - Desconto + Adicional_Diesel -> ((G{r}*E{r})*J{r})+M{r}-L{r}+N{r}
            ws.cell(row=r, column=15, value=f"=((G{r}*E{r})*J{r})+M{r}-L{r}+N{r}").alignment = align_right
            
            # Formatações numéricas
            ws.cell(row=r, column=5).number_format = "#,##0.0"
            ws.cell(row=r, column=7).number_format = "#,##0.00"
            ws.cell(row=r, column=10).number_format = 'R$ #,##0.00'
            ws.cell(row=r, column=12).number_format = 'R$ #,##0.00'
            ws.cell(row=r, column=13).number_format = 'R$ #,##0.00'
            ws.cell(row=r, column=14).number_format = 'R$ #,##0.00'
            ws.cell(row=r, column=15).number_format = 'R$ #,##0.00'
            
            for c in range(1, 16):
                cell = ws.cell(row=r, column=c)
                cell.font = font_regular
                cell.border = thin_border
                if i % 2 == 1:
                    cell.fill = fill_zebra

        tot_row = start_row + len(lista_fretes)
        ws.cell(row=tot_row, column=1, value="TOTAL GERAL").font = font_bold
        ws.cell(row=tot_row, column=5, value=f"=SUM(E{start_row}:E{tot_row-1})")
        ws.cell(row=tot_row, column=7, value=f"=SUM(G{start_row}:G{tot_row-1})")
        ws.cell(row=tot_row, column=10, value=f"=AVERAGE(J{start_row}:J{tot_row-1})")
        ws.cell(row=tot_row, column=12, value=f"=SUM(L{start_row}:L{tot_row-1})")
        ws.cell(row=tot_row, column=13, value=f"=SUM(M{start_row}:M{tot_row-1})")
        ws.cell(row=tot_row, column=14, value=f"=SUM(N{start_row}:N{tot_row-1})")
        ws.cell(row=tot_row, column=15, value=f"=SUM(O{start_row}:O{tot_row-1})")

        ws.cell(row=tot_row, column=5).number_format = "#,##0.0"
        ws.cell(row=tot_row, column=7).number_format = "#,##0.00"
        ws.cell(row=tot_row, column=10).number_format = 'R$ #,##0.00'
        ws.cell(row=tot_row, column=12).number_format = 'R$ #,##0.00'
        ws.cell(row=tot_row, column=13).number_format = 'R$ #,##0.00'
        ws.cell(row=tot_row, column=14).number_format = 'R$ #,##0.00'
        ws.cell(row=tot_row, column=15).number_format = 'R$ #,##0.00'

        for c in range(1, 16):
            cell = ws.cell(row=tot_row, column=c)
            cell.font = font_bold
            cell.fill = fill_total
            cell.border = thin_border

        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_len + 4, 15)

        output = io.BytesIO()
        wb.save(output)
        output.seek(0)

        with col_btn2:
            st.download_button(
                label="📥 Gerar Planilha Excel",
                data=output,
                file_name="Controle_de_Registro_de_Frete_Ellenco.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
    except Exception as e:
        st.error(f"Ocorreu um erro ao gerar o arquivo Excel: {e}")
