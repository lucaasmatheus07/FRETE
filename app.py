import io
from datetime import datetime
import openpyxl
import pandas as pd
import streamlit as st
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# Configuração da página
st.set_page_config(
    page_title="Controle de Fretes - Ellenco",
    page_icon="🚚",
    layout="wide"
)

# --- ESTILIZAÇÃO CSS CUSTOMIZADA (TEMA ELLENCO) ---
st.markdown("""
    <style>
        /* Cores da Ellenco: Azul Escuro #00234B | Vermelho #D9232A */
        .stApp {
            background-color: #f4f6f9;
        }
        
        /* Cabeçalho do Título */
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

        /* Estilização dos Containers / Cards */
        [data-testid="stForm"] {
            background-color: #FFFFFF;
            border-radius: 10px;
            border: 1px solid #E2E8F0;
            box-shadow: 0 4px 12px rgba(0,0,0,0.05);
            padding: 20px;
        }

        /* Títulos dentro do Formulário */
        h5 {
            color: #00234B !important;
            border-left: 4px solid #D9232A;
            padding-left: 8px;
            margin-bottom: 12px !important;
            font-weight: 700 !important;
        }

        /* Botões de Ação com fundo BRANCO e borda azul/cinza */
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

        /* Estilização dos Cartões de Métrica */
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

# --- CABEÇALHO COM LOGÓTIPO DA RAIZ DO GITHUB E TEMA VISUAL ---
col_logo, col_title = st.columns([1, 5])
with col_logo:
    try:
        # Aumentado o parâmetro width para 190 (um pouco maior)
        st.image("ellenco.png", width=190) 
    except Exception:
        st.write("🚚 **ELLENCO**")

with col_title:
    st.markdown("""
        <div class="main-header" style="margin-bottom: 0px; padding: 15px 20px;">
            <h1>Sistema de Controlo de Fretes e Logística</h1>
            <p>Registo, gestão operacional e consolidação automática de relatórios em Excel.</p>
        </div>
    """, unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# Inicializar a lista de fretes na memória da sessão
if "lista_fretes" not in st.session_state:
    st.session_state.lista_fretes = []

# --- 1. FORMULÁRIO COM O TEMA ELLENCO ---
with st.form("frete_form", clear_on_submit=True):
    st.subheader("➕ Adicionar Novo Frete")
    
    # Bloco 1: Dados Gerais do Transporte
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

    # Bloco 2: Destino, Fornecedor e Documentação
    with st.container(border=True):
        st.markdown("##### 🏗️ Origem, Destino e Documentação")
        c1, c2, c3 = st.columns(3)
        with c1:
            obra = st.selectbox("🏗️ Obra / Destino", ["013.0427-2", "013.0434-2", "013.0437-2", "013.0438-2"])
        with c2:
            fornecedor = st.text_input("🏭 Fornecedor", value="Ellenco")
        with c3:
            nota_fiscal = st.text_input("📄 Nota Fiscal (NF)", value="12345")

    # Bloco 3: Quilometragem, Carga e Valores
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

    # Bloco 4: Descontos
    with st.container(border=True):
        st.markdown("##### 🏷️ Descontos e Abatimentos")
        c1, c2 = st.columns(2)
        with c1:
            tipo_desconto = st.selectbox("🏷️ Tipo de Desconto", ["Nenhum", "Diesel", "Marmitex"])
        with c2:
            valor_desconto = st.number_input("➖ Valor do Desconto (R$)", min_value=0.0, value=0.00, step=5.00, format="%.2f")

    # --- CÁLCULO E PRÉ-VISUALIZAÇÃO EM TEMPO REAL ---
    km_total_preview = quantidade * km
    subtotal_preview = (km_total_preview * valor_frete) + pedagio
    total_geral_preview = subtotal_preview - valor_desconto

    st.markdown("---")
    
    # Exibição de Métricas
    m1, m2, m3, m4 = st.columns(4)
    m1.metric(label="KM Total (Qtd x KM)", value=f"{km_total_preview:,.2f} km")
    m2.metric(label="Subtotal da Viagem", value=f"R$ {subtotal_preview:,.2f}")
    m3.metric(label="Desconto Aplicado", value=f"R$ {valor_desconto:,.2f}")
    m4.metric(label="Total Geral Final", value=f"R$ {total_geral_preview:,.2f}")

    st.markdown("---")
    submitted_add = st.form_submit_button("➕ Adicionar Este Frete à Lista", use_container_width=True)

if submitted_add:
    st.session_state.lista_fretes.append({
        "Data": data_frete,
        "Fretista": fretista,
        "Placa": placa_veiculo,
        "Obra": obra,
        "Fornecedor": fornecedor,
        "NF": nota_fiscal,
        "Nº Vale": num_vale,
        "Qtd (Ton)": float(quantidade),
        "KM": float(km),
        "KM Total": float(km_total_preview),
        "Pedágio (R$)": float(pedagio),
        "Valor Frete (R$)": float(valor_frete),
        "Tipo Desconto": tipo_desconto,
        "Desconto (R$)": float(valor_desconto),
        "Total Geral (R$)": float(total_geral_preview)
    })
    st.success("Frete adicionado com sucesso!")
    st.rerun()

# --- 2. CONTROLE DE CRUD (TABELA INTERATIVA) ---
if len(st.session_state.lista_fretes) > 0:
    st.markdown("---")
    st.subheader(f"📋 Controlo de Registos ({len(st.session_state.lista_fretes)} registos)")
    st.markdown("💡 *Pode editar qualquer célula diretamente na tabela, adicionar novas linhas ou apagar linhas selecionando-as e clicando na lixeira.*")

    df_current = pd.DataFrame(st.session_state.lista_fretes)

    df_edited = st.data_editor(
        df_current,
        num_rows="dynamic",
        use_container_width=True,
        key="crud_editor"
    )

    # Recálculo automático em caso de edição manual na tabela
    for idx, row in df_edited.iterrows():
        try:
            q = float(row["Qtd (Ton)"])
            k = float(row["KM"])
            vf = float(row["Valor Frete (R$)"])
            ped = float(row["Pedágio (R$)"])
            desc = float(row["Desconto (R$)"])
            
            df_edited.loc[idx, "KM Total"] = q * k
            df_edited.loc[idx, "Total Geral (R$)"] = ((k * q * vf) + ped) - desc
        except Exception:
            pass

    st.session_state.lista_fretes = df_edited.to_dict(orient="records")

    col_btn1, col_btn2 = st.columns(2)

    if col_btn1.button("🗑️ Limpar Todos os Registos", use_container_width=True):
        st.session_state.lista_fretes = []
        st.rerun()

    # --- 3. GERAÇÃO DA PLANILHA EXCEL CONSOLIDADA ---
    try:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Controle de Fretes"
        ws.views.sheetView[0].showGridLines = True

        ELLENCO_NAVY = "00234B"
        ELLENCO_RED = "D9232A"
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
        ws["A1"] = "ELLENCO - REGISTRO DE CONTROLE DE FRETES E LOGÍSTICA"
        ws["A1"].font = font_title

        headers = [
            "Data", "Fretista", "Placa", "Obra", "Fornecedor", "Nota Fiscal", "Nº Vale",
            "Quantidade (Ton)", "KM", "KM Total (Qtd x KM)", 
            "Pedágio (R$)", "Valor do Frete (R$)", "Tipo Desconto", "Desconto (R$)", "Total Geral (R$)"
        ]

        for col_idx, header in enumerate(headers, 1):
            cell = ws.cell(row=3, column=col_idx, value=header)
            cell.font = font_header
            cell.fill = fill_header
            cell.alignment = align_center
            cell.border = thin_border

        start_row = 4
        for i, item in enumerate(st.session_state.lista_fretes):
            r = start_row + i
            
            ws.cell(row=r, column=1, value=str(item.get("Data", ""))).alignment = align_center
            ws.cell(row=r, column=2, value=str(item.get("Fretista", ""))).alignment = align_left
            ws.cell(row=r, column=3, value=str(item.get("Placa", ""))).alignment = align_center
            ws.cell(row=r, column=4, value=str(item.get("Obra", ""))).alignment = align_center
            ws.cell(row=r, column=5, value=str(item.get("Fornecedor", ""))).alignment = align_left
            ws.cell(row=r, column=6, value=str(item.get("NF", ""))).alignment = align_center
            ws.cell(row=r, column=7, value=str(item.get("Nº Vale", ""))).alignment = align_center
            ws.cell(row=r, column=8, value=item.get("Qtd (Ton)", 0)).alignment = align_right
            ws.cell(row=r, column=9, value=item.get("KM", 0)).alignment = align_right
            
            ws.cell(row=r, column=10, value=f"=H{r}*I{r}").alignment = align_right
            ws.cell(row=r, column=11, value=item.get("Pedágio (R$)", 0)).alignment = align_right
            ws.cell(row=r, column=12, value=item.get("Valor Frete (R$)", 0)).alignment = align_right
            ws.cell(row=r, column=13, value=str(item.get("Tipo Desconto", "Nenhum"))).alignment = align_center
            ws.cell(row=r, column=14, value=item.get("Desconto (R$)", 0)).alignment = align_right
            ws.cell(row=r, column=15, value=f"=((J{r}*L{r})+K{r})-N{r}").alignment = align_right
            
            ws.cell(row=r, column=8).number_format = "#,##0.00"
            ws.cell(row=r, column=9).number_format = "#,##0.0"
            ws.cell(row=r, column=10).number_format = "#,##0.00"
            ws.cell(row=r, column=11).number_format = 'R$ #,##0.00'
            ws.cell(row=r, column=12).number_format = 'R$ #,##0.00'
            ws.cell(row=r, column=14).number_format = 'R$ #,##0.00'
            ws.cell(row=r, column=15).number_format = 'R$ #,##0.00'
            
            for c in range(1, 16):
                cell = ws.cell(row=r, column=c)
                cell.font = font_regular
                cell.border = thin_border
                if i % 2 == 1:
                    cell.fill = fill_zebra

        tot_row = start_row + len(st.session_state.lista_fretes)
        ws.cell(row=tot_row, column=1, value="TOTAL GERAL").font = font_bold
        ws.cell(row=tot_row, column=8, value=f"=SUM(H{start_row}:H{tot_row-1})")
        ws.cell(row=tot_row, column=9, value=f"=SUM(I{start_row}:I{tot_row-1})")
        ws.cell(row=tot_row, column=10, value=f"=SUM(J{start_row}:J{tot_row-1})")
        ws.cell(row=tot_row, column=11, value=f"=SUM(K{start_row}:K{tot_row-1})")
        ws.cell(row=tot_row, column=12, value=f"=AVERAGE(L{start_row}:L{tot_row-1})")
        ws.cell(row=tot_row, column=14, value=f"=SUM(N{start_row}:N{tot_row-1})")
        ws.cell(row=tot_row, column=15, value=f"=SUM(O{start_row}:O{tot_row-1})")

        ws.cell(row=tot_row, column=8).number_format = "#,##0.00"
        ws.cell(row=tot_row, column=9).number_format = "#,##0.0"
        ws.cell(row=tot_row, column=10).number_format = "#,##0.00"
        ws.cell(row=tot_row, column=11).number_format = 'R$ #,##0.00'
        ws.cell(row=tot_row, column=12).number_format = 'R$ #,##0.00'
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

        col_btn2.download_button(
            label="📥 Descarregar Planilha Consolidada",
            data=output,
            file_name="Controle_de_Fretes_Ellenco.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    except Exception as e:
        st.error(f"Ocorreu um erro ao gerar o arquivo Excel: {e}")
