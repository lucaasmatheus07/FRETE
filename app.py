import io
import openpyxl
import pandas as pd
import streamlit as st
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# Configuração da página
st.set_page_config(
    page_title="Controle de Fretes",
    page_icon="🚚",
    layout="wide"
)

st.title("🚚 Sistema de Controlo de Fretes e Logística (com CRUD)")
st.markdown("Insira novos fretes abaixo ou **edite, adicione e exclua registros diretamente na tabela interativa**.")

# Inicializar a lista de fretes na memória da sessão
if "lista_fretes" not in st.session_state:
    st.session_state.lista_fretes = []

# --- 1. FORMULÁRIO PARA ADICIONAR NOVO FRETE ---
with st.form("frete_form", clear_on_submit=True):
    st.subheader("➕ Adicionar Novo Frete à Lista")
    col1, col2, col3 = st.columns(3)
    
    with col1:
        data_frete = str(st.date_input("Data do Frete"))
        fretista = st.text_input("Nome do Fretista", "Carlos Silva")
        placa_veiculo = st.text_input("Placa do Veículo", "ABC-1234")
        
    with col2:
        obra = st.text_input("Obra", "Obra Residencial Alpha")
        fornecedor = st.text_input("Fornecedor", "Ellenco")
        nota_fiscal = st.text_input("Nota Fiscal (NF)", "12345")
        
    with col3:
        quantidade = st.number_input("Quantidade (Ton)", min_value=0.0, value=2.5, step=0.1, format="%.2f")
        km = st.number_input("KM (Quilometragem)", min_value=0.0, value=45.0, format="%.1f")
        pedagio = st.number_input("Valor do Pedágio (R$)", min_value=0.0, value=12.50, format="%.2f")
        valor_frete = st.number_input("Valor do Frete (R$)", min_value=0.0, value=4.00, format="%.2f")
        
        tipo_desconto = st.selectbox("Tipo de Desconto", ["Nenhum", "Diesel", "Marmitex"])
        valor_desconto = st.number_input("Valor do Desconto (R$)", min_value=0.0, value=0.00, format="%.2f")

    # --- CÁLCULO EM TEMPO REAL ---
    subtotal_preview = (km * quantidade * valor_frete) + pedagio
    total_geral_preview = subtotal_preview - valor_desconto

    st.markdown("---")
    st.markdown(f"**Pré-visualização do Total deste Frete (com desconto):** R$ {total_geral_preview:,.2f}")
    
    submitted_add = st.form_submit_button("➕ Adicionar Este Frete à Lista")

if submitted_add:
    subtotal_viagem = (km * quantidade * valor_frete) + pedagio
    total_geral_viagem = subtotal_viagem - valor_desconto
    km_total_viagem = quantidade * km

    st.session_state.lista_fretes.append({
        "Data": data_frete,
        "Fretista": fretista,
        "Placa": placa_veiculo,
        "Obra": obra,
        "Fornecedor": fornecedor,
        "NF": nota_fiscal,
        "Qtd (Ton)": float(quantidade),
        "KM": float(km),
        "KM Total": float(km_total_viagem),
        "Pedágio (R$)": float(pedagio),
        "Valor Frete (R$)": float(valor_frete),
        "Tipo Desconto": tipo_desconto,
        "Desconto (R$)": float(valor_desconto),
        "Total Geral (R$)": float(total_geral_viagem)
    })
    st.success("Frete adicionado com sucesso!")
    st.rerun()

# --- 2. CONTROLE DE CRUD (TABELA INTERATIVA) ---
if len(st.session_state.lista_fretes) > 0:
    st.markdown("---")
    st.subheader(f"📋 Controlo de Registos (CRUD) - ({len(st.session_state.lista_fretes)} registos)")
    st.markdown("💡 *Pode editar qualquer célula diretamente na tabela, adicionar novas linhas ou apagar linhas selecionando-as e clicando na lixeira.*")

    # Converter lista em DataFrame para edição interativa
    df_current = pd.DataFrame(st.session_state.lista_fretes)

    # Tabela com suporte a edição, adição e exclusão de linhas (CRUD interativo)
    df_edited = st.data_editor(
        df_current,
        num_rows="dynamic",
        use_container_width=True,
        key="crud_editor"
    )

    # Atualizar automaticamente os cálculos baseados nas edições feitas na tabela
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

    # Atualizar o session_state com os dados editados
    st.session_state.lista_fretes = df_edited.to_dict(orient="records")

    col_btn1, col_btn2 = st.columns(2)

    # Botão para limpar toda a lista
    if col_btn1.button("🗑️ Limpar Todos os Registos", use_container_width=True):
        st.session_state.lista_fretes = []
        st.rerun()

    # --- 3. GERAÇÃO DA PLANILHA EXCEL CONSOLIDADA ---
    try:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Controle de Fretes"
        ws.views.sheetView[0].showGridLines = True

        NAVY = "1B365D"
        TOTAL_BG = "E6EEF4"
        ZEBRA_BG = "F8FAFC"
        BORDER_GRAY = "D9D9D9"

        font_title = Font(name="Segoe UI", size=16, bold=True, color=NAVY)
        font_header = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
        font_regular = Font(name="Segoe UI", size=11)
        font_bold = Font(name="Segoe UI", size=11, bold=True, color=NAVY)

        fill_header = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")
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

        ws.merge_cells("A1:N1")
        ws["A1"] = "REGISTRO DE CONTROLE DE FRETES E LOGÍSTICA"
        ws["A1"].font = font_title

        headers = [
            "Data", "Fretista", "Placa", "Obra", "Fornecedor", "Nota Fiscal", 
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
            ws.cell(row=r, column=4, value=str(item.get("Obra", ""))).alignment = align_left
            ws.cell(row=r, column=5, value=str(item.get("Fornecedor", ""))).alignment = align_left
            ws.cell(row=r, column=6, value=str(item.get("NF", ""))).alignment = align_center
            ws.cell(row=r, column=7, value=item.get("Qtd (Ton)", 0)).alignment = align_right
            ws.cell(row=r, column=8, value=item.get("KM", 0)).alignment = align_right
            
            # Fórmulas Excel aplicadas nas colunas calculadas
            ws.cell(row=r, column=9, value=f"=G{r}*H{r}").alignment = align_right
            ws.cell(row=r, column=10, value=item.get("Pedágio (R$)", 0)).alignment = align_right
            ws.cell(row=r, column=11, value=item.get("Valor Frete (R$)", 0)).alignment = align_right
            ws.cell(row=r, column=12, value=str(item.get("Tipo Desconto", "Nenhum"))).alignment = align_center
            ws.cell(row=r, column=13, value=item.get("Desconto (R$)", 0)).alignment = align_right
            ws.cell(row=r, column=14, value=f"=((I{r}*K{r})+J{r})-M{r}").alignment = align_right
            
            # Formatação numérica
            ws.cell(row=r, column=7).number_format = "#,##0.00"
            ws.cell(row=r, column=8).number_format = "#,##0.0"
            ws.cell(row=r, column=9).number_format = "#,##0.00"
            ws.cell(row=r, column=10).number_format = 'R$ #,##0.00'
            ws.cell(row=r, column=11).number_format = 'R$ #,##0.00'
            ws.cell(row=r, column=13).number_format = 'R$ #,##0.00'
            ws.cell(row=r,
