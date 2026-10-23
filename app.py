import streamlit as st
import openpyxl
import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io

# Configuração da página
st.set_page_config(
    page_title="Controle de Fretes",
    page_icon="🚚",
    layout="wide"
)

st.title("🚚 Sistema de Controlo de Fretes e Logística")
st.markdown("Insira e **guarde os fretes um por um**. O valor total de cada viagem aparece na tabela abaixo.")

# Inicializar a lista de fretes na memória da sessão
if "lista_fretes" not in st.session_state:
    st.session_state.lista_fretes = []

with st.form("frete_form", clear_on_submit=True):
    st.subheader("Adicionar Novo Frete à Lista")
    col1, col2, col3 = st.columns(3)
    
    with col1:
        data_frete = st.date_input("Data do Frete")
        fretista = st.text_input("Nome do Fretista", "Carlos Silva")
        obra = st.text_input("Obra", "Obra Residencial Alpha")
        
    with col2:
        fornecedor = st.text_input("Fornecedor", "Ellenco")
        # Quantidade em toneladas (aceita números quebrados)
        quantidade = st.number_input("Quantidade (Ton)", min_value=0.0, value=2.5, step=0.1, format="%.2f")
        km = st.number_input("KM (Quilometragem)", min_value=0.0, value=45.0, format="%.1f")
        
    with col3:
        pedagio = st.number_input("Valor do Pedágio (R$)", min_value=0.0, value=12.50, format="%.2f")
        valor_frete = st.number_input("Valor do Frete (R$)", min_value=0.0, value=4.00, format="%.2f")

    # --- CÁLCULO EM TEMPO REAL: (KM * Quantidade * Valor do Frete) + Pedágio ---
    total_geral_preview = (km * quantidade * valor_frete) + pedagio

    st.markdown("---")
    st.markdown(f"**Pré-visualização do Total deste Frete:** R$ {total_geral_preview:,.2f}")
    
    submitted_add = st.form_submit_button("➕ Adicionar Este Frete à Lista")

if submitted_add:
    # Calcular o total geral da viagem para guardar na lista
    total_geral_viagem = (km * quantidade * valor_frete) + pedagio
    km_total_viagem = quantidade * km

    # Adicionar o registro atual na lista da sessão
    st.session_state.lista_fretes.append({
        "Data": str(data_frete),
        "Fretista": fretista,
        "Obra": obra,
        "Fornecedor": fornecedor,
        "Qtd (Ton)": quantidade,
        "KM": km,
        "KM Total": km_total_viagem,
        "Pedágio (R$)": pedagio,
        "Valor Frete (R$)": valor_frete,
        "Total Geral (R$)": total_geral_viagem
    })
    st.success("Frete adicionado com sucesso! Veja o registo na tabela abaixo.")

# Se houver fretes guardados, exibe a tabela acumulada com o Total Geral e o botão de exportar
if len(st.session_state.lista_fretes) > 0:
    st.markdown("---")
    st.subheader(f"📋 Fretes Acumulados ({len(st.session_state.lista_fretes)} registos)")
    
    # Exibir tabela prévia na tela com os totais calculados
    df_preview = pd.DataFrame(st.session_state.lista_fretes)
    
    # Formatar colunas monetárias e numéricas na exibição visual da tabela
    st.dataframe(df_preview, use_container_width=True)

    col_btn1, col_btn2 = st.columns(2)

    # Botão para limpar a lista
    if col_btn1.button("🗑️ Limpar Todos os Registos"):
        st.session_state.lista_fretes = []
        st.rerun()

    # Gerar a Planilha Excel consolidada com todos os itens guardados
    try:
        wb = openpyxl.Workbook()
        ws_dash = wb.active
        ws_dash.title = "Resumo Geral"
        ws_dash.views.sheetView[0].showGridLines = True

        ws = wb.create_sheet(title="Controle de Fretes")
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

        ws.merge_cells("A1:J1")
        ws["A1"] = "REGISTRO DE CONTROLE DE FRETES E LOGÍSTICA"
        ws["A1"].font = font_title

        headers = [
            "Data", "Fretista", "Obra", "Fornecedor", 
            "Quantidade (Ton)", "KM", "KM Total (Qtd x KM)", 
            "Pedágio (R$)", "Valor do Frete (R$)", "Total Geral (R$)"
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
            
            ws.cell(row=r, column=1, value=item["Data"]).alignment = align_center
            ws.cell(row=r, column=2, value=item["Fretista"]).alignment = align_left
            ws.cell(row=r, column=3, value=item["Obra"]).alignment = align_left
            ws.cell(row=r, column=4, value=item["Fornecedor"]).alignment = align_left
            ws.cell(row=r, column=5, value=item["Qtd (Ton)"]).alignment = align_right
            ws.cell(row=r, column=6, value=item["KM"]).alignment = align_right
            
            # FÓRMULA DE MULTIPLICAÇÃO: Quantidade x KM (Coluna G)
            ws.cell(row=r, column=7, value=f"=E{r}*F{r}").alignment = align_right
            
            ws.cell(row=r, column=8, value=item["Pedágio (R$)"]).alignment = align_right
            ws.cell(row=r, column=9, value=item["Valor Frete (R$)"]).alignment = align_right
            
            # FÓRMULA PRINCIPAL: (KM x Quantidade x Valor do Frete) + Pedágio -> =(G * Valor_Frete) + Pedágio
            ws.cell(row=r, column=10, value=f"=(G{r}*I{r})+H{r}").alignment = align_right
            
            ws.cell(row=r, column=1).number_format = "yyyy-mm-dd"
            ws.cell(row=r, column=5).number_format = "#,##0.00"
            ws.cell(row=r, column=6).number_format = "#,##0.0"
            ws.cell(row=r, column=7).number_format = "#,##0.00"
            ws.cell(row=r, column=8).number_format = 'R$ #,##0.00'
            ws.cell(row=r, column=9).number_format = 'R$ #,##0.00'
            ws.cell(row=r, column=10).number_format = 'R$ #,##0.00'
            
            for c in range(1, 11):
                cell = ws.cell(row=r, column=c)
                cell.font = font_regular
                cell.border = thin_border
                if i % 2 == 1:
                    cell.fill = fill_zebra

        tot_row = start_row + len(st.session_state.lista_fretes)
        ws.cell(row=tot_row, column=1, value="TOTAL GERAL").font = font_bold
        ws.cell(row=tot_row, column=5, value=f"=SUM(E{start_row}:E{tot_row-1})")
        ws.cell(row=tot_row, column=6, value=f"=SUM(F{start_row}:F{tot_row-1})")
        ws.cell(row=tot_row, column=7, value=f"=SUM(G{start_row}:G{tot_row-1})")
        ws.cell(row=tot_row, column=8, value=f"=SUM(H{start_row}:H{tot_row-1})")
        ws.cell(row=tot_row, column=9, value=f"=SUM(I{start_row}:I{tot_row-1})")
        ws.cell(row=tot_row, column=10, value=f"=SUM(J{start_row}:J{tot_row-1})")

        ws.cell(row=tot_row, column=5).number_format = "#,##0.00"
        ws.cell(row=tot_row, column=7).number_format = "#,##0.00"

        for c in range(1, 11):
            cell = ws.cell(row=tot_row, column=c)
            cell.font = font_bold
            cell.fill = fill_total
            cell.border = thin_border

        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_len + 4, 15)

        # Salvar em buffer de memória para download
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)

        col_btn2.download_button(
            label="📥 Descarregar Planilha Consolidada",
            data=output,
            file_name="Controle_de_Fretes.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    except Exception as e:
        st.error(f"Ocorreu um erro ao gerar o arquivo Excel: {e}")
