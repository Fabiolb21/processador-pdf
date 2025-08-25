import streamlit as st
import fitz  # PyMuPDF
import pandas as pd
import re
import io # Necessário para salvar o Excel na memória

# --- Configuração da Página ---
st.set_page_config(
    page_title="Extrator de Pedidos",
    page_icon="📄",
    layout="wide"
)

def extrair_informacoes_pdf(arquivo_pdf):
    """
    Extrai informações de PRODUTO, QTD. e o NÚMERO DA PÁGINA impresso no documento.
    """
    dados_finais = []
    try:
        documento = fitz.open(stream=arquivo_pdf.read(), filetype="pdf")
        for pagina_num in range(len(documento)):
            pagina = documento.load_page(pagina_num)
            blocos = pagina.get_text("blocks")
            numero_pagina_impresso = "N/A"

            for b in blocos:
                texto_bloco = b[4]
                if "PÁGINA:" in texto_bloco:
                    match = re.search(r'(\d+)', texto_bloco)
                    if match:
                        numero_pagina_impresso = match.group(1)
                        break
            
            produtos_com_coords = []
            qtds_com_coords = []
            coord_x_produto_inicio = 0; coord_x_produto_fim = 100
            coord_x_qtd_inicio = 400; coord_x_qtd_fim = 450

            for b in blocos:
                x0, y0, _, _, texto, _, _ = b
                texto_limpo = texto.strip()
                if coord_x_produto_inicio < x0 < coord_x_produto_fim:
                    match = re.search(r'(JBGF\d+)', texto_limpo)
                    if match:
                        produtos_com_coords.append({'produto': match.group(1), 'y': y0})
                if coord_x_qtd_inicio < x0 < coord_x_qtd_fim and texto_limpo.isdigit():
                    qtds_com_coords.append({'qtd': int(texto_limpo), 'y': y0})

            for prod in produtos_com_coords:
                qtd_correspondente = next((qtd['qtd'] for qtd in qtds_com_coords if abs(prod['y'] - qtd['y']) < 10), 0)
                dados_finais.append({"PRODUTO": prod['produto'], "QTD.": qtd_correspondente, "PÁGINA": numero_pagina_impresso})
        documento.close()
    except Exception as e:
        st.error(f"Erro ao processar o arquivo {arquivo_pdf.name}: {e}")
    return dados_finais

# --- Interface Principal da Aplicação ---

st.title("📄 Extrator de Pedidos de Arquivos PDF")
st.markdown("Faça o upload de um ou mais arquivos de pedido para extrair os dados de forma rápida e precisa.")

with st.container(border=True):
    st.header("📤 1. Faça o Upload dos Arquivos")
    arquivos_pdf = st.file_uploader(
        "Selecione ou arraste os arquivos PDF aqui",
        type="pdf",
        accept_multiple_files=True,
        label_visibility="collapsed"
    )

if arquivos_pdf:
    todos_os_dados = []
    with st.spinner('Processando arquivos... Por favor, aguarde.'):
        for arquivo in arquivos_pdf:
            dados = extrair_informacoes_pdf(arquivo)
            todos_os_dados.extend(dados)

    if todos_os_dados:
        st.success(f"🎉 Extração concluída com sucesso!")
        
        df = pd.DataFrame(todos_os_dados)
        df = df[['PRODUTO', 'QTD.', 'PÁGINA']]
        df.insert(0, 'ÍNDICE', range(1, 1 + len(df)))

        st.header("📊 Resultados da Extração")
        
        col1, col2 = st.columns((3, 1))

        with col1:
            st.dataframe(df, height=500, use_container_width=True,
                         column_config={
                             "ÍNDICE": st.column_config.NumberColumn(width="small"),
                             "PRODUTO": st.column_config.TextColumn(width="large"),
                             "QTD.": st.column_config.NumberColumn(width="small"),
                             "PÁGINA": st.column_config.TextColumn(width="small"),
                         })

   

        # --- SEÇÃO DE DOWNLOAD (MODIFICADA PARA EXCEL) ---
        @st.cache_data
        def converter_df_para_excel(df):
            output = io.BytesIO()
            # Usa o 'engine='openpyxl'' para escrever no formato .xlsx
            # 'index=False' para não incluir o índice do DataFrame no arquivo Excel
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Pedidos')
            # Pega os dados do buffer da memória
            processed_data = output.getvalue()
            return processed_data

        excel_data = converter_df_para_excel(df)
        
        st.download_button(
           label="💾 Baixar dados como Excel (.xlsx)",
           data=excel_data,
           file_name='produtos_extraidos.xlsx', # Nome do arquivo alterado
           # MIME type para arquivos Excel .xlsx
           mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
           use_container_width=True
        )
    else:
        st.warning("Nenhum dado de produto pôde ser extraído dos arquivos fornecidos. Verifique o formato dos PDFs.")
else:
    st.info("Aguardando o upload de arquivos para iniciar a extração.")
