# --- APLICATIVO GERADOR DE DARF (STREAMLIT) ---

# 1. Importação dos módulos necessários
import streamlit as st
import pandas as pd
from pypdf import PdfReader, PdfWriter
import io
import re
import os
import shutil

# --- FUNÇÕES AUXILIARES DE FORMATAÇÃO (versão final, sem dependências) ---

def parse_value_to_float(value):
    if pd.isna(value): return 0.0
    s = str(value).strip()
    dots = s.count('.'); commas = s.count(',')
    if dots >= 1 and commas == 1: s = s.replace('.', '').replace(',', '.')
    elif commas >= 1: s = s.replace(',', '')
    try: return float(s)
    except (ValueError, TypeError): return 0.0

def format_value_for_pdf(value):
    numeric_value = parse_value_to_float(value)
    return f'{numeric_value:_.2f}'.replace('.', ',').replace('_', '.')

def format_cpf_cnpj(value):
    s = re.sub(r'\D', '', str(value))
    if len(s) == 11: return f"{s[:3]}.{s[3:6]}.{s[6:9]}-{s[9:]}"
    if len(s) == 14: return f"{s[:2]}.{s[2:5]}.{s[5:8]}/{s[8:12]}-{s[12:]}"
    return str(value)

def format_date(date_obj):
    if pd.notna(date_obj):
        try: return pd.to_datetime(date_obj).strftime('%d/%m/%Y')
        except (ValueError, TypeError): return ""
    return ""

# --- INTERFACE DO APLICATIVO STREAMLIT ---

st.set_page_config(page_title="Gerador de DARF em Lote", layout="centered")

st.title("🚀 Gerador de DARF em Lote")
st.write("Esta ferramenta preenche múltiplos DARFs a partir de uma planilha Excel.")

# 1. Upload dos arquivos
st.header("1. Faça o upload dos seus arquivos")
uploaded_excel_file = st.file_uploader("Selecione a planilha Excel", type=["xlsx"])
uploaded_pdf_template = st.file_uploader("Selecione o modelo de DARF (PDF preenchível)", type=["pdf"])

if uploaded_excel_file and uploaded_pdf_template:
    st.success("Arquivos carregados com sucesso!")

    # Botão para iniciar o processamento
    if st.button("Gerar DARFs", type="primary"):
        with st.spinner('Processando... Por favor, aguarde.'):
            try:
                # Mapeamento dos campos
                field_map = {
                    'Nome/Telefone': 'Nome', 'Período de Apuração': 'Apuração', 'CNPJ': 'NI',
                    'Código da Receita': 'Receita', 'Data de vencimento': 'Vencimento',
                    'Valor do principal': 'Principal', 'Valor dos juros': 'Juros', 'Valor Total': 'Total'
                }

                df = pd.read_excel(uploaded_excel_file)
                pdf_model_data = uploaded_pdf_template.getvalue()

                output_dir = 'darfs_preenchidos'
                if os.path.exists(output_dir): shutil.rmtree(output_dir)
                os.makedirs(output_dir)

                # Barra de progresso
                progress_bar = st.progress(0, text="Iniciando geração...")
                total_rows = len(df)

                for index, row in df.iterrows():
                    reader = PdfReader(io.BytesIO(pdf_model_data))
                    writer = PdfWriter()
                    writer.append(reader)

                    # Atenção aos nomes das colunas com espaços no final
                    data_to_fill = {
                        field_map['Nome/Telefone']: str(row.get('Nome/Telefone', '')),
                        field_map['Período de Apuração']: format_date(row.get('Período de Apuração')),
                        field_map['CNPJ']: format_cpf_cnpj(row.get('CNPJ')),
                        field_map['Código da Receita']: str(int(parse_value_to_float(row.get('Código da Receita', 0)))),
                        field_map['Data de vencimento']: format_date(row.get('Data de vencimento')),
                        field_map['Valor do principal']: format_value_for_pdf(row.get('Valor do principal ')),
                        field_map['Valor dos juros']: format_value_for_pdf(row.get('Valor dos juros ')),
                        field_map['Valor Total']: format_value_for_pdf(row.get('Valor Total '))
                    }

                    writer.update_page_form_field_values(writer.pages[0], data_to_fill)

                    contribuinte_nome = re.sub(r'\W+', '_', str(row.get('Nome/Telefone', 'Contribuinte')))
                    periodo = format_date(row.get('Período de Apuração')).replace('/', '-')
                    output_filename = f"DARF_{index+1}_{contribuinte_nome}_{periodo}.pdf"
                    
                    with open(os.path.join(output_dir, output_filename), "wb") as output_stream:
                        writer.write(output_stream)
                    
                    # Atualiza a barra de progresso
                    progress_bar.progress((index + 1) / total_rows, text=f"Gerando DARF {index + 1}/{total_rows}")

                # Compacta a pasta de saída
                zip_filename = 'DARFs_Preenchidos'
                shutil.make_archive(zip_filename, 'zip', output_dir)
                
                st.success("🎉 Todos os DARFs foram gerados com sucesso!")
                st.balloons()

                # Oferece o arquivo ZIP para download
                with open(f"{zip_filename}.zip", "rb") as fp:
                    st.download_button(
                        label="Clique aqui para baixar o ZIP com os DARFs",
                        data=fp,
                        file_name=f"{zip_filename}.zip",
                        mime="application/zip"
                    )

            except Exception as e:
                st.error(f"Ocorreu um erro durante o processamento: {e}")
                st.error("Verifique se os nomes das colunas na sua planilha Excel correspondem exatamente ao esperado (ex: 'Nome/Telefone', 'Período de Apuração', etc.)")