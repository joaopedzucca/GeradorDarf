# --- APLICATIVO GERADOR DE DARF (V9 - MODO DE DEPURAÇÃO) ---

import streamlit as st
import pandas as pd
from pypdf import PdfReader, PdfWriter
import io
import re
import os
import shutil

# --- FUNÇÕES AUXILIARES COM MODO DE DEPURAÇÃO ---

import re
import pandas as pd

def parse_value_to_float(value):
    """
    Converte strings como:
      '1.234,56'  → 1234.56
      '1234,56'   → 1234.56
      '1,234.56'  → 1234.56
      'R$ 1.234'  → 1234.0
      '1 234,56'  → 1234.56
    """
    if pd.isna(value):
        return 0.0
    s = str(value).strip()
    # tira tudo que não for dígito, ponto ou vírgula
    s = re.sub(r"[^\d\.,\-]", "", s)
    # se tiver ponto e vírgula, assume ponto como milhares e vírgula decimal
    if s.count('.') > 0 and s.count(',') > 0:
        s = s.replace('.', '').replace(',', '.')
    # só vírgula → decimal
    elif s.count(',') > 0:
        s = s.replace(',', '.')
    # só ponto → já é decimal/padrão python
    try:
        return float(s)
    except ValueError:
        return 0.0

def format_value_for_pdf(value):
    """
    Gera string no formato brasileiro:
      1.234,56
      12,30
      0,00
    """
    numeric = parse_value_to_float(value)
    # formata com vírgula decimal e ponto milhar
    s = f"{numeric:,.2f}"       # ex: '1,234.56'
    s = s.replace(',', '#')     # '1#234.56'
    s = s.replace('.', ',')     # '1#234,56'
    s = s.replace('#', '.')     # '1.234,56'
    return s


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

# --- INTERFACE DO APLICATIVO ---

st.set_page_config(page_title="Gerador de DARF em Lote", layout="centered")
st.title("🚀 Gerador de DARF em Lote")
st.write("Esta ferramenta preenche múltiplos DARFs a partir de uma planilha Excel.")

DARF_TEMPLATE_FILENAME = "ModeloDarf.pdf"

if not os.path.exists(DARF_TEMPLATE_FILENAME):
    st.error(f"Erro Crítico: O arquivo modelo '{DARF_TEMPLATE_FILENAME}' não foi encontrado no repositório do aplicativo.")
    st.stop()

st.header("1. Faça o upload da sua planilha Excel")
uploaded_excel_file = st.file_uploader("Selecione a planilha com os dados dos DARFs", type=["xlsx"])

if uploaded_excel_file:
    if st.button("Gerar DARFs", type="primary", use_container_width=True):
        with st.spinner('Processando... Por favor, aguarde.'):
            try:
                # O restante do código permanece o mesmo
                field_map = {
                    'Nome/Telefone': 'Nome', 'Período de Apuração': 'Apuração', 'CNPJ': 'NI',
                    'Código da Receita': 'Receita', 'Data de vencimento': 'Vencimento',
                    'Valor do principal': 'Principal', 'Valor dos juros': 'Juros', 'Valor Total': 'Total'
                }
                df = pd.read_excel(uploaded_excel_file)

                with open(DARF_TEMPLATE_FILENAME, "rb") as f:
                    pdf_model_data = f.read()

                output_dir = 'darfs_preenchidos'
                if os.path.exists(output_dir): shutil.rmtree(output_dir)
                os.makedirs(output_dir)

                progress_bar = st.progress(0, text="Iniciando geração...")
                total_rows = len(df)

                for index, row in df.iterrows():
                    reader = PdfReader(io.BytesIO(pdf_model_data))
                    writer = PdfWriter(); writer.append(reader)
                    
                    data_to_fill = {
                        field_map['Nome/Telefone']: str(row.get('Nome/Telefone', '')),
                        field_map['Período de Apuração']: format_date(row.get('Período de Apuração')),
                        field_map['CNPJ']: format_cpf_cnpj(row.get('CNPJ')),
                        field_map['Código da Receita']: str(int(parse_value_to_float(row.get('Código da Receita', 0)))),
                        field_map['Data de vencimento']: format_date(row.get('Data de vencimento')),
                        field_map['Valor do principal']: format_value_for_pdf(row.get('Valor do principal')),
                        field_map['Valor dos juros']: format_value_for_pdf(row.get('Valor dos juros')),
                        field_map['Valor Total']: format_value_for_pdf(row.get('Valor Total'))
                    }
                    writer.update_page_form_field_values(writer.pages[0], data_to_fill)
                    
                    contribuinte_nome = re.sub(r'\W+', '_', str(row.get('Nome/Telefone', 'Contribuinte')))
                    periodo = format_date(row.get('Período de Apuração')).replace('/', '-')
                    output_filename = f"DARF_{index+1}_{contribuinte_nome}_{periodo}.pdf"
                    
                    with open(os.path.join(output_dir, output_filename), "wb") as output_stream:
                        writer.write(output_stream)
                    
                    progress_bar.progress((index + 1) / total_rows, text=f"Gerando DARF {index + 1}/{total_rows}")

                zip_filename = 'DARFs_Preenchidos'
                shutil.make_archive(zip_filename, 'zip', output_dir)
                
                st.success("🎉 Processamento concluído!")
                st.balloons()

                with open(f"{zip_filename}.zip", "rb") as fp:
                    st.download_button(
                        label="Clique aqui para baixar o ZIP com os DARFs",
                        data=fp,
                        file_name=f"{zip_filename}.zip",
                        mime="application/zip",
                        use_container_width=True
                    )
            except Exception as e:
                st.error(f"Ocorreu um erro: {e}")
                st.error("Verifique se os nomes das colunas na sua planilha Excel estão corretos e tente novamente.")
