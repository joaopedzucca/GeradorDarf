# --- APLICATIVO GERADOR DE DARF (V16 - FLATTEN DIRETO) ---

import streamlit as st
import pandas as pd
from pypdf import PdfReader, PdfWriter
import io
import re
import os
import shutil

# --- FUNÇÕES AUXILIARES (sem alterações) ---

def get_safe_value(row, column_name, default=''):
    """
    Pega um valor de uma linha do DataFrame de forma segura,
    retornando o primeiro item se encontrar uma lista (colunas duplicadas).
    """
    value = row.get(column_name, default)
    if isinstance(value, pd.Series):
        return value.iloc[0] if not value.empty else default
    return value

def parse_value_to_float(value):
    s = str(value).strip()
    if not s or s == 'nan': return 0.0
    s_limpo = re.sub(r'[^\d.,-]', '', s)
    last_dot = s_limpo.rfind('.')
    last_comma = s_limpo.rfind(',')
    if last_comma > last_dot:
        s_final = s_limpo.replace('.', '').replace(',', '.')
    elif last_dot > last_comma:
        s_final = s_limpo.replace(',', '')
    else:
        s_final = s_limpo.replace(',', '.')
    try:
        if not s_final: return 0.0
        return float(s_final)
    except (ValueError, TypeError):
        return 0.0

def format_value_for_pdf(value):
    numeric_value = parse_value_to_float(value)
    s = f"{numeric_value:.2f}"
    partes = s.split('.')
    parte_inteira, parte_decimal = partes[0], partes[1]
    parte_inteira_reversa = parte_inteira[::-1]
    chunks = [parte_inteira_reversa[i:i+3] for i in range(0, len(parte_inteira_reversa), 3)]
    parte_inteira_formatada = ".".join(chunks)[::-1]
    return f"{parte_inteira_formatada},{parte_decimal}"

def format_cpf_cnpj(value):
    s = re.sub(r'\D', '', str(value))
    if len(s) == 11: return f"{s[:3]}.{s[3:6]}.{s[6:9]}-{s[9:]}"
    if len(s) == 14: return f"{s[:2]}.{s[2:5]}.{s[5:8]}/{s[8:12]}-{s[12:]}"
    return str(value)

def format_date(date_obj):
    if date_obj is None or str(date_obj).strip() == '' or pd.isna(date_obj): return ""
    try:
        return pd.to_datetime(date_obj).strftime('%d/%m/%Y')
    except (ValueError, TypeError):
        return ""

# --- INTERFACE DO APLICATIVO ---

st.set_page_config(page_title="Gerador de DARF em Lote", layout="centered")
st.title("📄 Gerador de DARF em Lote (Versão Estática)")
st.write("Esta ferramenta preenche e 'achata' múltiplos DARFs para garantir compatibilidade total.")

DARF_TEMPLATE_FILENAME = "ModeloDarf.pdf"

if not os.path.exists(DARF_TEMPLATE_FILENAME):
    st.error(f"Erro Crítico: O arquivo modelo '{DARF_TEMPLATE_FILENAME}' não foi encontrado.")
    st.stop()

st.header("1. Faça o upload da sua planilha Excel")
uploaded_excel_file = st.file_uploader("Selecione a planilha com os dados dos DARFs", type=["xlsx"])

if uploaded_excel_file:
    if st.button("Gerar DARFs Estáticos", type="primary", use_container_width=True):
        with st.spinner('Processando e achatando os PDFs... Por favor, aguarde.'):
            try:
                field_map = {
                    'Nome/Telefone': 'Nome', 'Período de Apuração': 'Apuração', 'CNPJ': 'NI',
                    'Código da Receita': 'Receita', 'Data de vencimento': 'Vencimento',
                    'Valor do principal': 'Principal', 'Valor dos juros': 'Juros', 'Valor Total': 'Total'
                }

                df = pd.read_excel(uploaded_excel_file, dtype=str)

                with open(DARF_TEMPLATE_FILENAME, "rb") as f:
                    pdf_model_data = f.read()

                output_dir = 'darfs_preenchidos'
                if os.path.exists(output_dir): shutil.rmtree(output_dir)
                os.makedirs(output_dir)

                progress_bar = st.progress(0, text="Iniciando geração...")
                total_rows = len(df)

                for index, row in df.iterrows():
                    reader = PdfReader(io.BytesIO(pdf_model_data))
                    writer = PdfWriter()
                    writer.append(reader)

                    data_to_fill = {
                        field_map['Nome/Telefone']: str(get_safe_value(row, 'Nome/Telefone')),
                        field_map['Período de Apuração']: format_date(get_safe_value(row, 'Período de Apuração')),
                        field_map['CNPJ']: format_cpf_cnpj(get_safe_value(row, 'CNPJ')),
                        field_map['Código da Receita']: str(int(parse_value_to_float(get_safe_value(row, 'Código da Receita', 0)))),
                        field_map['Data de vencimento']: format_date(get_safe_value(row, 'Data de vencimento')),
                        field_map['Valor do principal']: format_value_for_pdf(get_safe_value(row, 'Valor do principal ')),
                        field_map['Valor dos juros']: format_value_for_pdf(get_safe_value(row, 'Valor dos juros ')),
                        field_map['Valor Total']: format_value_for_pdf(get_safe_value(row, 'Valor Total '))
                    }
                    
                    # --- A MÁGICA ACONTECE AQUI ---
                    # Preenche os campos e os "achata" em uma única operação.
                    writer.update_page_form_field_values(
                        writer.pages[0], data_to_fill, flatten=True
                    )
                    
                    contribuinte_nome = re.sub(r'\W+', '_', str(get_safe_value(row, 'Nome/Telefone', 'Contribuinte')))
                    periodo = format_date(get_safe_value(row, 'Período de Apuração')).replace('/', '-')
                    output_filename = f"DARF_{index+1}_{contribuinte_nome}_{periodo}.pdf"
                    
                    with open(os.path.join(output_dir, output_filename), "wb") as output_stream:
                        writer.write(output_stream)

                    progress_bar.progress((index + 1) / total_rows, text=f"Gerando DARF {index + 1}/{total_rows}")

                zip_filename = 'DARFs_Preenchidos_Estaticos'
                shutil.make_archive(zip_filename, 'zip', output_dir)

                st.success("🎉 Todos os DARFs foram gerados com sucesso!")
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
                st.error(f"Ocorreu um erro inesperado: {e}")
                st.error("Dica: Verifique se os nomes das colunas na sua planilha Excel estão exatamente como o esperado.")
