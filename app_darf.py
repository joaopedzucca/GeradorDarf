# --- APLICATIVO GERADOR DE DARF (V13 - ACHATAMENTO E ROBUSTEZ) ---

import streamlit as st
import pandas as pd
from pypdf import PdfReader, PdfWriter
import io
import re
import os
import shutil

# --- FUNÇÕES AUXILIARES (sem alterações, pois são robustas para seu propósito) ---

def parse_value_to_float(value):
    """
    Função "à prova de balas" para converter qualquer formato de
    número (pt-br, en-us, com ou sem R$, etc.) para um float.
    """
    s = str(value).strip()
    if not s or s == 'nan':
        return 0.0
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
    """
    Formata um número para o padrão brasileiro (ex: 250.000,00) de forma
    totalmente manual. Essencial para o processo de "flattening".
    """
    numeric_value = parse_value_to_float(value)
    s = f"{numeric_value:.2f}"
    partes = s.split('.')
    parte_inteira, parte_decimal = partes, partes
    parte_inteira_reversa = parte_inteira[::-1]
    chunks = [parte_inteira_reversa[i:i+3] for i in range(0, len(parte_inteira_reversa), 3)]
    parte_inteira_formatada_reversa = ".".join(chunks)
    parte_inteira_formatada = parte_inteira_formatada_reversa[::-1]
    return f"{parte_inteira_formatada},{parte_decimal}"

def format_cpf_cnpj(value):
    s = re.sub(r'\D', '', str(value))
    if len(s) == 11: return f"{s[:3]}.{s[3:6]}.{s[6:9]}-{s[9:]}"
    if len(s) == 14: return f"{s[:2]}.{s[2:5]}.{s[5:8]}/{s[8:12]}-{s[12:]}"
    return str(value)

def format_date(date_obj):
    if not date_obj or str(date_obj).strip() == '' or pd.isna(date_obj): return ""
    try:
        return pd.to_datetime(date_obj).strftime('%d/%m/%Y')
    except (ValueError, TypeError):
        return ""

# --- INTERFACE DO APLICATIVO ---

st.set_page_config(page_title="Gerador de DARF em Lote", layout="centered")
st.title("🚀 Gerador de DARF em Lote")
st.write("Esta ferramenta preenche múltiplos DARFs a partir de uma planilha Excel.")

DARF_TEMPLATE_FILENAME = "ModeloDarf.pdf"

if not os.path.exists(DARF_TEMPLATE_FILENAME):
    st.error(f"Erro Crítico: O arquivo modelo '{DARF_TEMPLATE_FILENAME}' não foi encontrado.")
    st.stop()

st.header("1. Faça o upload da sua planilha Excel")
uploaded_excel_file = st.file_uploader("Selecione a planilha com os dados dos DARFs", type=["xlsx"])

if uploaded_excel_file:
    if st.button("Gerar DARFs", type="primary", use_container_width=True):
        with st.spinner('Processando... Por favor, aguarde.'):
            try:
                # Mapeamento dos nomes de coluna da planilha para os nomes dos campos no PDF
                field_map = {
                    'Nome/Telefone': 'Nome', 'Período de Apuração': 'Apuração', 'CNPJ': 'NI',
                    'Código da Receita': 'Receita', 'Data de vencimento': 'Vencimento',
                    'Valor do principal': 'Principal', 'Valor dos juros': 'Juros', 'Valor Total': 'Total'
                }
                
                # Lê todas as colunas como texto para evitar problemas de formatação do Excel
                df = pd.read_excel(uploaded_excel_file, dtype=str)

                # >>> INÍCIO DA CORREÇÃO DE ROBUSTEZ (SANITIZAÇÃO) <<<
                # Remove espaços em branco do início e do fim de todos os nomes de colunas.
                # Isso torna o código imune a erros de digitação de espaços nos cabeçalhos do Excel.
                df.columns = df.columns.str.strip()
                # >>> FIM DA CORREÇÃO DE ROBUSTEZ <<<

                # Validação de colunas essenciais
                required_columns = list(field_map.keys())
                missing_columns = [col for col in required_columns if col not in df.columns]
                if missing_columns:
                    st.error(f"Erro na planilha: As seguintes colunas obrigatórias não foram encontradas: {', '.join(missing_columns)}. Por favor, corrija o arquivo Excel e tente novamente.")
                    st.stop()

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
                        field_map: str(row.get('Nome/Telefone', '')),
                        field_map['Período de Apuração']: format_date(row.get('Período de Apuração')),
                        field_map['CNPJ']: format_cpf_cnpj(row.get('CNPJ')),
                        field_map: str(int(parse_value_to_float(row.get('Código da Receita', 0)))),
                        field_map: format_date(row.get('Data de vencimento')),
                        # Agora usando os nomes de coluna sanitizados (sem espaços no final)
                        field_map['Valor do principal']: format_value_for_pdf(row.get('Valor do principal')),
                        field_map['Valor dos juros']: format_value_for_pdf(row.get('Valor dos juros')),
                        field_map: format_value_for_pdf(row.get('Valor Total'))
                    }
                    
                    # >>> INÍCIO DA MUDANÇA CRÍTICA (ACHATAMENTO) <<<
                    # O parâmetro 'flatten=True' é a solução definitiva.
                    # Ele "achata" o PDF, convertendo os campos de formulário preenchidos
                    # em conteúdo estático, eliminando qualquer erro de renderização
                    # relacionado a 'locale' ou a softwares de PDF.
                    writer.update_page_form_field_values(
                        writer.pages, 
                        data_to_fill,
                        flatten=True
                    )
                    # >>> FIM DA MUDANÇA CRÍTICA <<<
                    
                    contribuinte_nome = re.sub(r'\W+', '_', str(row.get('Nome/Telefone', 'Contribuinte')))
                    periodo = format_date(row.get('Período de Apuração')).replace('/', '-')
                    output_filename = f"DARF_{index+1}_{contribuinte_nome}_{periodo}.pdf"
                    
                    with open(os.path.join(output_dir, output_filename), "wb") as output_stream:
                        writer.write(output_stream)
                    
                    progress_bar.progress((index + 1) / total_rows, text=f"Gerando DARF {index + 1}/{total_rows}")

                zip_filename = 'DARFs_Preenchidos'
                shutil.make_archive(zip_filename, 'zip', output_dir)
                
                st.success("🎉 Todos os DARFs foram gerados e achatados com sucesso!")
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
                st.error(f"Ocorreu um erro inesperado durante o processamento: {e}")
                st.error("Dica: Verifique se os dados na planilha estão corretos e se os nomes das colunas correspondem ao esperado.")
