# --- APLICATIVO GERADOR DE DARF (V12 - FORMATAÇÃO MANUAL DEFINITIVA) ---

import streamlit as st
import pandas as pd
from pypdf import PdfReader, PdfWriter
import io
import re
import os
import shutil

# --- FUNÇÕES AUXILIARES FINAIS E ROBUSTAS ---

def parse_value_to_float(value):
    """
    Função "à prova de balas" para converter qualquer formato de
    número (pt-br, en-us, com ou sem R$, etc.) para um float.
    """
    # Garante que temos uma string para trabalhar
    s = str(value).strip()

    if not s or s == 'nan':
        return 0.0
    
    # Passo 1: Limpa o valor, removendo tudo que não for dígito, vírgula, ponto ou sinal de menos.
    s_limpo = re.sub(r'[^\d.,-]', '', s)
    
    # Passo 2: Adivinha qual o separador decimal (o último que aparece)
    last_dot = s_limpo.rfind('.')
    last_comma = s_limpo.rfind(',')

    # Se a vírgula vem por último, assume formato PT-BR: "1.234,56"
    if last_comma > last_dot:
        # Remove os pontos de milhar e troca a vírgula decimal por ponto
        s_final = s_limpo.replace('.', '').replace(',', '.')
    # Se o ponto vem por último, assume formato EN-US: "1,234.56"
    elif last_dot > last_comma:
        # Remove as vírgulas de milhar
        s_final = s_limpo.replace(',', '')
    # Se não há separadores, ou só um tipo
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
    totalmente manual, garantindo que não haverá erros de formatação.
    """
    numeric_value = parse_value_to_float(value)

    # Converte para string com 2 casas decimais usando ponto (formato universal)
    s = f"{numeric_value:.2f}"

    # Separa a parte inteira da decimal
    partes = s.split('.')
    parte_inteira = partes[0]
    parte_decimal = partes[1]

    # Adiciona os pontos como separadores de milhar na parte inteira
    # Faz isso de trás para frente para garantir o agrupamento correto de 3 em 3
    parte_inteira_reversa = parte_inteira[::-1]
    # Ex: '0000052' -> ['000', '005', '2']
    chunks = [parte_inteira_reversa[i:i+3] for i in range(0, len(parte_inteira_reversa), 3)]
    # Junta com ponto: '000.005.2'
    parte_inteira_formatada_reversa = ".".join(chunks)
    # Reverte para a ordem correta: '2.500.000'
    parte_inteira_formatada = parte_inteira_formatada_reversa[::-1]

    # Junta a parte inteira formatada com a decimal, usando a vírgula
    return f"{parte_inteira_formatada},{parte_decimal}"


def format_cpf_cnpj(value):
    s = re.sub(r'\D', '', str(value))
    if len(s) == 11: return f"{s[:3]}.{s[3:6]}.{s[6:9]}-{s[9:]}"
    if len(s) == 14: return f"{s[:2]}.{s[2:5]}.{s[5:8]}/{s[8:12]}-{s[12:]}"
    return str(value)

def format_date(date_obj):
    if not date_obj or str(date_obj).strip() == '': return ""
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
    st.error(f"Erro Crítico: O arquivo modelo '{DARF_TEMPLATE_FILENAME}' não foi encontrado. Por favor, faça o upload dele para o repositório do GitHub junto com o 'app.py'.")
    st.stop()

st.header("1. Faça o upload da sua planilha Excel")
uploaded_excel_file = st.file_uploader("Selecione a planilha com os dados dos DARFs", type=["xlsx"])

if uploaded_excel_file:
    if st.button("Gerar DARFs", type="primary", use_container_width=True):
        with st.spinner('Processando... Por favor, aguarde.'):
            try:
                field_map = {
                    'Nome/Telefone': 'Nome', 'Período de Apuração': 'Apuração', 'CNPJ': 'NI',
                    'Código da Receita': 'Receita', 'Data de vencimento': 'Vencimento',
                    'Valor do principal': 'Principal', 'Valor dos juros': 'Juros', 'Valor Total': 'Total'
                }
                
                # Lê todas as colunas como texto para evitar problemas de formatação do Excel
                df = pd.read_excel(uploaded_excel_file, dtype=str)

                with open(DARF_TEMPLATE_FILENAME, "rb") as f:
                    pdf_model_data = f.read()

                output_dir = 'darfs_preenchidos'
                if os.
