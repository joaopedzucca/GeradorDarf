import streamlit as st
import pandas as pd
from pypdf import PdfReader, PdfWriter
from pypdf.generic import NameObject, BooleanObject, DictionaryObject
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
    s = str(value).strip()
    if not s or s.lower() == 'nan':
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
        return float(s_final) if s_final else 0.0
    except (ValueError, TypeError):
        return 0.0

def format_value_for_pdf(value):
    """
    (Ainda disponível, mas não usado no preenchimento automático
    de campos numéricos; deixei aqui caso queira usar em textos livres.)
    Formata um número para o padrão brasileiro (ex: 250.000,00).
    """
    numeric_value = parse_value_to_float(value)
    s = f"{numeric_value:.2f}"
    partes = s.split('.')
    int_rev = partes[0][::-1]
    chunks = [int_rev[i:i+3] for i in range(0, len(int_rev), 3)]
    int_fmt = ".".join(chunks)[::-1]
    return f"{int_fmt},{partes[1]}"

def format_cpf_cnpj(value):
    s = re.sub(r'\D', '', str(value))
    if len(s) == 11:
        return f"{s[:3]}.{s[3:6]}.{s[6:9]}-{s[9:]}"
    if len(s) == 14:
        return f"{s[:2]}.{s[2:5]}.{s[5:8]}/{s[8:12]}-{s[12:]}"
    return str(value)

def format_date(date_obj):
    if not date_obj or str(date_obj).strip() == '':
        return ""
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

if uploaded_excel_file and st.button("Gerar DARFs", type="primary", use_container_width=True):
    with st.spinner('Processando... Por favor, aguarde.'):
        try:
            field_map = {
                'Nome/Telefone': 'Nome',
                'Período de Apuração': 'Apuração',
                'CNPJ': 'NI',
                'Código da Receita': 'Receita',
                'Data de vencimento': 'Vencimento',
                'Valor do principal': 'Principal',
                'Valor dos juros': 'Juros',
                'Valor Total': 'Total'
            }
            df = pd.read_excel(uploaded_excel_file, dtype=str)
            with open(DARF_TEMPLATE_FILENAME, "rb") as f:
                pdf_model_data = f.read()

            output_dir = 'darfs_preenchidos'
            if os.path.exists(output_dir):
                shutil.rmtree(output_dir)
            os.makedirs(output_dir)

            progress_bar = st.progress(0, text="Iniciando geração...")
            total_rows = len(df)

            for index, row in df.iterrows():
                reader = PdfReader(io.BytesIO(pdf_model_data))
                writer = PdfWriter()
                writer.append(reader)

                # -- Forçar recalcular aparências no PDF --
                writer._root_object.setdefault(
                    NameObject("/AcroForm"),
                    DictionaryObject()
                )[NameObject("/NeedAppearances")] = BooleanObject(True)

                # -- Preenche campos com tipos corretos --
                data_to_fill = {
                    field_map['Nome/Telefone']: str(row.get('Nome/Telefone', '')),
                    field_map['Período de Apuração']: format_date(row.get('Período de Apuração')),
                    field_map['CNPJ']: format_cpf_cnpj(row.get('CNPJ')),
                    field_map['Código da Receita']: str(int(parse_value_to_float(row.get('Código da Receita', 0)))),
                    field_map['Data de vencimento']: format_date(row.get('Data de vencimento')),

                    # valores numéricos puros (float) – centavos preservados
                    field_map['Valor do principal']: parse_value_to_float(row.get('Valor do principal ', 0)),
                    field_map['Valor dos juros']:    parse_value_to_float(row.get('Valor dos juros ',   0)),
                    field_map['Valor Total']:         parse_value_to_float(row.get('Valor Total ',       0)),
                }
                writer.update_page_form_field_values(writer.pages[0], data_to_fill)

                contribuinte_nome = re.sub(r'\W+', '_', str(row.get('Nome/Telefone', 'Contribuinte')))
                periodo = format_date(row.get('Período de Apuração')).replace('/', '-')
                output_filename = f"DARF_{index+1}_{contribuinte_nome}_{periodo}.pdf"
                with open(os.path.join(output_dir, output_filename), "wb") as output_stream:
                    writer.write(output_stream)

                progress_bar.progress((index + 1) / total_rows,
                                      text=f"Gerando DARF {index + 1}/{total_rows}")

            # Empacota em ZIP
            zip_filename = 'DARFs_Preenchidos'
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
            st.error("Dica: Verifique se os nomes das colunas na sua planilha Excel estão exatamente como o esperado (incluindo possíveis espaços).")
