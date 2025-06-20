import streamlit as st
import pandas as pd
from pypdf import PdfReader, PdfWriter
from pypdf.generic import (
    NameObject,
    BooleanObject,
    DictionaryObject,
    ArrayObject
)
import io
import re
import os
import shutil

# --- FUNÇÕES AUXILIARES ---

def parse_value_to_float(value):
    """
    Converte qualquer formato pt-BR ou en-US, com ou sem R$, em float.
    """
    s = str(value).strip()
    if not s or s.lower() == 'nan':
        return 0.0
    # remove tudo que não seja dígito, ponto, vírgula ou sinal
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
    except:
        return 0.0

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
    except:
        return ""

# --- STREAMLIT APP ---

st.set_page_config(page_title="Gerador de DARF em Lote", layout="centered")
st.title("🚀 Gerador de DARF em Lote")
st.write("Preenche múltiplos DARFs a partir de uma planilha Excel.")

DARF_TEMPLATE_FILENAME = "ModeloDarf.pdf"
if not os.path.exists(DARF_TEMPLATE_FILENAME):
    st.error(f"Arquivo modelo '{DARF_TEMPLATE_FILENAME}' não encontrado.")
    st.stop()

st.header("1. Faça o upload da sua planilha Excel")
uploaded = st.file_uploader("Selecione sua planilha (.xlsx)", type="xlsx")

if uploaded and st.button("Gerar DARFs", use_container_width=True):
    with st.spinner("Processando..."):
        try:
            # mapeamento Excel → PDF
            field_map = {
                'Nome/Telefone': 'Nome',
                'Período de Apuração': 'Apuração',
                'CNPJ': 'NI',
                'Código da Receita': 'Receita',
                'Data de vencimento': 'Vencimento',
                'Valor do principal': 'Principal',
                'Valor dos juros': 'Juros',
                'Valor Total': 'Total',
            }

            # lê tudo como string e remove espaços extras nos cabeçalhos
            df = pd.read_excel(uploaded, dtype=str)
            df.columns = df.columns.str.strip()

            with open(DARF_TEMPLATE_FILENAME, "rb") as f:
                pdf_bytes = f.read()

            # prepara pasta de saída
            out_dir = "darfs_preenchidos"
            if os.path.exists(out_dir):
                shutil.rmtree(out_dir)
            os.makedirs(out_dir)

            total = len(df)
            prog = st.progress(0, text="Iniciando geração...")

            for i, row in df.iterrows():
                reader = PdfReader(io.BytesIO(pdf_bytes))
                writer = PdfWriter()
                writer.append(reader)

                # === configura NeedAppearances corretamente ===
                root = writer._root_object.get_object()
                acro_ref = root.get(NameObject("/AcroForm"))
                if acro_ref is None:
                    # cria AcroForm se não existir
                    acro = DictionaryObject({
                        NameObject("/Fields"): ArrayObject(),
                        NameObject("/NeedAppearances"): BooleanObject(True)
                    })
                    acro_ref = writer._add_object(acro)
                    root[NameObject("/AcroForm")] = acro_ref
                else:
                    # obtém dicionário real e seta NeedAppearances
                    acro = acro_ref.get_object()
                    acro[NameObject("/NeedAppearances")] = BooleanObject(True)

                # === preenche campos ===
                data = {
                    field_map['Nome/Telefone']: str(row.get('Nome/Telefone', '')),
                    field_map['Período de Apuração']: format_date(row.get('Período de Apuração')),
                    field_map['CNPJ']: format_cpf_cnpj(row.get('CNPJ')),
                    field_map['Código da Receita']: str(int(parse_value_to_float(row.get('Código da Receita', 0)))),
                    field_map['Data de vencimento']: format_date(row.get('Data de vencimento')),
                    # valores numéricos puros para o PDF aplicar máscara
                    field_map['Valor do principal']: parse_value_to_float(row.get('Valor do principal', 0)),
                    field_map['Valor dos juros']:    parse_value_to_float(row.get('Valor dos juros',   0)),
                    field_map['Valor Total']:         parse_value_to_float(row.get('Valor Total',       0)),
                }
                writer.update_page_form_field_values(writer.pages[0], data)

                # salva PDF individual
                nome = re.sub(r'\W+', '_', str(row.get('Nome/Telefone', 'Contribuinte')))
                periodo = format_date(row.get('Período de Apuração')).replace("/", "-")
                out_path = os.path.join(out_dir, f"DARF_{i+1}_{nome}_{periodo}.pdf")
                with open(out_path, "wb") as out_f:
                    writer.write(out_f)

                prog.progress((i + 1) / total, text=f"Gerando DARF {i+1}/{total}")

            # empacota em ZIP
            zip_name = "DARFs_Preenchidos"
            shutil.make_archive(zip_name, "zip", out_dir)

            st.success("🎉 DARFs gerados com sucesso!")
            with open(f"{zip_name}.zip", "rb") as fp:
                st.download_button(
                    "Baixar ZIP com os DARFs",
                    data=fp,
                    file_name=f"{zip_name}.zip",
                    mime="application/zip",
                    use_container_width=True
                )

        except Exception as e:
            st.error(f"Ocorreu um erro inesperado: {e}")
            st.error("Verifique se os nomes das colunas na sua planilha estão corretos (sem espaços extras).")
