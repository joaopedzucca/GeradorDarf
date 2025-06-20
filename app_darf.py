# --- APLICATIVO GERADOR DE DARF (VERSÃO FINAL - INTERATIVA INTELIGENTE) ---

import streamlit as st
import pandas as pd
from pypdf import PdfReader, PdfWriter
import io, re, os, shutil

# --- Funções Auxiliares ---

def parse_value_to_float(value):
    s = str(value).strip()
    if not s or s.lower() == 'nan':
        return 0.0
    s2 = re.sub(r'[^\d.,-]', '', s)
    d = s2.rfind('.'); c = s2.rfind(',')
    if c > d:
        s3 = s2.replace('.', '').replace(',', '.')
    elif d > c:
        s3 = s2.replace(',', '')
    else:
        s3 = s2.replace(',', '.')
    try: return float(s3)
    except: return 0.0

def format_for_js(value):
    """
    NOVA FUNÇÃO: Formata o número como um texto simples com ponto decimal,
    que é o formato que o JavaScript do PDF espera receber. Ex: 1234.50
    """
    v = parse_value_to_float(value)
    return f"{v:.2f}"

def format_cpf_cnpj(x):
    s = re.sub(r'\D','',str(x))
    if len(s)==11: return f"{s[:3]}.{s[3:6]}.{s[6:9]}-{s[9:]}"
    if len(s)==14: return f"{s[:2]}.{s[2:5]}.{s[5:8]}/{s[8:12]}-{s[12:]}"
    return str(x)

def fmt_date(d):
    if not d or str(d).strip()=="":
        return ""
    try:
        return pd.to_datetime(d).strftime("%d/%m/%Y")
    except:
        return ""

# --- Interface do Aplicativo ---

st.set_page_config(page_title="Gerador de DARF Interativo", layout="centered")
st.title("📄 Gerador de DARF Interativo Inteligente")
st.write("Esta ferramenta cria DARFs interativos que usam a formatação nativa do modelo PDF.")

TEMPLATE = "ModeloDarf.pdf"
if not os.path.exists(TEMPLATE):
    st.error(f"Modelo '{TEMPLATE}' não encontrado."); st.stop()

u = st.file_uploader("📊 Planilha (.xlsx)", type="xlsx")
if not u: st.stop()

if st.button("Gerar DARFs Interativos", use_container_width=True):
    with st.spinner("Gerando DARFs..."):
        try:
            df = pd.read_excel(u, dtype=str)
            df.columns = df.columns.str.strip()

            M = {
                "Nome/Telefone":"Nome", "Período de Apuração":"Apuração",
                "CNPJ":"NI", "Código da Receita":"Receita",
                "Data de vencimento":"Vencimento", "Valor do principal":"Principal",
                "Valor dos juros":"Juros", "Valor Total":"Total",
            }

            outdir = "darfs_interativos"
            if os.path.exists(outdir): shutil.rmtree(outdir)
            os.makedirs(outdir)

            template_bytes = open(TEMPLATE,"rb").read()
            prog = st.progress(0, text="Iniciando..."); total=len(df)

            for i, row in df.iterrows():
                prog.progress((i+1)/total, text=f"Gerando DARF {i+1}/{total}...")
                
                reader_template = PdfReader(io.BytesIO(template_bytes))
                writer = PdfWriter()
                writer.append(reader_template)
                
                # Preenchemos os dados usando o formato que o PDF espera
                data_to_fill = {
                    M["Nome/Telefone"]: str(row.get("Nome/Telefone","")),
                    M["Período de Apuração"]: fmt_date(row.get("Período de Apuração")),
                    M["CNPJ"]: format_cpf_cnpj(row.get("CNPJ")),
                    M["Código da Receita"]: str(int(parse_value_to_float(row.get("Código da Receita",0)))),
                    M["Data de vencimento"]: fmt_date(row.get("Data de vencimento")),
                    # Para valores, usamos a nova função para entregar o dado "cru"
                    M["Valor do principal"]: format_for_js(row.get("Valor do principal",0)),
                    M["Valor dos juros"]: format_for_js(row.get("Valor dos juros",0)),
                    M["Valor Total"]: format_for_js(row.get("Valor Total",0)),
                }

                # Simplesmente preenchemos os valores. O PDF fará o resto.
                # Não usamos mais nenhum tipo de achatamento (flattening).
                writer.update_page_form_field_values(writer.pages[0], data_to_fill)
                
                nm = re.sub(r'\W+','_', row.get("Nome/Telefone","Contribuinte"))
                per = fmt_date(row.get("Período de Apuração","")).replace("/","-")
                fname = f"DARF_{i+1}_{nm}_{per}.pdf"
                with open(os.path.join(outdir,fname),"wb") as f:
                    writer.write(f)

            # zip e download
            zipf="DARFs_Gerados"
            shutil.make_archive(zipf,"zip",outdir)
            st.success("🎉 Pronto! DARFs interativos gerados com sucesso.")
            st.balloons()
            st.download_button("📥 Baixar ZIP com DARFs", open(f"{zipf}.zip","rb"),
                              file_name=f"{zipf}.zip", mime="application/zip",
                              use_container_width=True)

        except Exception as e:
            st.error(f"Ocorreu um erro inesperado: {e}")
            st.error("Dica: Verifique se os nomes das colunas na planilha correspondem exatamente ao esperado.")
