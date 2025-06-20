import streamlit as st
import pandas as pd
from pypdf import PdfReader, PdfWriter
from pypdf.generic import NameObject, BooleanObject, DictionaryObject, ArrayObject
import io, re, os, shutil

# --- helpers ---

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

def format_br(value):
    """Retorna string no formato 1.234,56"""
    v = parse_value_to_float(value)
    s = f"{v:,.2f}"           # ex: "2,52300.00" em en-US
    s = s.replace(",", "X").replace(".", ",").replace("X", ".")
    return s

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

# --- app ---

st.set_page_config(page_title="Gerador de DARF em Lote", layout="centered")
st.title("🚀 Gerador de DARF em Lote")
st.write("Preenche DARFs em lote a partir de Excel")

TEMPLATE = "ModeloDarf.pdf"
if not os.path.exists(TEMPLATE):
    st.error(f"Modelo '{TEMPLATE}' não encontrado."); st.stop()

u = st.file_uploader("📊 Planilha (.xlsx)", type="xlsx")
if not u: st.stop()

if st.button("Gerar DARFs", use_container_width=True):
    with st.spinner("Iniciando geração..."):
        try:
            df = pd.read_excel(u, dtype=str)
            df.columns = df.columns.str.strip()     # tira espaços finais/iniciais

            # mapeamento Excel → campos PDF
            M = {
              "Nome/Telefone":"Nome",
              "Período de Apuração":"Apuração",
              "CNPJ":"NI",
              "Código da Receita":"Receita",
              "Data de vencimento":"Vencimento",
              "Valor do principal":"Principal",
              "Valor dos juros":"Juros",
              "Valor Total":"Total",
            }

            # prepara saída
            outdir = "darfs"
            if os.path.exists(outdir): shutil.rmtree(outdir)
            os.makedirs(outdir)

            pdf_bytes = open(TEMPLATE,"rb").read()
            prog = st.progress(0); total=len(df)

            for i, row in df.iterrows():
                # carrega e clona
                reader = PdfReader(io.BytesIO(pdf_bytes))
                writer = PdfWriter(); writer.append(reader)

                # força aparências
                root = writer._root_object.get_object()
                ac = root.get(NameObject("/AcroForm"))
                if ac is None:
                    ac = DictionaryObject({
                        NameObject("/Fields"): ArrayObject(),
                        NameObject("/NeedAppearances"): BooleanObject(True)
                    })
                    ac_ref = writer._add_object(ac)
                    root[NameObject("/AcroForm")] = ac_ref
                else:
                    ac.get_object()[NameObject("/NeedAppearances")] = BooleanObject(True)

                # preenche usando strings formatadas
                data = {
                  M["Nome/Telefone"]: str(row.get("Nome/Telefone","")),
                  M["Período de Apuração"]: fmt_date(row.get("Período de Apuração")),
                  M["CNPJ"]: format_cpf_cnpj(row.get("CNPJ")),
                  M["Código da Receita"]: str(int(parse_value_to_float(row.get("Código da Receita",0)))),
                  M["Data de vencimento"]: fmt_date(row.get("Data de vencimento")),
                  M["Valor do principal"]: format_br(row.get("Valor do principal",0)),
                  M["Valor dos juros"]:    format_br(row.get("Valor dos juros",0)),
                  M["Valor Total"]:         format_br(row.get("Valor Total",0)),
                }
                # create_appearances=True garante que o PDF vai gravar a aparência
                writer.update_page_form_field_values(writer.pages[0], data, create_appearances=True)

                # salva
                nm = re.sub(r'\W+','_', row.get("Nome/Telefone","Contribuinte"))
                per = fmt_date(row.get("Período de Apuração","")).replace("/","-")
                fname = f"DARF_{i+1}_{nm}_{per}.pdf"
                with open(os.path.join(outdir,fname),"wb") as f:
                    writer.write(f)

                prog.progress((i+1)/total)

            # zip e download
            zipf="DARFs"
            shutil.make_archive(zipf,"zip",outdir)
            st.success("🎉 Pronto!")
            st.download_button("📥 Baixar ZIP", open(f"{zipf}.zip","rb"),
                               file_name=f"{zipf}.zip", mime="application/zip",
                               use_container_width=True)

        except Exception as e:
            st.error(f"Erro: {e}")
            st.error("Confira nomes de coluna sem espaços extras.")
