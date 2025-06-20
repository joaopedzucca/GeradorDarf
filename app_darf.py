# --- APLICATIVO GERADOR DE DARF (VERSÃO FINAL - ESTÁTICO PERFEITO) ---

import streamlit as st
import pandas as pd
from pypdf import PdfReader, PdfWriter
from pypdf.generic import NameObject, NumberObject
import io, re, os, shutil

# --- Funções Auxiliares (usando suas funções originais que são ótimas) ---

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
    """Retorna string no formato 1.234,56, que será 'desenhada' no PDF."""
    v = parse_value_to_float(value)
    s = f"{v:,.2f}"
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

# --- Interface do Aplicativo ---

st.set_page_config(page_title="Gerador de DARF Estático", layout="centered")
st.title("📄 Gerador de DARF 100% Estático e Formatado")
st.write("Cria DARFs com formatação e alinhamento perfeitos, travados para visualização universal.")

TEMPLATE = "ModeloDarf.pdf"
if not os.path.exists(TEMPLATE):
    st.error(f"Modelo '{TEMPLATE}' não encontrado."); st.stop()

u = st.file_uploader("📊 Planilha (.xlsx)", type="xlsx")
if not u: st.stop()

if st.button("Gerar DARFs Finais", use_container_width=True):
    with st.spinner("Formatando, alinhando e travando os PDFs..."):
        try:
            df = pd.read_excel(u, dtype=str)
            df.columns = df.columns.str.strip()

            M = {
                "Nome/Telefone":"Nome", "Período de Apuração":"Apuração",
                "CNPJ":"NI", "Código da Receita":"Receita",
                "Data de vencimento":"Vencimento", "Valor do principal":"Principal",
                "Valor dos juros":"Juros", "Valor Total":"Total",
            }
            
            campos_numericos = [M["Valor do principal"], M["Valor dos juros"], M["Valor Total"]]

            outdir = "darfs_finais_estaticos"
            if os.path.exists(outdir): shutil.rmtree(outdir)
            os.makedirs(outdir)

            template_bytes = open(TEMPLATE,"rb").read()
            prog = st.progress(0, text="Iniciando..."); total=len(df)

            for i, row in df.iterrows():
                prog.progress((i+1)/total, text=f"Processando DARF {i+1}/{total}...")
                
                # ETAPA 1: Preencher e formatar em memória
                reader_template = PdfReader(io.BytesIO(template_bytes))
                writer_filled = PdfWriter()
                writer_filled.append(reader_template)

                for page in writer_filled.pages:
                    if "/Annots" in page:
                        for annot in page["/Annots"]:
                            field = annot.get_object()
                            if field.get("/T") in campos_numericos:
                                field.update({ NameObject("/Q"): NumberObject(2) })

                data_to_fill = {
                    M["Nome/Telefone"]: str(row.get("Nome/Telefone","")),
                    M["Período de Apuração"]: fmt_date(row.get("Período de Apuração")),
                    M["CNPJ"]: format_cpf_cnpj(row.get("CNPJ")),
                    M["Código da Receita"]: str(int(parse_value_to_float(row.get("Código da Receita",0)))),
                    M["Data de vencimento"]: fmt_date(row.get("Data de vencimento")),
                    M["Valor do principal"]: format_br(row.get("Valor do principal",0)),
                    M["Valor dos juros"]: format_br(row.get("Valor dos juros",0)),
                    M["Valor Total"]: format_br(row.get("Valor Total",0)),
                }
                writer_filled.update_page_form_field_values(writer_filled.pages[0], data_to_fill)
                
                filled_buffer = io.BytesIO()
                writer_filled.write(filled_buffer)
                filled_buffer.seek(0)

                # ETAPA 2: Achatamento por "Estampa" (merge)
                reader_background = PdfReader(io.BytesIO(template_bytes))
                page_background = reader_background.pages[0]
                reader_foreground = PdfReader(filled_buffer)
                page_foreground = reader_foreground.pages[0]
                page_background.merge_page(page_foreground)
                
                # ETAPA 3: Salvar o resultado final 100% estático
                writer_final = PdfWriter()
                writer_final.add_page(page_background)
                
                nm = re.sub(r'\W+','_', row.get("Nome/Telefone","Contribuinte"))
                per = fmt_date(row.get("Período de Apuração","")).replace("/","-")
                fname = f"DARF_{i+1}_{nm}_{per}.pdf"
                with open(os.path.join(outdir,fname),"wb") as f:
                    writer_final.write(f)

            zipf="DARFs_Gerados"
            shutil.make_archive(zipf,"zip",outdir)
            st.success("🎉 Pronto! DARFs estáticos e perfeitos gerados.")
            st.balloons()
            st.download_button("📥 Baixar ZIP com DARFs", open(f"{zipf}.zip","rb"),
                              file_name=f"{zipf}.zip", mime="application/zip",
                              use_container_width=True)

        except Exception as e:
            st.error(f"Ocorreu um erro inesperado: {e}")
