# --- APLICATIVO GERADOR DE DARF (VERSÃO FINAL - DESENHO PRECISO) ---

import streamlit as st
import pandas as pd
from pypdf import PdfReader, PdfWriter
import io, re, os, shutil

# --- NOVAS BIBLIOTECAS PARA DESENHO ---
from reportlab.pdfgen import canvas
from reportlab.lib.units import pt
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# --- Funções Auxiliares (mantidas como no seu código) ---

def parse_value_to_float(value):
    s = str(value).strip()
    if not s or s.lower() == 'nan': return 0.0
    s2 = re.sub(r'[^\d.,-]', '', s)
    d = s2.rfind('.'); c = s2.rfind(',')
    if c > d: s3 = s2.replace('.', '').replace(',', '.')
    elif d > c: s3 = s2.replace(',', '')
    else: s3 = s2.replace(',', '.')
    try: return float(s3)
    except: return 0.0

def format_br(value):
    v = parse_value_to_float(value)
    s = f"{v:,.2f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")

def format_cpf_cnpj(x):
    s = re.sub(r'\D','',str(x))
    if len(s)==11: return f"{s[:3]}.{s[3:6]}.{s[6:9]}-{s[9:]}"
    if len(s)==14: return f"{s[:2]}.{s[2:5]}.{s[5:8]}/{s[8:12]}-{s[12:]}"
    return str(x)

def fmt_date(d):
    if not d or str(d).strip()=="": return ""
    try: return pd.to_datetime(d).strftime("%d/%m/%Y")
    except: return ""

# --- COORDENADAS EXTRAÍDAS DO SEU PDF ---
# [x_inferior_esquerdo, y_inferior_esquerdo, x_superior_direito, y_superior_direito]
COORDINATES = {
    'Nome': [45.35, 603.11, 303.30, 625.78],
    'Apuração': [425.85, 698.52, 561.91, 721.19],
    'NI': [425.85, 674.67, 561.91, 697.34],
    'Receita': [425.85, 650.82, 561.91, 673.49],
    'Referência': [425.85, 626.97, 561.91, 649.64],
    'Vencimento': [425.85, 603.12, 561.91, 625.79],
    'Principal': [425.85, 579.27, 561.91, 601.94],
    'Multa': [425.85, 555.42, 561.91, 578.09],
    'Juros': [425.85, 531.57, 561.91, 554.24],
    'Total': [425.85, 507.62, 561.91, 530.30],
}

# --- Interface do Aplicativo ---
st.set_page_config(page_title="Gerador de DARF por Desenho", layout="centered")
st.title("🎯 Gerador de DARF por Desenho de Precisão")
st.write("Esta ferramenta desenha os dados diretamente no PDF para um resultado perfeito.")

TEMPLATE = "ModeloDarf.pdf"
if not os.path.exists(TEMPLATE):
    st.error(f"Modelo '{TEMPLATE}' não encontrado."); st.stop()

# Registra a fonte Helvetica padrão para uso no ReportLab
pdfmetrics.registerFont(TTFont('Helv', 'Helvetica.ttf'))

u = st.file_uploader("📊 Planilha (.xlsx)", type="xlsx")
if not u: st.stop()

if st.button("Gerar DARFs por Desenho", use_container_width=True):
    with st.spinner("Desenhando os DARFs com precisão..."):
        try:
            df = pd.read_excel(u, dtype=str)
            df.columns = df.columns.str.strip()

            outdir = "darfs_desenhados"
            if os.path.exists(outdir): shutil.rmtree(outdir)
            os.makedirs(outdir)

            template_reader = PdfReader(TEMPLATE)
            template_page = template_reader.pages[0]

            prog = st.progress(0); total=len(df)

            for i, row in df.iterrows():
                packet = io.BytesIO()
                # Cria a "folha de rascunho" transparente (canvas)
                can = canvas.Canvas(packet, pagesize=template_page.mediabox.upper_right)
                can.setFont('Helv', 10) # Define a fonte e tamanho exatos do original

                # Dicionário de dados formatados
                data = {
                    "Nome": str(row.get("Nome/Telefone", "")),
                    "Apuração": fmt_date(row.get("Período de Apuração")),
                    "NI": format_cpf_cnpj(row.get("CNPJ")),
                    "Receita": str(int(parse_value_to_float(row.get("Código da Receita", 0)))),
                    "Vencimento": fmt_date(row.get("Data de vencimento")),
                    "Principal": format_br(row.get("Valor do principal", 0)),
                    "Juros": format_br(row.get("Valor dos juros", 0)),
                    "Total": format_br(row.get("Valor Total", 0)),
                    "Referência": str(row.get("Número de Referência", "")),
                    "Multa": format_br(row.get("Valor da multa",0))
                }
                
                # Desenha cada informação no canvas
                for field_name, text in data.items():
                    if field_name in COORDINATES:
                        coords = COORDINATES[field_name]
                        x0, y0, x1, y1 = coords
                        
                        # Calcula uma posição Y verticalmente centralizada
                        y_pos = y0 + 4 

                        if field_name == 'Nome': # Alinhamento à esquerda
                            can.drawString(x0 + 2, y_pos, text)
                        else: # Alinhamento à direita
                            can.drawRightString(x1 - 2, y_pos, text)

                can.save() # Salva o canvas
                packet.seek(0)
                
                # Lê o canvas que acabamos de criar
                overlay_reader = PdfReader(packet)
                overlay_page = overlay_reader.pages[0]

                # Estampa o canvas sobre a página do modelo
                template_page.merge_page(overlay_page)

                # Salva o resultado
                writer = PdfWriter()
                writer.add_page(template_page)
                
                nm = re.sub(r'\W+','_', str(row.get("Nome/Telefone","Contribuinte")))
                per = fmt_date(row.get("Período de Apuração","")).replace("/","-")
                fname = f"DARF_{i+1}_{nm}_{per}.pdf"
                with open(os.path.join(outdir, fname), "wb") as f:
                    writer.write(f)

                # Reseta a página do template para a próxima iteração
                template_page = PdfReader(TEMPLATE).pages[0]
                prog.progress((i+1)/total)

            zipf="DARFs_Gerados"
            shutil.make_archive(zipf,"zip",outdir)
            st.success("🎉 Pronto! DARFs desenhados com perfeição.")
            st.balloons()
            st.download_button("📥 Baixar ZIP com DARFs", open(f"{zipf}.zip","rb"),
                              file_name=f"{zipf}.zip", mime="application/zip",
                              use_container_width=True)

        except Exception as e:
            st.error(f"Ocorreu um erro inesperado: {e}")
            st.info("Dica: Verifique se a biblioteca 'reportlab' está instalada (`pip install reportlab`).")
