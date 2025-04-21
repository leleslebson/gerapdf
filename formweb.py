import streamlit as st
import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib.utils import ImageReader
from PIL import Image
from io import BytesIO
import requests
from datetime import datetime
import time

CAMPOS_EXIBIDOS = [
    'Número OS', 'Descrição', 'Área', 'Tag', 'Equipamento', 'Centro Trabalho', 'Solicitante',
    'Altura', 'Largura', 'Comprimento', 'Volume', 'Qtd Piso Extra', 'Executante', 'Data Execução', 'Observação'
]

LOGO_URL = "https://www.contrex.com.br/wp-content/themes/contrex_2019/images/logo.png"

def gerar_pdf_multipagina(df):
    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=letter)
    largura, altura = letter

    for i, linha in df.iterrows():
        margem = 40
        espacamento = 14
        y = altura - 40

        # LOGO
        try:
            response = requests.get(LOGO_URL, timeout=10)
            logo = Image.open(BytesIO(response.content)).resize((120, 40))
            logo_io = BytesIO()
            logo.save(logo_io, format='PNG')
            logo_io.seek(0)
            c.drawImage(ImageReader(logo_io), margem, y - 40, mask='auto')
        except Exception:
            pass
        y -= 20

        # TÍTULO
        c.setFont("Helvetica-Bold", 16)
        c.drawCentredString(largura / 2, y, "REGISTRO DE PONTO DE ANDAIME")
        y -= espacamento * 2

        # PLACA
        c.setFont("Helvetica-Bold", 16)
        placa = str(linha.get('Placa', ''))
        c.drawCentredString(largura / 2, y, f"{placa}")
        text_width = c.stringWidth(placa, "Helvetica-Bold", 16)
        c.line(largura / 2 - text_width / 2, y - 2, largura / 2 + text_width / 2, y - 2)
        y -= espacamento * 2

        # CAMPOS
        for campo in CAMPOS_EXIBIDOS:
            valor = linha.get(campo, '')
            if pd.notna(valor) and str(valor).strip() != "":
                c.setFont("Helvetica-Bold", 10)
                c.drawString(margem, y, f"{campo}:")
                c.setFont("Helvetica", 10)
                if campo == "Descrição":
                    texto = str(valor)
                    largura_max = 400
                    linhas = []
                    while texto:
                        for i in range(len(texto), 0, -1):
                            if c.stringWidth(texto[:i], "Helvetica", 10) <= largura_max:
                                linhas.append(texto[:i])
                                texto = texto[i:].lstrip()
                                break
                        else:
                            linhas.append(texto)
                            break
                    for linha_desc in linhas[:2]:
                        c.drawString(margem + 120, y, linha_desc)
                        y -= espacamento
                else:
                    c.drawString(margem + 120, y, str(valor))
                y -= espacamento

        # IMAGEM FINAL (Fotografias)
        foto_url = linha.get('Fotografias', '')
        if pd.notna(foto_url) and foto_url.startswith('http'):
            try:
                response = requests.get(foto_url, timeout=10)
                imagem = Image.open(BytesIO(response.content))
                imagem.thumbnail((400, 300))
                img_buffer = BytesIO()
                imagem.save(img_buffer, format='PNG')
                img_buffer.seek(0)
                c.drawImage(ImageReader(img_buffer), margem, y - 300, width=400, preserveAspectRatio=True, mask='auto')
            except Exception:
                pass

        c.showPage()
    c.save()
    buffer.seek(0)
    return buffer

# ---------------- STREAMLIT APP -----------------

st.set_page_config(page_title="Gerador de PDF", layout="centered")

st.title("📄 Gerador de PDF a partir de Excel")
st.markdown("Faça upload da planilha com os dados e clique em **Gerar PDF**.")

arquivo = st.file_uploader("Selecione a planilha Excel", type=["xlsx", "xls"])

if arquivo:
    try:
        df = pd.read_excel(arquivo)
        st.success(f"Arquivo carregado com sucesso! {len(df)} linhas encontradas.")
        if st.button("📤 Gerar PDF"):
            progresso = st.progress(0, text="Gerando PDF...")
            buffer = BytesIO()
            for i in range(101):
                time.sleep(0.01)
                progresso.progress(i, text=f"Gerando PDF... {i}%")
            buffer = gerar_pdf_multipagina(df)
            progresso.empty()
            st.success("PDF gerado com sucesso!")
            st.download_button("⬇️ Baixar PDF", buffer, file_name="relatorio.pdf", mime="application/pdf")
    except Exception as e:
        st.error(f"Erro ao processar o arquivo: {e}")



