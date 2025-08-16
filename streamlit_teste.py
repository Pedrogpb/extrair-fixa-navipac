import streamlit as st
from PIL import Image
import openpyxl
from openpyxl.styles import Font, Alignment
import re
from io import BytesIO
import zipfile

#----------------------------------------------#
### FUNÇÕES ###
#----------------------------------------------#
def limpar_e_converter(valor, corte):
    """Limpa e converte um valor de linha para um número inteiro."""
    if len(valor) > corte:
        texto = valor[corte:]
    else:
        texto = ""
    texto = texto.split(',')[0].split('.')[0].strip()
    try:
        return int(texto)
    except ValueError:
        return 0

#--------------------------------------------------------------#
# SIDEBAR STREAMLIT
#--------------------------------------------------------------#
st.set_page_config(page_title="Extrair Fixas- Navipac", page_icon="📉", layout="centered")

st.sidebar.markdown("# Survey Info")
st.sidebar.markdown("### ")
imagem = Image.open("lh2_foto.jpg")
st.sidebar.image(imagem, width=250)
st.sidebar.markdown("""---""")

st.sidebar.markdown("""
<style>
    .sidebar-text {
        text-align: center;
        font-size: 10px;
    }
</style>
""", unsafe_allow_html=True)
st.sidebar.markdown('<p class="sidebar-text">Powered by Pedro Garcia.</p>', unsafe_allow_html=True)

#----------------------------------------------#
### PÁGINA STREAMLIT ###
#----------------------------------------------#
st.title("Extrair Fixas - NAVIPAC 🚢")
st.markdown("")
st.markdown("")

st.markdown("### 1. Envie a pasta com as fixas:")
arquivo_zip = st.file_uploader("", type="zip")

st.markdown("---")
st.markdown("### 2. Extrair fixas formato NPC:")
if st.button("Extrair"):
    if not arquivo_zip:
        st.error("Informe um arquivo .zip válido.")
    else:
        # --- CRIA EXCEL EM MEMÓRIA ---
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Dados NPC"

        fonte_padrao = Font(name='Arial', size=10)
        fonte_negrito = Font(name='Arial', size=10, bold=True)
        alinhamento_centralizado = Alignment(horizontal='center')

        ws.append(["Nome do Arquivo", "NORTE", "ESTE", "PROF"])
        for col in range(1, 5):
            cell = ws.cell(row=1, column=col)
            cell.alignment = alinhamento_centralizado
            cell.font = fonte_negrito

        # --- LÊ ARQUIVOS DO ZIP ---
        with zipfile.ZipFile(arquivo_zip) as z:
            arquivos_npc = [f for f in z.namelist() if f.lower().endswith('.npc')]

            for nome_arquivo in arquivos_npc:
                conteudo = None
                for encoding in ['utf-8', 'latin-1', 'windows-1252']:
                    try:
                        with z.open(nome_arquivo) as arquivo:
                            conteudo = [line.decode(encoding).strip() for line in arquivo.readlines()]
                            break
                    except:
                        continue

                if conteudo is None:
                    continue

                linha_11 = conteudo[10] if len(conteudo) > 10 else ""
                linha_12 = conteudo[11] if len(conteudo) > 11 else ""
                linha_13 = conteudo[12] if len(conteudo) > 12 else ""

                este = limpar_e_converter(linha_11, 57)
                norte = limpar_e_converter(linha_12, 56)
                prof = limpar_e_converter(linha_13, 61)
                nome_arquivo_sem_extensao = nome_arquivo.split('/')[-1].replace('.npc','')

                ws.append([nome_arquivo_sem_extensao, norte, este, prof])

        for col in ws.columns:
            max_length = 0
            col_letter = col[0].column_letter
            for cell in col:
                if cell.value:
                    cell.font = fonte_padrao
                    max_length = max(max_length, len(str(cell.value)))
            ws.column_dimensions[col_letter].width = max_length + 2

        # --- SALVA NA MEMÓRIA PARA DOWNLOAD ---
        output = BytesIO()
        wb.save(output)
        output.seek(0)

        st.success("Dados extraídos com sucesso!")
        st.markdown("---")
        st.markdown("### 3. Download tabela fixas:")
        st.download_button(
            label="📥 Baixar Excel",
            data=output,
            file_name="Tabela_fixas.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
