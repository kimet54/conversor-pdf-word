import streamlit as st
import pdfplumber
from docx import Document
import io

# Configuració bàsica de la pàgina
st.set_page_config(page_title="PDF a Word", page_icon="📄", layout="centered")

# Títol i descripció de la utilitat web
st.title("Conversor de PDF a Word 🪄")
st.write("Puja un arxiu PDF per convertir-lo a format Word (.docx). Utilitza les opcions per intentar eliminar els encapçalaments i els números de pàgina.")

# Giny per pujar l'arxiu
uploaded_file = st.file_uploader("Selecciona un arxiu PDF", type="pdf")

# Opcions d'eliminació en dues columnes
col1, col2 = st.columns(2)
with col1:
    remove_header = st.checkbox("Elimina encapçalaments (10% superior)", value=True)
with col2:
    remove_footer = st.checkbox("Elimina números de pàgina (10% inferior)", value=True)

# Lògica de conversió quan es puja un arxiu
if uploaded_file is not None:
    if st.button("Converteix a Word", type="primary"):
        
        # Mostrem un missatge de càrrega mentre treballa
        with st.spinner("Processant el document... Això pot trigar una mica depenent de la mida del PDF."):
            try:
                # Creem un document Word buit
                doc = Document()
                
                # Obrim el PDF amb pdfplumber
                with pdfplumber.open(uploaded_file) as pdf:
                    
                    # Iterem per cada pàgina del PDF
                    for i, page in enumerate(pdf.pages):
                        width = page.width
                        height = page.height
                        
                        # Calculem on tallem si les caselles estan marcades
                        top_margin = 0.1 * height if remove_header else 0
                        bottom_margin = 0.9 * height if remove_footer else height
                        
                        # Definim la caixa de tall (Bounding Box): (esquerra, dalt, dreta, baix)
                        bbox = (0, top_margin, width, bottom_margin)
                        
                        # Retallem la pàgina i n'extraiem el text
                        cropped_page = page.within_bbox(bbox)
                        text = cropped_page.extract_text()
                        
                        # Si hi ha text a la pàgina, l'afegim al Word
                        if text:
                            doc.add_paragraph(text)
                            # Opcional: Afegir un salt de pàgina al Word després de cada pàgina del PDF
                            # doc.add_page_break()

                # Guardem el document Word en la memòria temporal (BytesIO) per poder-lo descarregar
                word_buffer = io.BytesIO()
                doc.save(word_buffer)
                word_buffer.seek(0)
                
                st.success("Conversió completada amb èxit! 🎉")
                
                # Creem el botó de descàrrega
                st.download_button(
                    label="Descarrega l'arxiu Word",
                    data=word_buffer,
                    file_name=uploaded_file.name.replace(".pdf", ".docx"),
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
                
            except Exception as e:
                st.error(f"S'ha produït un error durant la conversió: {e}")
