import streamlit as st
import os
import tempfile
from pdf_reader import read_pdf_and_identify_model
from bling_processor import process_bling_pdf
from sgbr_processor import process_sgbr_pdf 

def main():
    st.title("Processador de Notas")
    st.write("Faça upload de um arquivo PDF para processar e baixar como Excel.")
    
    # Upload do PDF
    uploaded_file = st.file_uploader("Escolha um arquivo PDF", type="pdf")
    
    if uploaded_file is not None:
        # Salvar o arquivo temporariamente
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_file:
            tmp_file.write(uploaded_file.read())
            tmp_file_path = tmp_file.name
        
        # Identificar o modelo do PDF
        modelo, texto_completo = read_pdf_and_identify_model(tmp_file_path)
        
        st.write(f"Modelo identificado: {modelo}")

        # Processar apenas se for modelo Bling ou SGBr
        if modelo == 'Bling' or modelo == 'SGBr Sistemas':
            try:
                # Processar o PDF e gerar o Excel em memória
                excel_data = process_bling_pdf(texto_completo) if modelo == 'Bling' else process_sgbr_pdf(texto_completo)

                # Botão para baixar o Excel
                st.download_button(
                    label="Baixar Excel",
                    data=excel_data,
                    file_name="output.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
            except Exception as e:
                st.error(f"Erro ao processar o PDF: {str(e)}")
        else:
            st.error("Modelo de PDF não suportado. Atualmente, apenas PDFs do modelo Bling são processados.")
        
        # Remover arquivo temporário
        os.unlink(tmp_file_path)

if __name__ == "__main__":
    main()