import streamlit as st
import pandas as pd
from parser import extraer_observaciones
import io

st.set_page_config(page_title="Extractor de Observaciones", layout="wide")
st.title("📄 Extractor de Observaciones desde Word")

uploaded_file = st.file_uploader("Sube el archivo Word (.docx)", type="docx")

if uploaded_file:
    with st.spinner("Procesando..."):
        data = extraer_observaciones(uploaded_file)
        df = pd.DataFrame(data)

        st.success("¡Procesado correctamente!")
        st.dataframe(df)

        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name="Observaciones")
        
        st.download_button(
            label="📥 Descargar Excel",
            data=buffer.getvalue(),
            file_name="observaciones_extraidas.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
