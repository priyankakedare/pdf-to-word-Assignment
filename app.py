import streamlit as st
import tempfile
from generate_doc import generate_word

# ---------------- Page Config ----------------
st.set_page_config(
    page_title="PDF to Word Automation",
    page_icon="📄",
    layout="centered"
)

# ---------------- Custom Styling ----------------
st.markdown("""
<style>
.main {
    background-color: #f7f9fc;
}
.title {
    text-align: center;
    font-size: 34px;
    font-weight: 700;
    color: #1f2937;
}
.subtitle {
    text-align: center;
    font-size: 16px;
    color: #4b5563;
}
.card {
    background-color: white;
    padding: 25px;
    border-radius: 12px;
    box-shadow: 0px 4px 15px rgba(0,0,0,0.08);
}
.footer {
    text-align: center;
    font-size: 13px;
    color: #6b7280;
}
</style>
""", unsafe_allow_html=True)

# ---------------- Header ----------------
st.markdown('<div class="title">PDF to Word Converter</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="subtitle">Accurate legal document replication using Python</div>',
    unsafe_allow_html=True
)
st.write("")

# ---------------- Card UI ----------------
st.markdown('<div class="card">', unsafe_allow_html=True)

uploaded_file = st.file_uploader(
    "📤 Upload PDF file",
    type=["pdf"],
    help="Upload the Mediation Application Form PDF"
)

st.write("")

if uploaded_file:
    st.success("PDF uploaded successfully ✅")

    if st.button("✨ Generate Word Document"):
        with st.spinner("Recreating document layout..."):
            output_path = generate_word()

        with open(output_path, "rb") as f:
            st.download_button(
                label="⬇️ Download Word File",
                data=f,
                file_name="Mediation_Application_Form.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

st.markdown("</div>", unsafe_allow_html=True)

# ---------------- Footer ----------------
st.write("")
st.markdown(
    '<div class="footer">Built with Python & Streamlit • Document Automation Demo</div>',
    unsafe_allow_html=True
)
