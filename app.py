import streamlit as st

st.set_page_config(
    page_title="HYDAC Lead Processor",
    layout="wide"
)

st.title("HYDAC Lead Processor")
st.success("App is running successfully.")

st.write("Step 1 completed: Streamlit deployment is working.")

uploaded_file = st.file_uploader("Upload MSG file", type=["msg"])

if uploaded_file:
    st.info(f"File uploaded: {uploaded_file.name}")
    st.write("Next step will be MSG parsing after deployment is confirmed.")
