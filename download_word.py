import streamlit as st
import base64

def download_word(file_path):
    try:
        with open(file_path, "rb") as file:
            file_data = file.read()
            b64 = base64.b64encode(file_data).decode()
            file_name = file_path.split("/")[-1]
            href = f'<a href="data:application/octet-stream;base64,{b64}" download="{file_name}">📥 Download {file_name}</a>'
            st.markdown(href, unsafe_allow_html=True)
    except FileNotFoundError:
        st.error("File not found. Please try again.")
    except Exception as e:
        st.error(f"Error downloading file: {e}")