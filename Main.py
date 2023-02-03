from bad_links_doc import detect_url_doc
from bad_links_ppt import detect_url_ppt
import streamlit as st
import os

st.title("Broken Links Detector for Word Documents and PPTs")

input_folder_path = st.text_input("Enter the input folder path:")
output_folder_path = st.text_input("Enter the output folder path:")
if input_folder_path and output_folder_path:
    if os.path.isdir(input_folder_path) and os.path.isdir(output_folder_path):
        try:
            detect_url_doc(input_folder_path, output_folder_path)
            detect_url_ppt(input_folder_path, output_folder_path)
            st.write("Broken links detected and saved to broken_links.csv in the output folder.")
        except Exception as e:
            st.write("An error occurred:", e)
    else:
        st.write("Invalid input or output folder path.")


