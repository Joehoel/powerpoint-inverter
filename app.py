import streamlit as st
import os
import zipfile
import shutil
import tempfile
from pptx_reader import read_pptx
from pptx_writer import save_presentation

def main():
    st.title('PowerPoint Inverter App')
    st.write('Welcome to the PowerPoint Inverter app! Please upload .pptx files or a .zip file containing .pptx files to invert.')

    uploaded_files = st.file_uploader('Choose PowerPoint files (.pptx) or a .zip file', accept_multiple_files=True)
    if uploaded_files:
        st.write('Files uploaded successfully.')
        with tempfile.TemporaryDirectory() as tmpdirname:
            pp_filenames = []
            for uploaded_file in uploaded_files:
                if uploaded_file.name.endswith('.pptx'):
                    prs = read_pptx(uploaded_file)
                    output_file_path = save_presentation(prs, tmpdirname, os.path.splitext(uploaded_file.name)[0] + '_inverted.pptx')
                    pp_filenames.append(output_file_path)
                elif uploaded_file.name.endswith('.zip'):
                    with zipfile.ZipFile(uploaded_file, 'r') as zip_ref:
                        zip_ref.extractall(tmpdirname)
                    for file_name in os.listdir(tmpdirname):
                        if file_name.endswith('.pptx'):
                            prs = read_pptx(os.path.join(tmpdirname, file_name))
                            output_file_path = save_presentation(prs, tmpdirname, os.path.splitext(file_name)[0] + '_inverted.pptx')
                            pp_filenames.append(output_file_path)
            zip_path = os.path.join(tmpdirname, 'inverted_presentations.zip')
            with zipfile.ZipFile(zip_path, 'w') as zipf:
                for file_name in pp_filenames:
                    zipf.write(file_name, os.path.basename(file_name))
                    os.unlink(file_name)
            with open(zip_path, 'rb') as file:
                st.download_button(
                    label='Download Inverted Presentations',
                    data=file,
                    file_name='inverted_presentations.zip',
                    mime='application/zip'
                )
            shutil.rmtree(tmpdirname)

if __name__ == '__main__':
    main()