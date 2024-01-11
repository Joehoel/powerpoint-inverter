from datetime import datetime
import streamlit as st
import os
import zipfile
import shutil
import tempfile
from pptx_reader import read_pptx
from pptx_writer import save_presentation


def main():
    date = datetime.now().strftime('%Y-%m-%d')

    st.set_page_config(page_title='PowerPoint Inverter', page_icon='🔃')
    st.title('PowerPoint Inverter')
    st.write('Welcome to the PowerPoint Inverter! Please upload .pptx files or a .zip file containing .pptx files to invert.')

    st.sidebar.header("Options")
    options_form = st.sidebar.form(key="options")

    file_suffix = options_form.text_input('Suffix for inverted files', value='(inverted)',key='file_suffix')
    folder_name = options_form.text_input('Folder name for inverted files', value='Inverted Presentations', key='folder_name')
    folder_suffix = "" if options_form.selectbox('Suffix for folder name', options=[date, 'Nothing'], index=0, key='folder_suffix') == 'Nothing' else date

    submitted = options_form.form_submit_button("Apply")

    if submitted:
        options_form.success('Options applied successfully.')

    uploaded_files = st.file_uploader('Choose PowerPoint files (.pptx) or a .zip file', accept_multiple_files=True)
    
    if uploaded_files:
        st.success('Files uploaded successfully.')
        with tempfile.TemporaryDirectory() as tmpdirname:
            pp_filenames = []
            
            with st.spinner('Inverting presentations...'):
                for uploaded_file in uploaded_files:
                    if uploaded_file.name.endswith('.pptx'):
                        prs = read_pptx(uploaded_file)
                        output_file_path = save_presentation(prs, tmpdirname, os.path.splitext(uploaded_file.name)[0] + f' {file_suffix}.pptx')
                        pp_filenames.append(output_file_path)
                    elif uploaded_file.name.endswith('.zip'):
                        with zipfile.ZipFile(uploaded_file, 'r') as zip_ref:
                            zip_ref.extractall(tmpdirname)
                        for file_name in os.listdir(tmpdirname):
                            if file_name.endswith('.pptx'):
                                prs = read_pptx(os.path.join(tmpdirname, file_name))
                                output_file_path = save_presentation(prs, tmpdirname, os.path.splitext(file_name)[0] + f' {file_suffix}.pptx')
                                pp_filenames.append(output_file_path)
              
                zip_name = f'{folder_name}.zip' if folder_suffix == "" else f'{folder_name} - {folder_suffix}.zip'
                zip_path = os.path.join(tmpdirname, zip_name)
                with zipfile.ZipFile(zip_path, 'w') as zipf:
                    for file_name in pp_filenames:
                        zipf.write(file_name, os.path.basename(file_name))
                        os.unlink(file_name)
                with open(zip_path, 'rb') as file:
                    st.download_button(
                        label='Download',
                        data=file,
                        file_name=zip_name,
                        mime='application/zip'
                    )
                shutil.rmtree(tmpdirname)

if __name__ == '__main__':
    main()
