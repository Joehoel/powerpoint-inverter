from pptx import Presentation
import os

def save_presentation(prs, output_dir, output_file_name):
    output_file_path = os.path.join(output_dir, output_file_name)
    prs.save(output_file_path)
    return output_file_path
