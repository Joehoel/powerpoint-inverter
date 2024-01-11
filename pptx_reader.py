from pptx import Presentation
from color_inverter import invert_text_and_background
from image_inverter import invert_image_colors  # New import statement
import streamlit as st

def read_pptx(file_stream):
    # Load the presentation
    prs = Presentation(file_stream)

    # Iterate over each slide
    for slide in prs.slides:
        # Invert colors of the text and background
        invert_text_and_background(slide)
        
        # Invert colors of the images on each slide
        invert_image_colors(slide)  # Call the new function here

    return prs  # Return the presentation object for further processing
