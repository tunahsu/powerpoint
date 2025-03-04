import os
import requests
from PIL import Image
from io import BytesIO
from together import Together
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.util import Inches
from pptx.slide import Slide

import logging
from typing import Literal, Union, List, Dict, Any

ChartTypes = Literal["bar", "line", "pie", "scatter", "area"]

class PresentationManager:
    def __init__(self):
        self.presentations: Dict[str, Any] = {}

    def add_picture_with_caption_slide(self, presentation_name: str, title: str,
                                       image_path: str, caption_text: str) -> Slide:
        """
        For the given presentation builds a slide with the picture with caption template.
        """
        try:
            prs = self.presentations[presentation_name]
        except KeyError as e:
            raise ValueError(f"Presentation '{presentation_name}' not found")
        slide_master = prs.slide_master

        # Add a new slide with layout 8 (Caption with Picture)
        slide_layout = prs.slide_layouts[8]
        slide = prs.slides.add_slide(slide_layout)

        # Set the title
        title_shape = slide.shapes.title
        title_shape.text = title

        # Insert the picture
        placeholder = slide.placeholders[1]
        picture = placeholder.insert_picture(image_path)

        # Set the caption
        caption = slide.placeholders[2]
        caption.text = caption_text
        return slide

    def add_title_with_content_slide(self, presentation_name: str, title: str, content: str) -> Slide:
        try:
            prs = self.presentations[presentation_name]
        except KeyError as e:
            raise ValueError(f"Presentation '{presentation_name}' not found")
        slide_master = prs.slide_master
        # Add a slide with title and content
        slide_layout = prs.slide_layouts[1]  # Use layout with title and content
        slide = prs.slides.add_slide(slide_layout)

        # Set the title
        title_shape = slide.shapes.title
        title_shape.text = title

        # Set the content
        content_shape = slide.placeholders[1]
        content_shape.text = content
        return slide

    def add_table_slide(self, presentation_name: str, title: str, headers: str, rows: str) -> Slide:

        try:
            prs = self.presentations[presentation_name]
        except KeyError as e:
            raise ValueError(f"Presentation '{presentation_name}' not found")

        slide_layout = prs.slide_layouts[5]  # Use layout with title only
        slide = prs.slides.add_slide(slide_layout)

        # Set the title
        title_shape = slide.shapes.title
        title_shape.text = title

        # Calculate table dimensions and position
        num_rows = len(rows) + 1  # +1 for header row
        num_cols = len(headers)

        # Position table in the middle of the slide with some margins
        x = Inches(1)  # Left margin
        y = Inches(2)  # Top margin below title

        # Make table width proportional to the number of columns
        width_per_col = Inches(8 / num_cols)  # Divide available width (8 inches) by number of columns
        height_per_row = Inches(0.4)  # Standard height per row

        # Create table
        shape = slide.shapes.add_table(
            num_rows,
            num_cols,
            x,
            y,
            width_per_col * num_cols,
            height_per_row * num_rows
        )
        table = shape.table

        # Add headers
        for col_idx, header in enumerate(headers):
            cell = table.cell(0, col_idx)
            cell.text = str(header)
            # Style header row
            paragraph = cell.text_frame.paragraphs[0]
            paragraph.font.bold = True
            paragraph.font.size = Pt(11)



        # Add data rows
        for row_idx, row_data in enumerate(rows, start=1):
            for col_idx, cell_value in enumerate(row_data):
                cell = table.cell(row_idx, col_idx)
                cell.text = str(cell_value)
                # Style data cells
                paragraph = cell.text_frame.paragraphs[0]
                paragraph.font.size = Pt(10)

        return slide

    def add_title_slide(self, presentation_name: str, title: str) -> Slide:
        try:
            prs = self.presentations[presentation_name]
        except KeyError as e:
            raise ValueError(f"Presentation '{presentation_name}' not found")

        # Add a slide with title and content
        slide_layout = prs.slide_layouts[0]  # Use layout with title only
        slide = prs.slides.add_slide(slide_layout)

        # Set the title
        title_shape = slide.shapes.title
        title_shape.text = title
        return slide


    async def generate_and_save_image(self, prompt: str, output_path: str) -> str:
        """Generate an image using Together AI/Flux Model and save it to the specified path."""

        api_key = os.environ.get('TOGETHER_AI_KEY')
        if not api_key:
            raise ValueError("TOGETHER_AI_KEY environment variable not set.")

        client = Together(api_key=api_key)

        try:
            # Generate the image
            response = client.images.generate(
                prompt=prompt,
                width=1024,
                height=1024,
                steps=4,
                model="black-forest-labs/FLUX.1-schnell-Free",
                n=1,
            )
        except Exception as e:
            raise ValueError(f"Failed to generate image: {str(e)}")

        image_url = response.data[0].url

        # Download the image
        try:
            response = requests.get(image_url)
            if response.status_code != 200:
                raise ValueError(f"Failed to download generated image: HTTP {response.status_code}")
        except requests.RequestException as e:
            raise ValueError(f"Network error downloading image: {str(e)}")

        # Save the image
        try:
            image = Image.open(BytesIO(response.content))
            # Ensure the save directory exists
            try:
                os.makedirs(os.path.dirname(output_path), exist_ok=True)
            except OSError as e:
                raise ValueError(f"Failed to create a directory for image: str({e})")
            # Save the image
            image.save(output_path)
        except (IOError, OSError) as e:
            raise ValueError(f"Failed to save image to {output_path}: {str(e)}")

        return output_path