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


