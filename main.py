from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import Mm

from utils import CellDirection, set_vertical_cell_direction

# A4 Portrait Page Setup, all values in millimeters
PAGE_ORIENTATION = WD_ORIENT.PORTRAIT
PAGE_WIDTH = 210
PAGE_HEIGHT = 297
MARGIN = 4.0

document = Document()
page_section = document.sections[0]

# Page Setup
page_section.orientation = PAGE_ORIENTATION
page_section.page_width = Mm(PAGE_WIDTH)
page_section.page_height = Mm(PAGE_HEIGHT)

# Margins
page_section.left_margin = Mm(MARGIN)
page_section.right_margin = Mm(MARGIN)
page_section.top_margin = Mm(MARGIN)
page_section.bottom_margin = Mm(MARGIN)

# Table Setup
table = document.add_table(rows=4, cols=2)

cell_indices = [[2, 1], [3, 8], [4, 7], [5, 6]]
cell_orientations = 4 * [[CellDirection.TB_RL, CellDirection.BT_LR]]

for row_id, row in enumerate(table.rows):
    row.height = Mm(PAGE_HEIGHT / len(table.rows) - MARGIN)
    for col_id, cell in enumerate(row.cells):
        cell_id = cell_indices[row_id][col_id]
        cell_orientation = cell_orientations[row_id][col_id]

        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

        cell_text = f"{cell_id}"

        cell.text = cell_text
        set_vertical_cell_direction(cell, cell_orientation)

        cell.paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

document.save("demo.docx")
