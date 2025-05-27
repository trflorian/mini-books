from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import Mm

from utils import CellDirection, set_vertical_cell_direction

document = Document()

section = document.sections[0]
section.orientation = WD_ORIENT.PORTRAIT
section.page_width = Mm(210)
section.page_height = Mm(297)
section.left_margin = Mm(4.0)
section.right_margin = Mm(4.0)
section.top_margin = Mm(4.0)
section.bottom_margin = Mm(0.0)

table = document.add_table(rows=4, cols=2)

cell_indices = [[2, 1], [3, 8], [4, 7], [5, 6]]
cell_orientations = 4 * [[CellDirection.TB_RL, CellDirection.BT_LR]]

for row_id, row in enumerate(table.rows):
    row.height = Mm(72)
    for col_id, cell in enumerate(row.cells):
        cell_id = cell_indices[row_id][col_id]
        cell_orientation = cell_orientations[row_id][col_id]

        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

        cell_text = f"{cell_id}"

        cell.text = cell_text
        set_vertical_cell_direction(cell, cell_orientation)

        cell.paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

document.add_page_break()

document.save("demo.docx")
