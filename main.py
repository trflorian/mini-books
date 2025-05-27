from docx import Document
from docx.shared import Mm
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.table import _Cell
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


def set_vertical_cell_direction(cell: _Cell, direction: str):
    """
    direction: "tbRl" for top-to-bottom, or "btLr" for bottom-to-top
    """
    assert direction in ("tbRl", "btLr")
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    text_direction = OxmlElement("w:textDirection")
    text_direction.set(qn("w:val"), direction)
    tcPr.append(text_direction)


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
cell_orientations = 4*[["tbRl", "btLr"]]

for row_id, row in enumerate(table.rows):
    row.height = Mm(72)
    for col_id, cell in enumerate(row.cells):
        cell_id = cell_indices[row_id][col_id]
        cell_orientation = cell_orientations[row_id][col_id]

        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

        cell_text = f"{cell_id}"#f"This is a text cell in row {row_id + 1}, column {col_id + 1}. The mini-book page number is {cell_id}."

        cell.text = cell_text
        set_vertical_cell_direction(cell, cell_orientation)

        cell.paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

document.add_page_break()

document.save("demo.docx")
