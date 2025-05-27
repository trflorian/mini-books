from enum import Enum

from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.table import _Cell


class CellDirection(str, Enum):
    """
    Enum for cell text direction.
    """

    TB_RL = "tbRl"  # Top to Bottom, Right to Left
    BT_LR = "btLr"  # Bottom to Top, Left to Right

    def __str__(self) -> str:
        return self.value


def set_vertical_cell_direction(
    cell: _Cell,
    direction: CellDirection,
) -> None:
    """
    Set the vertical text direction for a table cell.
    Args:
        cell (_Cell): The table cell to modify.
        direction (CellDirection): The text direction to set, either TB_RL or BT_LR.
    """
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    text_direction = OxmlElement("w:textDirection")
    text_direction.set(qn("w:val"), str(direction))
    tcPr.append(text_direction)
