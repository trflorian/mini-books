import contextlib
from enum import Enum
from collections.abc import Generator
import tempfile

import cv2

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


@contextlib.contextmanager
def create_image_for_direction(
    image_path: str,
    direction: CellDirection,
) -> Generator[str]:
    """
    Create a temporary image file with the specified direction.
    Args:
        image_path (str): The path to the original image.
        direction (CellDirection): The text direction for the image.
    Yields:
        str: The path to the temporary image file.
    """

    with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp_file:
        tmp_filename = tmp_file.name
        
        # Read the original image
        image = cv2.imread(image_path)
        if image is None:
            raise ValueError(f"Could not read image from {image_path}")
        
        # Rotate the image based on the direction
        if direction == CellDirection.TB_RL:
            rotated_image = cv2.rotate(image, cv2.ROTATE_90_CLOCKWISE)
        elif direction == CellDirection.BT_LR:
            rotated_image = cv2.rotate(image, cv2.ROTATE_90_COUNTERCLOCKWISE)
        else:
            raise ValueError(f"Unsupported direction: {direction}")

        # Save the rotated image to the temporary file
        cv2.imwrite(tmp_filename, rotated_image)

        yield tmp_filename