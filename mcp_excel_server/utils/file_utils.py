"""
File utility functions for Word Document Server.
"""

import os
from typing import Tuple, Optional
import shutil
from typing import List

# Nueva función para obtener los directorios permitidos


def get_allowed_directories() -> List[str]:
    """Get the list of allowed directories from environment variables."""
    # Obtener de variable de entorno, con valor predeterminado si no existe
    allowed_dirs_str = os.environ.get("MCP_ALLOWED_DIRECTORIES", "./documents")
    # Dividir por comas si hay múltiples directorios
    allowed_dirs = [dir.strip() for dir in allowed_dirs_str.split(",")]
    # Asegurar que las rutas estén normalizadas
    return [os.path.abspath(dir) for dir in allowed_dirs]


# Nueva función para verificar si una ruta está en directorios permitidos
def is_path_in_allowed_directories(file_path: str) -> tuple[bool, Optional[str]]:
    """Check if the given file path is within allowed directories."""
    allowed_dirs = get_allowed_directories()
    abs_path = os.path.abspath(file_path)

    # Verificar si el archivo está en alguno de los directorios permitidos
    for allowed_dir in allowed_dirs:
        if os.path.commonpath([allowed_dir, abs_path]) == allowed_dir:
            return True, None

    return (
        False,
        f"Path '{file_path}' is not in allowed directories: {', '.join(allowed_dirs)}",
    )


def check_file_writeable(filepath: str) -> Tuple[bool, str]:
    """
    Check if a file can be written to.

    Args:
        filepath: Path to the file

    Returns:
        Tuple of (is_writeable, error_message)
    """
    # If file doesn't exist, check if directory is writeable
    if not os.path.exists(filepath):
        directory = os.path.dirname(filepath)
        # If no directory is specified (empty string), use current directory
        if directory == "":
            directory = "."
        if not os.path.exists(directory):
            return False, f"Directory {directory} does not exist"
        if not os.access(directory, os.W_OK):
            return False, f"Directory {directory} is not writeable"
        return True, ""

    # If file exists, check if it's writeable
    if not os.access(filepath, os.W_OK):
        return False, f"File {filepath} is not writeable (permission denied)"

    # Try to open the file for writing to see if it's locked
    try:
        with open(filepath, "a"):
            pass
        return True, ""
    except IOError as e:
        return False, f"File {filepath} is not writeable: {str(e)}"
    except Exception as e:
        return False, f"Unknown error checking file permissions: {str(e)}"


def create_document_copy(
    source_path: str, dest_path: Optional[str] = None
) -> Tuple[bool, str, Optional[str]]:
    """
    Create a copy of a document.

    Args:
        source_path: Path to the source document
        dest_path: Optional path for the new document. If not provided, will use source_path + '_copy.docx'

    Returns:
        Tuple of (success, message, new_filepath)
    """
    if not os.path.exists(source_path):
        return False, f"Source document {source_path} does not exist", None

    if not dest_path:
        # Generate a new filename if not provided
        base, ext = os.path.splitext(source_path)
        dest_path = f"{base}_copy{ext}"

    try:
        # Simple file copy
        shutil.copy2(source_path, dest_path)
        return True, f"Document copied to {dest_path}", dest_path
    except Exception as e:
        return False, f"Failed to copy document: {str(e)}", None


def ensure_xlsx_extension(filename: str) -> str:
    """
    Ensure filename has .xlsx extension.

    Args:
        filename: The filename to check

    Returns:
        Filename with .xlsx extension
    """
    if not filename.endswith(".xlsx"):
        return filename + ".xlsx"
    return filename
