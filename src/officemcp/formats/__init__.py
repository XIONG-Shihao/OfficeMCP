"""One handler per format. Importing this package registers every format's actions."""

from . import docx, pdf, pptx, xlsx

HANDLERS = {"docx": docx.Handler(), "xlsx": xlsx.Handler(), "pptx": pptx.Handler(), "pdf": pdf.Handler()}
LIBRARIES = {
    "docx": ("python-docx", "docx"),
    "xlsx": ("openpyxl", "openpyxl"),
    "pptx": ("python-pptx", "pptx"),
    "pdf": ("pypdf", "pypdf"),
}
