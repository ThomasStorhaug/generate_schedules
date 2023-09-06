# Konstanter

from docx.shared import Cm

WEEKS = "35-2023:25-2024"

CLASSES = [
    "1TIFA", 
    "1TIFB", 
    "1TIFC", 
    "1TIFD", 
    "1TIFE", 
    "1TIFF", 
    "1TIFG",
    ] 

COLORS = {
    "YFF" : {"text": "#ffffff", "background": "#8b4049"},
    "Naturfag" : {"text": "#000000", "background": "#c9cca1"},
    "Kroppsøving" : {"text": "#ffffff", "background": "#caa05a"},
    "Produktivitet og kvalitetsstyring" : {"text": "#ffffff", "background": "#ae6a47"},
    "Engelsk" : {"text": "#ffffff", "background": "#543344"},
    "Matematikk" : {"text": "#ffffff", "background": "#515262"},
    "Konstruksjon og styringsteknikk" : {"text": "#ffffff", "background": "#63787d"},
    "Produksjon og tjenester" : {"text": "#ffffff", "background": "#8ea091"},
    "accent": {"text": "#ffffff", "background": "#E2EFD9"},
    "fri": {"text": "#000000", "background": "#D9D9D9"}
}

SUBJECT_CODES = {
    "y": "YFF",
    "n": "Naturfag",
    "m": "Matematikk",
    "e": "Engelsk",
    "k": "Kroppsøving",
    "pk": "Produktivitet og kvalitetsstyring",
    "ks": "Konstruksjon og styringsteknikk",
    "pt": "Produksjon og tjenester"
}

CLASS_START = {
    1: "08:10",
    2: "08:55",
    3: "10:00",
    4: "10:45",
    5: "12:00",
    6: "12:45",
    7: "13:40",
    8: "14:25"
}

PAUSES = {
    1: "09:40-10:00",
    2: "11:30-12:00",
    3: "13:30-13:40"
}

CLASS_COLUMN_WIDTH = Cm(5)
TIMES_COLUMN_WIDTH = Cm(2)
CLASS_CELL_HEIGHT = Cm(1.5)

LEFT_MARGIN = Cm(0.5)
RIGHT_MARGIN = Cm(0.5)
TOP_MARGIN = Cm(0.5)
BOTTOM_MARGIN = Cm(0.5)

FONT_FAMILY = "Calibri"

TOP_ROW_TEXT = "Viktig info:"