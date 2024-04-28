from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import RGBColor

setter_gost = {
    "по ширине": WD_ALIGN_PARAGRAPH.JUSTIFY, "по левому": WD_ALIGN_PARAGRAPH.LEFT,
    "по правому": WD_ALIGN_PARAGRAPH.RIGHT, "по центру": WD_ALIGN_PARAGRAPH.CENTER,
    "по умолчанию": WD_ALIGN_PARAGRAPH is None
}

sel_gost = {
   WD_ALIGN_PARAGRAPH.JUSTIFY: "по ширине",
   WD_ALIGN_PARAGRAPH.LEFT: "по левому",
   WD_ALIGN_PARAGRAPH.RIGHT: "по правому",
   WD_ALIGN_PARAGRAPH.CENTER: "по центру",
   None: "по умолчанию"
}

templ_sel_gost = {
   3: "по ширине",
   0: "по левому",
   2: "по правому",
   1: "по центру",
   None: "по умолчанию"
}



template_gost = {"по ширине": 3, "по левому": 0,
    "по правому": 2, "по центру": 1,
    "по умолчанию": WD_ALIGN_PARAGRAPH is None}

setter_eng = {
    "JUSTIFY": WD_ALIGN_PARAGRAPH.JUSTIFY, "LEFT": WD_ALIGN_PARAGRAPH.LEFT,
    "RIGHT": WD_ALIGN_PARAGRAPH.RIGHT, "CENTER": WD_ALIGN_PARAGRAPH.CENTER,
    "None": WD_ALIGN_PARAGRAPH is None
}

TITLE = {'Заголовок': 'Title',
         'Заголовок 1': 'Heading 1',
         'Заголовок 2': 'Heading 2',
         'Заголовок 3': 'Heading 3',
         'Заголовок 4': 'Heading 4'}

chekout = ['Выравнивание', 'Межстрочный интервал', 'Абзацный отступ']

colors = {
    'red': RGBColor(255, 0, 0),
    'green': RGBColor(0, 255, 0),
    'blue': RGBColor(0, 0, 255),
    'yellow': RGBColor(196, 120, 20),
    'pink': RGBColor(250, 105, 165)
}


basic_ans = 'не соответствует ГОСТу.\n\t  Он должен быть: '
exceptions = {
    'line_spacing': f'  Межстрочный интервал {basic_ans}',
    'alignment': f'  Выравнивание {basic_ans}',
    'indent': f'  Абзацный отступ {basic_ans}',
    'font-size': f'  Размер шрифта {basic_ans}',
    'font-style': f'  Стиль шрифта {basic_ans}',
}

color_exceptions = {
    'green': 'alignment',
    'blue': 'line_spacing',
    'yellow': 'indent',
    'pink': 'font-size',
    'red': 'font-style'
}
