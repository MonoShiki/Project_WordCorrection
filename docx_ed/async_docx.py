import asyncio
import os

import docx

import docx_ed.cfg as c
from docx_ed.file_reader import FileReader


def painter(paragraph: docx, errors: list[tuple]):
    for color, comment in errors:
        paint(paragraph, color, comment)


def paint(paragraph: docx, color: str, comment: str):
    paragraph.add_comment(comment)
    for word in paragraph.runs:
        word.font.color.rgb = c.colors[color]


def join_numbers(numbers):
    if not numbers:
        return ""
    numbers = list(map(int, numbers))
    ranges = []
    start = end = numbers[0]

    for num in numbers[1:]:
        if num == end + 1:
            end = num
        else:
            if start == end:
                ranges.append(str(start))
            else:
                ranges.append(f"{start}-{end}")
            start = end = num

    if start == end:
        ranges.append(str(start))
    else:
        ranges.append(f"{start}-{end}")

    return ", ".join(ranges)


class FileManager:
    functions: dict = {
        'alignment'
    }

    def __init__(self, user_id: int, docx_: docx.Document, name: str, doc_rej: bool = False,
                 gost=None):
        self.gost = gost
        self.user_id = user_id
        self.user_file = docx_
        self.user_file_name = name
        self.alignment = None
        self.indent = None
        self.interval = None
        self.fsize = None
        self.fname = None
        self.doc_rej = doc_rej

    @staticmethod
    def msg_errors(errors) -> str:
        errors = list(errors)
        if errors:
            comments = []
            for main_err, i in errors:
                comments.append(i)
            ans = main_err[1][1]
            abzac = join_numbers(comments)
            ans += f'\nВ абзацах: \n{abzac}\n\n\t'
            return ans
        return "Все соответствует госту"

    def lineal_is_choosen(self, function_name) -> str:
        paragraphs = self.user_file.paragraphs
        errors = []
        for i, paragraph in enumerate(paragraphs):
            errors.append((self.get_f_dict()[function_name](paragraph), i))
        return self.msg_errors(filter(lambda tup: tup[0][0] is True, errors))

    async def get_params_from_gost(self):
        if self.gost in FileReader.get_files().keys():
            params = await FileReader(self.gost + '.json').read_file()
            self.alignment = c.setter_gost[params['alignment']]
            self.indent = params['paragraph-indent']
            self.interval = params['interval']
            self.fname = params['font-style']
            self.fsize = params['font-size']
            return True
        return False

    def get_f_dict(self):
        return {
            'font-size': self.is_correct_font_size,
            'font-style': self.is_correct_font_style,
            'alignment': self.is_correct_alignment,
            'indent': self.is_correct_indent,
            'interval': self.is_correct_interval
        }

    def is_correct_font_size(self, paragraph) -> tuple:
        error = (False, ('pink', 'All-okey'))
        for run in paragraph.runs:
            font_size = run.font.size.pt if run.font.size else None
            if font_size is None: return error
            if isinstance(self.fsize, list):
                left, right = map(float, self.fsize)
                if not (left <= font_size <= right):
                    error = (True, ('pink', c.exceptions['font-size'] + '-'.join(self.fsize)))
                    break
            else:
                if font_size != float(self.fsize):
                    error = (True, ('pink', c.exceptions['font-size'] + str(self.fsize)))
                    break
        return error

    def is_correct_font_style(self, paragraph):
        error = (False, ('red', 'All-okey'))
        for run in paragraph.runs:
            font_style = run.font.name
            if font_style:
                if isinstance(self.fname, list):
                    if font_style not in self.fname:
                        error = (True, ('red', c.exceptions['font-style'] + '-'.join(self.fname)))
                else:
                    if font_style not in self.fname:
                        error = (True, ('red', c.exceptions['font-style'] + str(self.fname)))
        return error

    def is_correct_alignment(self, paragraph):
        error = (False, ('green', 'All-okey'))
        alignment = paragraph.alignment
        alignment = alignment if alignment else paragraph.style.paragraph_format.alignment
        if alignment:
            if paragraph.text.strip():
                if alignment != self.alignment:
                    error = (True, ('green', c.exceptions['alignment'] + str(self.alignment)))
        return error

    def is_correct_interval(self, paragraph):
        interval = paragraph.style.paragraph_format.line_spacing
        error = (False, ('yellow', 'All-okey'))
        if interval is None: return error
        interval = round(interval, 2)
        if isinstance(self.interval, list):
            left, right = map(float, self.interval)
            if not (left <= interval <= right):
                error = (True, ('yellow', c.exceptions['line_spacing'] + '-'.join(self.interval)))

        else:
            if interval != float(self.interval):
                error = (True, ('yellow', c.exceptions['line_spacing'] + str(self.interval)))
        return error

    def is_correct_indent(self, paragraph):
        error = (False, ('blue', 'All-okey'))
        if paragraph.paragraph_format.first_line_indent is not None:
            doc_indent = round(paragraph.paragraph_format.first_line_indent.cm, 2)
            if isinstance(self.indent, list):
                left, right = map(float, self.indent)
                if not (left <= doc_indent <= right):
                    error = (True, ('blue', c.exceptions['indent'] + '-'.join(self.indent)))
            else:
                if doc_indent != float(self.indent):
                    error = (True, ('blue', c.exceptions['indent'] + str(self.indent)))
        return error

    @staticmethod
    def is_not_Heading(style):
        our_style = style.lower()
        for r_h, en_h in c.TITLE.items():
            r_h = r_h.split()[0].lower()
            en_h = en_h.split()[0].lower()
            if r_h in our_style or en_h in our_style:
                return False
        return True

    @staticmethod
    def is_picture_or_figure(style):
        our_style = style.lower()
        paint_allies = ['рисунок', 'picture', 'figure', 'shape', 'фигур']
        for st in paint_allies:
            if st in our_style:
                return True
        return False

    @staticmethod
    def is_listing(paragraph):
        return len(paragraph._element.xpath('./w:pPr/w:numPr')) > 0

    async def is_correct_document(self):
        if await self.get_params_from_gost():
            errors = {'alignment': [],
                      'line_spacing': [],
                      'indent': [],
                      'font-size': [],
                      'font-style': []
                      }

            first_page = False  # Скипает первую страницы до Введения у отчетов
            for i, paragraph in enumerate(self.user_file.paragraphs):
                if first_page:
                    if not self.is_not_Heading(paragraph.style.name):
                        first_page = False
                    else:
                        continue

                if self.is_picture_or_figure(paragraph.style.name): continue
                if self.is_not_Heading(paragraph.style.name):
                    if self.is_listing(paragraph): continue
                    actual_font_size_info = self.is_correct_font_size(paragraph)  # Размер шрифта
                    actual_font_style_info = self.is_correct_font_style(paragraph)  # Стиль шрифта
                    actual_alignment_info = self.is_correct_alignment(paragraph)  # Выравнивание
                    actual_interval_info = self.is_correct_interval(paragraph)  # Межстрочный интервал
                    actual_indent_info = self.is_correct_indent(paragraph)  # Отступы
                    reports = [actual_font_size_info,
                               actual_font_style_info,
                               actual_alignment_info,
                               actual_interval_info,
                               actual_indent_info]
                    err_filter = filter(lambda tup: tup[0] is True, reports)

                    if err_filter:

                        err_comments = [err[1] for err in err_filter]
                        if self.doc_rej:
                            painter(paragraph, err_comments)
                        else:
                            for color, comment in err_comments:
                                errors[c.color_exceptions[color]].append(comment + f'\n На строке {i}')
            return self.answer(errors)
        return False

    def answer(self, errors: dict = None):
        if self.doc_rej:
            self.saver()
        else:
            answer = f'Проверка госта {self.gost}\n'
            for keyh in errors:
                if errors[keyh]:
                    abzac = []
                    for er in errors[keyh]:
                        reason, fix = er.split(':')
                        spec, strings = fix.split('\n')
                        abzac.append(strings.split()[-1])
                    abzac = join_numbers(abzac)
                    answer += f' Проблемы возникли с {keyh}: \n\t{reason.strip()} {spec}\n\t  В абзацах: \n{abzac}\n\n\t'
            if answer != f'Проверка госта {self.gost}\n':
                return answer
            else:
                return 'Все соотвествует ГОСТу'

    def saver(self):
        self.user_file.save(f'../files/edited_Docx/{self.user_id}_ready_file.docx')

    def close(self):
        os.remove(self.user_file_name)


if __name__ == '__main__':
    obj = FileManager(1, docx.Document('../test2.docx'), 'tur', gost="2.105-2019", doc_rej=True)
    asyncio.run(obj.is_correct_document())
