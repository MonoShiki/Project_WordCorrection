import json

import docx
from docx.shared import RGBColor

import docx_ed.style_parser as sp


class Template:
    def __init__(self,user_id,docx_: docx.Document):
        self.user_id = user_id
        self.docx_ = docx_

    def writeTemplates(self, data):
        with open(f'../files/user_json/{self.user_id}.json', 'w') as outfile:
            json.dump(data, outfile)


    @staticmethod
    def takeTemplate(paragraph):
        data = {
        }
        font_size = None
        font_style = None
        font_color = None
        font_bold = None
        font_italic = None

        for run in paragraph.runs:
            font_size = run.font.size.pt if run.font.size else None
            font_style = [run.font.name] + font_style if font_style else [run.font.name]
            font_bold = True if run.font.bold else False
            font_italic = True if run.font.italic else False
            font_color = run.font.color.rgb if run.font.color else RGBColor(0, 0, 0)
            if font_size and font_style: break

        alignment = paragraph.alignment if paragraph.alignment else paragraph.style.paragraph_format.alignment
        interval = paragraph.style.paragraph_format.line_spacing
        interval = interval if interval else 0
        if paragraph.paragraph_format.first_line_indent is not None:
            indent = round(paragraph.paragraph_format.first_line_indent.cm, 2)
        else:
            if paragraph.style.paragraph_format.left_indent:
                left_indent = round(paragraph.style.paragraph_format.left_indent.cm, 2)
            else:
                left_indent = 0

            # отступ справа
            if paragraph.style.paragraph_format.right_indent:
                right_indent = paragraph.style.paragraph_format.right_indent
            else:
                right_indent = 0
            indent = [left_indent, right_indent]

        data['indent'] = indent
        data["alignment"] = alignment
        data['font-style'] = font_style
        data["font-size"] = font_size
        data['interval'] = round(interval, 2)
        data['font_bold'] = font_bold
        data['font_italic'] = font_italic
        data["font_color"] = font_color
        return data

    def summarize(self, style: list[dict]) -> dict:
        false_mark = [0, False, None]
        if not len(style): return {}
        true_style = {key: [] for key in style[0]}
        for styl in style:
            for key in styl:
                val = styl[key]
                if val not in false_mark:
                    if isinstance(val, list):
                        if val[0] in false_mark: continue
                    true_style[key].append(val)
        for key in true_style:
            tp = true_style[key]
            tup = []
            for kk in tp:
                if isinstance(kk, list):
                    if isinstance(kk[0], float):
                        tup = sorted(kk)

            true_style[key] = tup if tup else list(set(tp))
            tr_val = true_style[key]
            if len(tr_val) < 2:
                if tr_val:
                    true_style[key] = tr_val[0]
                else:
                    true_style[key] = None
        return true_style

    def generate_gost(self) -> dict:
        headings = []
        main_text = []
        fig_pic = []
        listing = []
        for i, paragraph in enumerate(self.docx_.paragraphs):
            if sp.is_Heading(paragraph.style.name):
                headings.append(self.takeTemplate(paragraph))
            elif sp.is_picture_or_figure(paragraph.style.name):
                fig_pic.append(self.takeTemplate(paragraph))
            elif sp.is_listing(paragraph):
                listing.append(self.takeTemplate(paragraph))
        gost = {
            "name": 'user_gost',
            "heading": self.summarize(headings),
            "main_text": self.summarize(headings),
            "picture_or_figure": self.summarize(fig_pic),
            "listing": self.summarize(listing)
        }
        return gost


if __name__ == "__main__":
    templ = Template(user_id='1',docx_=docx.Document('../test2.docx'))
    gost = templ.generate_gost()
    templ.writeTemplates(gost)
