import docx_ed.cfg as c


def is_Heading(style):
    our_style = style.lower()
    for r_h, en_h in c.TITLE.items():
        r_h = r_h.split()[0].lower()
        en_h = en_h.split()[0].lower()
        if r_h in our_style or en_h in our_style:
            return True
    return False


def is_picture_or_figure(style):
    our_style = style.lower()
    paint_allies = ['рисунок', 'picture', 'figure', 'shape', 'фигур']
    for st in paint_allies:
        if st in our_style:
            return True
    return False


def is_listing(paragraph):
    return len(paragraph._element.xpath('./w:pPr/w:numPr')) > 0
