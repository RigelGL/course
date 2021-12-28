from docxtpl import DocxTemplate, RichText, R
import jinja2
from docx.oxml.shared import qn
from datetime import datetime
from PIL import Image


def main():
    im = open("cheched.png")
    doc = DocxTemplate("fax.docx")
    context = {'to': 'Иванов Иван Иванович', 'from': 'Петров Пётр Петрович',
               'to_fax': '1234567', 'pages_count': '1',
               'to_phone': '+7 900 900 90 90', 'date': datetime.now().strftime('%d:%m:%Y'),
               'to_at': '', 'copy': 'Нет',
               'extra': 'Python is used',
               'speed':'+', 'broadcast': ''
               }

    doc.render(context)
    doc.save("generated_doc.docx")


if __name__ == '__main__':
    main()
