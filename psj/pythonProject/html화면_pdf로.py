import pdfkit

import os

config = pdfkit.configuration(wkhtmltopdf='C:/Program Files/wkhtmltopdf/bin/wkhtmltopdf.exe')

global folder

folder = input("폴더를 입력 하세요")

output = input("아웃풋 폴더를 입력 해주세요")


def htmltopdf(filename):
    name, ext = os.path.splitext(filename)

    copy_name = os.path.basename(filename)

    print(filename, name + ".pdf")

    pdfkit.from_file(filename, name + ".pdf", configuration=config)


def loop(folder):
    for dirpath, dirnames, filenames in os.walk(folder):

        for filename in filenames:

            name, ext = os.path.splitext(filename)

            ext = ext.upper()

            if "HTML" in ext:
                htmltopdf(dirpath + "/" + filename)


loop(folder)