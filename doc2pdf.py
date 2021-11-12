from docx2pdf import convert
import os


"""
use library: https://github.com/AlJohri/docx2pdf
"""

def convert_pdf(input_dir, output_dir):
    """
    docxファイルの保存されたフォルダを指定して、フォルダ格納データを全てpdfにして、指定フォルダに保存する
    :param input_dir: dir_name, default:output/,  outputフォルダを利用
    :param output_dir: dir_name, default:output_pdf/, output_pdfフォルダを利用
    :return: output_pdfフォルダにoutputフォルダのpdfが全て保存される
    """
    convert(input_dir, output_dir)

if __name__ == '__main__':
    input_dir = "C:/Users/terashima/Desktop/input/".replace('/', os.sep)
    output_dir = "C:/Users/terashima/Desktop/input/pdf/".replace('/', os.sep)
    convert_pdf(input_dir, output_dir)
