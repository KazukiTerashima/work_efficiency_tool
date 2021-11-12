import docx2pdf
import win32com.client
import PyPDF2
import re
import os

def excel2pdf(input_file, output_file):
    #エクセルを開く
    app = win32com.client.Dispatch("Excel.Application")
    app.Visible = True
    app.DisplayAlerts = False
    # Excelでワークブックを読み込む
    book = app.Workbooks.Open(input_file)
    # PDF形式で保存
    xlTypePDF = 0
    book.ExportAsFixedFormat(xlTypePDF, output_file)
    #エクセルを閉じる
    app.Quit()

if __name__ == '__main__':
    print("Which dir(full path)?:", end="")
    # 対象フォルダ
    input_dir = (input()+"\\").replace('/', os.sep)
    filenames = os.listdir(input_dir)
    output_dir = (input_dir + "pdf/").replace('/', os.sep)
    # ディレクトリが存在しない場合、ディレクトリを作成する
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    for file in filenames:
        word_match = re.search("\.docx$", file) 
        if word_match: 
            docx2pdf.convert(input_dir+file, output_dir+file[:-5]+".pdf")
            print(file)
        excel_match = re.search("\.xlsx$", file) 
        if excel_match: 
            excel2pdf(input_dir+file, output_dir+file[:-5]+".pdf")
            print(file)
    print()
    print("-------------------------------")
    print("merging PDF from:" + output_dir)
    print("↓↓↓↓↓↓↓↓↓↓↓↓ target file ↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓")
    merger = PyPDF2.PdfFileMerger()

    filenames = os.listdir(output_dir) 
    num = 0
    for file in filenames: 
            merger.append(output_dir + file)
            print(file)
            num += 1
    print("↑↑↑↑↑↑↑↑↑↑↑↑ target file ↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑")
    print("total: " + str(num) + "file")
    print()
    print("-------------------------------")
    print("What merged file name?:", end="")
    merged_file_name = input()
    if merged_file_name[-4:] != ".pdf":
            merged_file_name += ".pdf"
    merger.write(output_dir + merged_file_name)
    merger.close()
