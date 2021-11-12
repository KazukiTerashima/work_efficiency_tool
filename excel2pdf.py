import win32com.client
import re 
import os

def excel2pdf(target):
    # 読み込み元ファイルと書き込み先を指定する
    input_file = "C:/Users/terashima/Desktop/input/" + target
    output_file = "C:/Users/terashima/Desktop/input/pdf/" + target[:-4] + "pdf"
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



filenames = os.listdir("C:/Users/terashima/Desktop/input/".replace('/', os.sep)) 
for file in filenames: 
    match = re.search("\.xlsx$", file) 
    if match: 
        print(file)
        excel2pdf(file)
