# merge_pdf_app.py（ボタン動作を追加した最終版）
import tkinter
from tkinter import ttk
# インポートは以下の３行を追加
from tkinter import filedialog
from tkinter import messagebox

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

def merge_pdf(inp_dir, out_dir):
    # 対象フォルダ
    input_dir = (inp_dir+"\\").replace('/', os.sep)
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
    merger = PyPDF2.PdfFileMerger()
    filenames = os.listdir(output_dir) 
    for file in filenames: 
            merger.append(output_dir + file)
    merger.write(out_dir)
    merger.close()

def ask_folder():
    """ 参照ボタンの動作
    """
    path = filedialog.askdirectory()
    folder_path.set(path)


def app():
    """ 実行ボタンの動作
    """
    input_dir = folder_path.get()
    # 保存するPDFファイルを指定
    output_file = filedialog.asksaveasfilename(
        filetypes=[("PDF files", "*.pdf")], defaultextension=".pdf"
    )
    if not input_dir or not output_file:
        return
    # 結合実行
    merge_pdf(input_dir, output_file)
    # メッセージボックス
    messagebox.showinfo("完了", "完了しました。")


# メインウィンドウ
# 窓を作る
main_win = tkinter.Tk()
# 窓のタイトルを設定
main_win.title("PDFを結合する")
# 窓の大きさを設定
main_win.geometry("500x100")

# メインフレーム
main_frm = ttk.Frame(main_win)
main_frm.grid(column=0, row=0, sticky=tkinter.NSEW, padx=5, pady=10)

# パラメータ
folder_path = tkinter.StringVar()

# ウィジェット（フォルダ名）
folder_label = ttk.Label(main_frm, text="フォルダ指定")
folder_box = ttk.Entry(main_frm, textvariable=folder_path)
folder_btn = ttk.Button(main_frm, text="参照", command=ask_folder)

# ウィジェット（実行ボタン）
app_btn = ttk.Button(main_frm, text="実行", command=app)

# ウィジェットの配置
folder_label.grid(column=0, row=0, pady=10)
folder_box.grid(column=1, row=0, sticky=tkinter.EW, padx=5)
folder_btn.grid(column=2, row=0)
app_btn.grid(column=1, row=1)

# 配置設定
main_win.columnconfigure(0, weight=1)
main_win.rowconfigure(0, weight=1)
main_frm.columnconfigure(1, weight=1)

# ウインドウ状態の維持
main_win.mainloop()
