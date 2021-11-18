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

def pdf_merger(m_out = ""):
    if not m_out:
        print("対象ディレクトリを入力:", end="")
        m_out = (input()+"\\").replace('/', os.sep)
    print()
    print("-------------------------------")
    print("merging PDF from:" + m_out)
    print("※※※※※※ ファイルの順番と数を確認してください ※※※※※※")
    merger = PyPDF2.PdfFileMerger()

    filenames = os.listdir(m_out) 
    num = 0
    for file in filenames: 
        merger.append(m_out + file)
        print(file)
        num += 1
    print("※※※※※※ ファイルの順番と数を確認してください ※※※※※※")
    print("total: " + str(num) + "file")
    print("-------------------------------")
    print()
    print("結合したPDFの名前を入力:", end="")
    merged_file_name = input()
    if merged_file_name[-4:] != ".pdf":
            merged_file_name += ".pdf"
    merger.write(m_out + merged_file_name)
    merger.close()

if __name__ == '__main__':
    print("excelファイルとwordファイルをPDF化して全てを結合します。")
    print("PDF化のみ：\"1\"、PDF結合のみ：\"2\",すべて実行する：0、でEnterキーを押してください")
    ptn = int(input())
    if ptn in [0, 1, 2]:
        if ptn == 0 or ptn == 1:
            print("対象ディレクトリを入力:", end="")
            # 対象フォルダ
            input_dir = (input()+"\\").replace('/', os.sep)
            filenames = os.listdir(input_dir)
            output_dir = (input_dir + "pdf/").replace('/', os.sep)
            # ディレクトリが存在しない場合、ディレクトリを作成する
            if not os.path.exists(output_dir):
                os.makedirs(output_dir)
            
            for file in filenames:
                if not os.path.exists(output_dir+file[:-5]+".pdf"):
                    word_match = re.search("\.docx$", file)
                    if word_match:
                        docx2pdf.convert(input_dir+file, output_dir+file[:-5]+".pdf")
                        print(file)
                    excel_match = re.search("\.xlsx$", file) 
                    if excel_match:
                        excel2pdf(input_dir+file, output_dir+file[:-5]+".pdf")
                        print(file)
                else:
                    word_match = re.search("\.docx$", file)
                    if word_match and os.path.getmtime(input_dir+file) > os.path.getmtime(output_dir+file[:-5]+".pdf"): 
                        docx2pdf.convert(input_dir+file, output_dir+file[:-5]+".pdf")
                        print(file)
                    excel_match = re.search("\.xlsx$", file) 
                    if excel_match and os.path.getmtime(input_dir+file) > os.path.getmtime(output_dir+file[:-5]+".pdf"): 
                        excel2pdf(input_dir+file, output_dir+file[:-5]+".pdf")
                        print(file)
            if ptn == 1:
                print("-------------------------------")
                print("merging PDF from:" + output_dir)
                print("※※※※※※ ファイルの順番と数を確認してください ※※※※※※")

                filenames = os.listdir(output_dir) 
                num = 0
                for file in filenames: 
                    print(file)
                print("※※※※※※ ファイルの順番と数を確認してください ※※※※※※")
                print("total: " + str(num) + "file")
                print("-------------------------------")
                print()
                print("結合したPDFの名前を入力:", end="")
                print("正常に処理が終了しました。")

        if ptn == 0 or ptn == 2:
            if ptn == 0:
                pdf_merger(output_dir)
            if ptn == 2:
                pdf_merger()
            print("正常に処理が終了しました。")
    else:
        print("入力が不正です。")

