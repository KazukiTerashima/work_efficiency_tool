import docx2pdf
import win32com.client
import PyPDF2
import re
import os

if __name__ == '__main__':
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
                print(file)
            excel_match = re.search("\.xlsx$", file) 
            if excel_match:
                print(file)
        else:
            word_match = re.search("\.docx$", file)
            if word_match and os.path.getmtime(input_dir+file) > os.path.getmtime(output_dir+file[:-5]+".pdf"): 
                print(file)
            excel_match = re.search("\.xlsx$", file) 
            if excel_match and os.path.getmtime(input_dir+file) > os.path.getmtime(output_dir+file[:-5]+".pdf"): 
                print(file)
