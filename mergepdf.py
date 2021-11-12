import PyPDF2
import os

merger = PyPDF2.PdfFileMerger()

filenames = os.listdir("C:/Users/terashima/Desktop/input/pdf/".replace('/', os.sep)) 
for file in filenames: 
        print(file)
        merger.append("C:/Users/terashima/Desktop/input/pdf/" + file)

merger.write("C:/Users/terashima/Desktop/input/pdf/" + input())
merger.close()
