import os
from pathlib import Path
from pdf2image import convert_from_path
from dotenv import load_dotenv


load_dotenv()

# poppler/binを環境変数PATHに追加する
poppler_dir = os.environ['POPPLER']
os.environ["PATH"] += os.pathsep + str(poppler_dir)

# PDFファイルのパス
input_dir = os.environ['PDF']
filenames = os.listdir(input_dir)
for file in filenames:
    pdf_path = Path(input_dir+"/"+file)
    
    # PDF -> Image に変換（150dpi）
    pages = convert_from_path(str(pdf_path), 150)

    # 画像ファイルを１ページずつ保存
    image_dir = Path(os.environ['IMAGE'])
    for i, page in enumerate(pages):
        file_name = pdf_path.stem + "_{:02d}".format(i + 1) + ".jpeg"
        image_path = image_dir / file_name
        # JPEGで保存
        page.save(str(image_path), "JPEG")
