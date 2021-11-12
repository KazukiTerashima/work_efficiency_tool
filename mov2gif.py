from moviepy.editor import *
import os
 
input_path = "C:/Users/terashima/Desktop/sample/qiita用.mp4".replace('/', os.sep)
output_path = "C:/Users/terashima/Desktop/sample/output.gif".replace('/', os.sep)
 
# 動画読み込み
clip = VideoFileClip(input_path)
# 動画のサイズ変更
clip = clip.resize(width=800)
# 動画をGIFアニメに変換
clip.write_gif(output_path, fps=1)
# 閉じる
clip.close()
