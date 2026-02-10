# 必要なライブラリ
# pip install pymupdf

import fitz  # PyMuPDF
import os

PDF_PATH = 'TAYO-LINE_Lite_開発提案.pdf'
# ファイルが現在のディレクトリにない場合、親ディレクトリを探す
if not os.path.exists(PDF_PATH) and os.path.exists(os.path.join('..', PDF_PATH)):
    PDF_PATH = os.path.join('..', PDF_PATH)

IMG_DIR = 'pdf_images'

def pdf_to_images(pdf_path, out_dir):
    os.makedirs(out_dir, exist_ok=True)
    doc = fitz.open(pdf_path)
    img_paths = []

    for page_index in range(len(doc)):
        page = doc[page_index]
        image_list = page.get_images()
        
        for image_index, img in enumerate(image_list, start=1):
            xref = img[0]
            try:
                pix = fitz.Pixmap(doc, xref)
                # CMYK等の場合はRGBに変換
                if pix.n - pix.alpha > 3:
                    pix = fitz.Pixmap(fitz.csRGB, pix)
                
                # 画像ファイル名を生成 (常にPNG)
                img_filename = f'page_{page_index + 1}_img_{image_index}.png'
                img_path = os.path.join(out_dir, img_filename)
                
                pix.save(img_path)
                img_paths.append(img_path)
            except Exception as e:
                print(f"警告: ページ{page_index + 1}の画像{image_index}の保存に失敗しました: {e}")
            
    return img_paths

if __name__ == '__main__':
    if os.path.exists(PDF_PATH):
        img_paths = pdf_to_images(PDF_PATH, IMG_DIR)
        print(f'画像保存完了: {len(img_paths)} 枚の画像を保存しました。')
    else:
        print(f'エラー: PDFファイルが見つかりません: {PDF_PATH}')
