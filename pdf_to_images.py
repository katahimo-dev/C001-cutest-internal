# 必要なライブラリ
# pip install pdf2image python-pptx pillow

from pdf2image import convert_from_path
import os

PDF_PATH = 'Integrated_Childcare_System_Foundation.pdf'
IMG_DIR = 'pdf_images'

def pdf_to_images(pdf_path, out_dir):
    os.makedirs(out_dir, exist_ok=True)
    images = convert_from_path(pdf_path)
    img_paths = []
    for i, img in enumerate(images):
        img_path = os.path.join(out_dir, f'page_{i+1}.png')
        img.save(img_path, 'PNG')
        img_paths.append(img_path)
    return img_paths

if __name__ == '__main__':
    img_paths = pdf_to_images(PDF_PATH, IMG_DIR)
    print(f'画像保存完了: {img_paths}')
