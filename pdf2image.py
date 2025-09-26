import os
import fitz  # PyMuPDF
from PIL import Image
from tkinter import Tk, filedialog
from docx2pdf import convert

# -------- 新增：高分屏优化 --------
try:
    import ctypes
    ctypes.windll.shcore.SetProcessDpiAwareness(1)  # Windows 8.1 及以上
except Exception:
    try:
        ctypes.windll.user32.SetProcessDPIAware()  # 兼容 Windows 7
    except Exception:
        pass
# ---------------------------------


def word_to_pdf(word_file, output_dir):
    """将 doc/docx 转换为 pdf"""
    pdf_file = os.path.join(output_dir, os.path.splitext(os.path.basename(word_file))[0] + ".pdf")
    convert(word_file, pdf_file)
    return pdf_file


def pdf_to_image_pdf(pdf_path, output_dir):
    """将 PDF 转换为纯图像 PDF"""
    image_files = []
    pdf_document = fitz.open(pdf_path)

    for page_num in range(len(pdf_document)):
        page = pdf_document.load_page(page_num)
        pix = page.get_pixmap()
        image_filename = f"{os.path.splitext(os.path.basename(pdf_path))[0]}_page_{page_num + 1}.png"
        image_path = os.path.join(output_dir, image_filename)
        pix.save(image_path)
        image_files.append(image_path)

    output_pdf = os.path.join(output_dir, f"{os.path.splitext(os.path.basename(pdf_path))[0]}_预览版.pdf")
    if image_files:
        images = [Image.open(img_file).convert("RGB") for img_file in image_files]
        images[0].save(output_pdf, save_all=True, append_images=images[1:])

        # 删除中间图片
        for img_file in image_files:
            os.remove(img_file)

    print(f"New image-based PDF created: {output_pdf}")
    return output_pdf


def main():
    # 弹出文件选择窗口
    root = Tk()
    root.withdraw()  # 不显示主窗口
    word_files = filedialog.askopenfilenames(
        title="选择 Word 文件",
        filetypes=[("Word files", "*.doc *.docx")]
    )
    if not word_files:
        print("未选择文件")
        return

    for word_file in word_files:
        directory = os.path.dirname(word_file)
        # 1. Word 转 PDF
        pdf_file = word_to_pdf(word_file, directory)
        # 2. PDF 转纯图像 PDF
        image_pdf = pdf_to_image_pdf(pdf_file, directory)
        # 3. 删除中间 PDF
        if os.path.exists(pdf_file):
            os.remove(pdf_file)
            print(f"Deleted intermediate PDF: {pdf_file}")


if __name__ == "__main__":
    main()
