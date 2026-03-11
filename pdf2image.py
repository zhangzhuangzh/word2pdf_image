import os
import pymupdf  # PyMuPDF
from PIL import Image
from tkinter import Tk, filedialog

try:
    import win32com.client
    HAS_PYWIN32 = True
except ImportError:
    HAS_PYWIN32 = False

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
    """将 doc/docx 转换为 pdf（使用微软 Word 直接转 PDF）"""
    try:
        if not HAS_PYWIN32:
            raise Exception("需要安装 pywin32: pip install pywin32")
        
        base_name = os.path.splitext(os.path.basename(word_file))[0]
        pdf_file = os.path.join(output_dir, base_name + ".pdf")
        
        # 使用 Word COM 接口转换
        word_app = win32com.client.Dispatch("Word.Application")
        word_app.Visible = False
        word_app.DisplayAlerts = False
        
        try:
            print(f"Opening: {word_file}...")
            doc = word_app.Documents.Open(os.path.abspath(word_file))
            
            # 如果是 .doc，先转为 .docx（Word 内部操作）
            if word_file.lower().endswith('.doc'):
                print(f"Converting .doc to PDF using Microsoft Word...")
            else:
                print(f"Converting .docx to PDF using Microsoft Word...")
            
            # 直接保存为 PDF（FileFormat=17 表示 PDF）
            doc.SaveAs2(os.path.abspath(pdf_file), FileFormat=17)
            doc.Close()
            print(f"Successfully created PDF: {pdf_file}")
        finally:
            word_app.Quit()
        
        return pdf_file
    except Exception as e:
        print(f"Word转PDF失败: {e}")
        raise


def pdf_to_image_pdf(pdf_path, output_dir):
    """将 PDF 转换为纯图像 PDF"""
    image_files = []
    pdf_document = pymupdf.open(pdf_path)

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
        try:
            directory = os.path.dirname(word_file)
            # 1. Word 转 PDF
            pdf_file = word_to_pdf(word_file, directory)
            # 2. PDF 转纯图像 PDF
            image_pdf = pdf_to_image_pdf(pdf_file, directory)
            # 3. 删除中间 PDF
            if os.path.exists(pdf_file):
                os.remove(pdf_file)
                print(f"Deleted intermediate PDF: {pdf_file}")
            print(f"处理成功: {word_file}")
        except Exception as e:
            print(f"处理文件 {word_file} 时失败: {e}")
            import traceback
            traceback.print_exc()


if __name__ == "__main__":
    main()
