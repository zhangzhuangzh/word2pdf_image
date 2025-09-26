# Word to Image-PDF Converter

## 功能说明
该工具用于批量选择 Word 文件（`.doc` / `.docx`），先转换为 PDF，再将 PDF 的每一页渲染为图像，最后合成为新的纯图像 PDF 文件（带 `_预览版` 后缀）。中间生成的临时 PNG 和 PDF 会自动删除，只保留最终的纯图像 PDF。支持多文件同时处理。

在 4K 等高分屏环境下，文件选择对话框经过 DPI 优化，显示清晰。

## 主要依赖
- [Python 3.8+](https://www.python.org/)
- [PyMuPDF (fitz)](https://pymupdf.readthedocs.io/)
- [Pillow](https://pillow.readthedocs.io/)
- [docx2pdf](https://github.com/AlJohri/docx2pdf) （依赖 Microsoft Word，仅限 Windows / macOS）
- `tkinter`（Python 内置库）

安装依赖：
pip install pymupdf pillow docx2pdf pyinstaller
打包命令：
pyinstaller --onefile --console word2pdf_image.py
