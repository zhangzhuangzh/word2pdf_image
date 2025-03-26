import os
import fitz  # PyMuPDF
from PIL import Image


def pdfs_to_images_and_recombine(pdf_directory):
    # Identify all PDF files in the current folder
    pdf_files = [f for f in os.listdir(pdf_directory) if f.endswith('.pdf')]

    for pdf_file in pdf_files:
        pdf_path = os.path.join(pdf_directory, pdf_file)

        # Step 1: Convert PDF files to images
        image_files = []

        # Open the PDF file
        pdf_document = fitz.open(pdf_path)

        # Convert each PDF file to images, with each page as a separate image
        for page_num in range(len(pdf_document)):
            page = pdf_document.load_page(page_num)
            pix = page.get_pixmap()

            image_filename = f"{pdf_file[:-4]}_page_{page_num + 1}.png"
            image_path = os.path.join(pdf_directory, image_filename)
            pix.save(image_path)
            image_files.append(image_path)

        # Step 2: Recombine images into a new PDF and cleanup
        output_pdf = os.path.join(pdf_directory, f"{pdf_file[:-4]}_image.pdf")
        if image_files:
            images = [Image.open(img_file).convert('RGB') for img_file in image_files]
            images[0].save(output_pdf, save_all=True, append_images=images[1:])

            # Delete the intermediate images
            for img_file in image_files:
                os.remove(img_file)

        print(f"New PDF created at: {output_pdf}")


# Usage
pdf_directory = '.'  # Current directory
pdfs_to_images_and_recombine(pdf_directory)
