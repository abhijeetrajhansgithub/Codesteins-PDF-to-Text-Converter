import fitz  # PyMuPDF
import os
import pytesseract
from PIL import Image
from pptx import Presentation
from pptx.util import Inches, Pt


def extract_images_from_pdf(pdf_path, output_folder):
    """
    Extracts images from a given PDF and saves them into the output_folder.
    """
    # Ensure output directory exists
    os.makedirs(output_folder, exist_ok=True)

    # Open the PDF file
    pdf_document = fitz.open(pdf_path)

    # Iterate over each page
    for page_index in range(len(pdf_document)):
        # Get the page
        page = pdf_document[page_index]

        # Get the images of the page
        image_list = page.get_images(full=True)

        # Print how many images found on this page
        if image_list:
            print(f"[+] Found a total of {len(image_list)} images in page {page_index}")
        else:
            print("[!] No images found on page", page_index)

        for image_index, img in enumerate(image_list):
            # Get the XREF of the image
            xref = img[0]

            # Extract the image bytes
            base_image = pdf_document.extract_image(xref)
            image_bytes = base_image["image"]

            # Get the image extension
            image_ext = base_image["ext"]

            # Construct the image name
            image_name = f"image{page_index + 1}_{image_index + 1}.{image_ext}"

            # Save the image
            image_path = os.path.join(output_folder, image_name)
            with open(image_path, "wb") as image_file:
                image_file.write(image_bytes)

            print(f"Saved image {image_index + 1} from page {page_index + 1}")

    pdf_document.close()


def extract_text_from_images(images_folder, text_folder):
    """
    Extracts text from all the images in images_folder and saves the resulting text files in text_folder.
    """
    # Ensure output directory exists
    os.makedirs(text_folder, exist_ok=True)

    # Iterate through all files in the images folder
    for image_name in os.listdir(images_folder):
        if image_name.endswith((".png", ".jpg", ".jpeg", ".tiff", ".bmp", ".gif")):
            # Construct image path
            image_path = os.path.join(images_folder, image_name)

            # Open the image
            image = Image.open(image_path)

            # Use pytesseract to do OCR on the image
            text = pytesseract.image_to_string(image, lang='eng')

            # Construct text file path
            text_file_name = os.path.splitext(image_name)[0] + '.txt'
            text_file_path = os.path.join(text_folder, text_file_name)

            # Save the text to a file
            with open(text_file_path, 'w', encoding='utf-8') as text_file:
                text_file.write(text)

            print(f"Processed {image_name}")


def create_ppt_from_text_in_images(images_folder, ppt_path):
    """
    Creates a PowerPoint presentation with the text extracted from images in images_folder and saves it to ppt_path.
    """
    prs = Presentation()
    for image_name in sorted(os.listdir(images_folder)):
        if image_name.endswith((".png", ".jpg", ".jpeg", ".tiff", ".bmp", ".gif")):
            # Construct image path
            image_path = os.path.join(images_folder, image_name)

            # Open the image and use pytesseract to do OCR on the image
            image = Image.open(image_path)
            text = pytesseract.image_to_string(image, lang='eng')

            # Create a new slide
            slide_layout = prs.slide_layouts[5]  # choosing a blank slide layout
            slide = prs.slides.add_slide(slide_layout)

            # Add text to slide
            txBox = slide.shapes.add_textbox(Pt(50), Pt(50), prs.slide_width - Pt(100), prs.slide_height - Pt(100))
            tf = txBox.text_frame
            tf.text = text  # Add the extracted text to the text frame

    prs.save(ppt_path)
    print(f"PowerPoint saved at {ppt_path}")
