from flask import Flask, request, render_template, send_from_directory
import os
from model import extract_images_from_pdf, extract_text_from_images, create_ppt_from_text_in_images

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'output'
TEXT_FOLDER = 'text'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Ensure that the directories exist
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)
os.makedirs(TEXT_FOLDER, exist_ok=True)


@app.route('/')
def home():
    return render_template('index.html')


@app.route('/upload', methods=['POST'])
def upload_file():
    if request.method == 'POST':
        if 'file' not in request.files:
            return 'No file part'
        file = request.files['file']
        if file.filename == '':
            return 'No selected file'
        if file:
            pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
            file.save(pdf_path)  # This should work now as the directory exists

            # Clear the OUTPUT_FOLDER and TEXT_FOLDER before processing new PDF
            for folder in [OUTPUT_FOLDER, TEXT_FOLDER]:
                for filename in os.listdir(folder):
                    file_path = os.path.join(folder, filename)
                    os.remove(file_path)

            # Process the PDF and extract images
            extract_images_from_pdf(pdf_path, OUTPUT_FOLDER)

            # Extract text from images (optional, based on your requirement)
            extract_text_from_images(OUTPUT_FOLDER, TEXT_FOLDER)

            # Create a PowerPoint presentation from the images
            ppt_path = os.path.join(TEXT_FOLDER, 'extracted_text_presentation.pptx')
            create_ppt_from_text_in_images(OUTPUT_FOLDER, ppt_path)

            # Send the PowerPoint file
            return send_from_directory(TEXT_FOLDER, 'extracted_text_presentation.pptx', as_attachment=True)
        else:
            return 'Text extraction failed or no text found.'


if __name__ == '__main__':
    app.run(debug=True)

