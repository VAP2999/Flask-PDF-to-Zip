from flask import Flask, render_template, request, send_file, flash
import os
from PIL import Image
import fitz
import io
from docx import Document
import zipfile
import shutil
from pdf2docx import Converter
import pdf2docx
import csv
from flask_mail import * 

app = Flask(__name__)
app.config['MAIL_SERVER'] = 'smtp.gmail.com'
app.config['MAIL_PORT'] = 465
app.config['MAIL_USERNAME'] = '30primero@gmail.com'
app.config['MAIL_PASSWORD'] = 'zilmflvkgtdyhixl'
app.config['MAIL_USE_TLS'] = False
app.config['MAIL_USE_SSL'] = True
app.secret_key = '1234'
mail = Mail(app)


@app.route('/')
@app.route('/index', methods=['GET','POST'])
def index():
    if request.method == 'POST':
        pdf_file = request.files['pdf']
        if pdf_file.filename != '':
            
            pdf_folder = pdf_file.filename.split(".")[0]
            os.mkdir(pdf_folder)
            os.mkdir(pdf_folder + r'/imgs')
            pdf_path = os.path.join(pdf_folder, pdf_file.filename)
            pdf_file.save(pdf_path)

            img_path = pdf_folder + r'/imgs'
            zip_file_path = extract_images_from_pdf(pdf_folder, pdf_path, img_path)
            shutil.rmtree(pdf_folder)

            msg = Message('Requested zip file', sender = 'no-reply@demo.com', recipients = [request.form['email']], body = pdf_folder + " zip file:")
            # Attach the file to the message
            with app.open_resource(pdf_folder + ".zip") as attachment:
                msg.attach(pdf_folder + ".zip", 'application/zip', attachment.read())

            try:
                # Send the email
                mail.send(msg)
                flash('Email sent successfully', 'success')
            except Exception as e:
                flash('An error occurred while sending the email', 'error')
            
            os.remove(zip_file_path)

        return render_template('index.html')        

    return render_template('index.html')

def extract_images_from_pdf(pdf_folder, pdf_path, img_path):


    pdf_file = fitz.open(pdf_path)
    
    # Iterate over PDF pages
    for page_index in range(len(pdf_file)):
        # Get the page itself
        page = pdf_file[page_index]
        image_list = page.get_images(full=True)
        # Printing the number of images found on this page
        if image_list:
            print(f"[+] Found a total of {len(image_list)} images in page {page_index}")
        else:
            print("[!] No images found on page", page_index)
        image_filenames = []  # Store image filenames for sorting
        for image_index, img in enumerate(page.get_images(), start=1):
            # Get the XREF of the image
            xref = img[0]
            # Extract the image bytes
            base_image = pdf_file.extract_image(xref)
            image_bytes = base_image["image"]
            # Get the image extension
            image_ext = base_image["ext"]
            rect = page.get_image_bbox(img[7])
            x0, y0, x1, y1 = rect.x0, rect.y0, rect.x1, rect.y1
            # Create a PIL Image from image data and convert to RGB mode
            pil_image = Image.open(io.BytesIO(image_bytes)).convert("RGB")
            # Save the image in the folder
            image_filename = f"{pdf_folder}_page{page_index}_{image_index}.jpg"
            image_path = os.path.join(img_path, image_filename)
            pil_image.save(image_path)
        
    doc_path = os.path.join(pdf_folder, pdf_folder + r'.docx')
    cv = Converter(pdf_path)
    cv.convert(doc_path, start=0, end=None)
    cv.close()

    doc = Document(doc_path)

    for table_index, table in enumerate(doc.tables):
        csv_file_name = os.path.join(pdf_folder, f"{pdf_folder}_{table_index + 1}.csv")

        
        with open(csv_file_name, 'w', newline='') as csv_file:
            csv_writer = csv.writer(csv_file)

            
            for row in table.rows:
                csv_writer.writerow([cell.text for cell in row.cells])


    # Define the folder you want to zip
    folder_to_zip = pdf_folder
    output_zip_filename = pdf_folder + r'.zip'
    
    with zipfile.ZipFile(output_zip_filename, 'w', zipfile.ZIP_DEFLATED) as zipf:
        # Walk through the folder and its subdirectories
        for foldername, subfolders, filenames in os.walk(folder_to_zip):
            # Add the current folder to the zip file
            zipf.write(foldername, os.path.relpath(foldername, folder_to_zip))

            # Add all files in the current folder to the zip file
            for filename in filenames:
                file_path = os.path.join(foldername, filename)
                arcname = os.path.relpath(file_path, folder_to_zip)
                zipf.write(file_path, arcname=arcname)

    return output_zip_filename

if __name__=='__main__':
    app.run(debug=True, port=8000)