# Flask-PDF-to-Zip


This project provides a simple web application built using Flask that allows users to upload a `.pdf` file, extract images from the PDF, convert the PDF to `.docx` format, and convert any tables in the document to `.csv`. All the extracted content (images, docx file, and csv files) is then packaged into a `.zip` file. The user can provide an email address, and the zip file will be sent as an attachment to that email.

## Features

- Upload a `.pdf` file to the server.
- Extract images from each page of the PDF and save them as `.jpg` files.
- Convert the PDF to `.docx` format.
- Convert tables from the `.docx` file into `.csv` files.
- Package all extracted content into a `.zip` file.
- Send the `.zip` file via email to the user.

## Project Structure

```
/PDF-to-Zip-Converter
│
├── /templates
│   ├── index.html               # HTML template for the upload form
├── app.py                       # Flask app script
├── requirements.txt             # Required Python packages
└── README.md                    # This README file
```

## Installation

### Prerequisites
Ensure you have Python 3.x installed on your machine.

1. **Clone the repository**:
   ```bash
   git clone https://github.com/yourusername/PDF-to-Zip-Converter.git
   cd PDF-to-Zip-Converter
   ```

2. **Create a virtual environment** (optional but recommended):
   ```bash
   python -m venv venv
   source venv/bin/activate  # On Windows, use `venv\Scripts\activate`
   ```

3. **Install dependencies**:
   You can install the required dependencies using the `requirements.txt` file.

   ```bash
   pip install -r requirements.txt
   ```

4. **Configure email settings**:
   Update the `MAIL_USERNAME` and `MAIL_PASSWORD` in `app.py` with your own email credentials (using a Gmail account or another SMTP server).

   **Important:** For Gmail, you may need to enable "Less Secure Apps" or use an [App Password](https://support.google.com/accounts/answer/185833?hl=en) if using 2-step verification.

### Run the App

1. **Start the Flask application**:
   ```bash
   python app.py
   ```

   The application will run locally at `http://127.0.0.1:8000/`.

2. **Upload your PDF file**:
   - Open the application in your browser (`http://127.0.0.1:8000/`).
   - Use the form to upload a `.pdf` file that you want to convert.
   - The app will:
     - Extract images from the PDF.
     - Convert the PDF to `.docx`.
     - Extract any tables from the `.docx` and save them as `.csv` files.
     - Package everything into a `.zip` file.
   - Provide your email address to receive the zip file.

### Usage

- Go to `http://127.0.0.1:8000/` in your browser.
- Select a `.pdf` file to upload.
- The app will extract the images, convert the PDF to a Word document, and convert any tables to CSV files.
- After the extraction, all files will be zipped into a `.zip` file and sent to the provided email address.

### Email Sending Feature

The app will send the resulting zip file via email. Ensure that the following fields are filled out:

- **Recipient's email**: The email where the zip file will be sent.
- **Sender's email credentials**: You need to configure the email details in the app (in `app.py`) for sending the email.

### Error Handling

The application checks for the following errors:

- **No file part**: If no file is provided in the request.
- **No selected file**: If the file selected is empty or not selected properly.
- **Invalid file format**: If the uploaded file is not a `.pdf` file.

These errors will be displayed as flash messages to the user.

### Requirements

- Python 3.x or higher
- Flask
- Pillow (for image processing)
- PyMuPDF (`fitz`) (for PDF processing)
- pdf2docx (for converting PDF to DOCX)
- python-docx (for handling DOCX files)
- Flask-Mail (for sending emails)
- CSV (for exporting tables as CSV)

You can install the required Python libraries via the `requirements.txt` file:

```bash
pip install -r requirements.txt
```
