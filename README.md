# Bulk Email Sender & PDF to DOCX Converter

-----

This project combines two useful Python utilities: a **GUI-based bulk email sender** for marketing or notifications, and a simple **PDF to DOCX converter**.

## Table of Contents

  * [Features](https://www.google.com/search?q=%23features)
  * [Bulk Email Sender](https://www.google.com/search?q=%23bulk-email-sender)
      * [How it Works](https://www.google.com/search?q=%23how-it-works)
      * [Prerequisites](https://www.google.com/search?q=%23prerequisites)
      * [Installation](https://www.google.com/search?q=%23installation)
      * [Usage](https://www.google.com/search?q=%23usage)
  * [PDF to DOCX Converter](https://www.google.com/search?q=%23pdf-to-docx-converter)
      * [How it Works](https://www.google.com/search?q=%23how-it-works-1)
      * [Prerequisites](https://www.google.com/search?q=%23prerequisites-1)
      * [Usage](https://www.google.com/search?q=%23usage-1)
  * [Contributing](https://www.google.com/search?q=%23contributing)
  * [License](https://www.google.com/search?q=%23license)

-----

## Features

  * **Bulk Email Sender (GUI)**:
      * Send personalized emails to a list of recipients from an Excel file.
      * User-friendly Tkinter graphical interface.
      * Configurable email subject and body.
      * **SMTP settings** (server, port, sender email, password) can be easily configured via a settings dialog.
      * Progress bar and live status log for monitoring sending progress.
      * Multi-threaded email sending to prevent GUI freezing.
  * **PDF to DOCX Converter**:
      * Simple script to convert PDF documents into editable Word (`.docx`) files.
      * Extracts text content page by page and adds it to a new Word document.

-----

## Bulk Email Sender

This tool allows you to send emails to multiple recipients by reading their email addresses from an Excel file.

### How it Works

The `BulkEmailApp` class creates a Tkinter GUI.

1.  **Settings:** You can configure your sender email address, password, SMTP server, and port through a dedicated "Settings" dialog. These details are stored within the application instance.
2.  **Excel Upload:** Users upload an Excel file (`.xlsx`) containing an 'Email' column. The application reads these emails into a pandas DataFrame.
3.  **Email Content:** Subject and body of the email can be entered directly into the GUI.
4.  **Sending Process:**
      * When "Send Emails" is clicked, a new thread is spawned to handle the email sending process in the background, keeping the GUI responsive.
      * It iterates through each email in the loaded DataFrame.
      * For each email, it constructs an email message using `MIMEMultipart` and sends it via `smtplib.SMTP_SSL`.
      * A progress bar updates as emails are sent, and a status log displays real-time feedback (success or failure).

### Prerequisites

  * Python 3.x
  * Required Python libraries: `tkinter` (usually built-in), `pandas`, `smtplib`, `email.mime.multipart`, `email.mime.text`, `threading`, `openpyxl` (for reading .xlsx files).

### Installation

1.  **Clone the repository:**

    ```bash
    git clone https://github.com/sidsrbh/bulk_email.git # Replace with your actual repo URL
    cd your-repo-name
    ```

2.  **Create a virtual environment (recommended):**

    ```bash
    python3 -m venv venv
    source venv/bin/activate  # On Windows: venv\Scripts\activate
    ```

3.  **Install dependencies:**
    Create a `requirements.txt` file in your project root with the following content:

    ```
    pandas
    openpyxl
    ```

    Then run:

    ```bash
    pip install -r requirements.txt
    ```

### Usage

1.  **Run the application:**

    ```bash
    python your_main_script_name.py # Assuming the BulkEmailApp code is in a file like main.py
    ```

    A GUI window will appear.

2.  **Configure Settings:**

      * Click the **"Settings"** button.
      * Enter your **sender email**, **password** (an app-specific password if you use Gmail/Outlook 2FA), **SMTP server** (e.g., `smtp.hostinger.com`, `smtp.gmail.com`), and **Port** (e.g., `465` for SSL, `587` for TLS).
      * Click "Ok" to save.

3.  **Prepare your Excel file:**

      * Create an Excel file (e.g., `recipients.xlsx`) with a column named `Email` containing the recipient email addresses.
      * Example `recipients.xlsx`:
        ```
        Name    Email             City
        John    john@example.com  New York
        Jane    jane@test.org     London
        ```

4.  **Upload Excel File:**

      * Click **"Upload Excel File"** and select your `recipients.xlsx`.
      * A confirmation message will appear if the file is loaded successfully.

5.  **Enter Subject and Body:**

      * Type your desired email subject in the "Subject" field.
      * Compose your email body in the "Body" text area.

6.  **Send Emails:**

      * Click **"Send Emails"**.
      * The progress bar will update, and the status log will show the sending activity.

-----

## PDF to DOCX Converter

This script provides a straightforward way to convert the text content of a PDF file into a Microsoft Word (`.docx`) document.

### How it Works

The `pdf_to_docx` function utilizes two key libraries:

  * **PyMuPDF (`fitz`):** To open and extract text from each page of the PDF document.
  * **python-docx (`docx`):** To create a new Word document and add the extracted text as paragraphs.

### Prerequisites

  * Python 3.x
  * Required Python libraries: `PyMuPDF` (`fitz`), `python-docx` (`docx`).

### Installation

1.  **Install dependencies:**
    If you haven't already from the email sender, add these to your `requirements.txt`:
    ```
    PyMuPDF
    python-docx
    ```
    Then run:
    ```bash
    pip install -r requirements.txt
    ```

### Usage

1.  **Place your PDF:**

      * Put the PDF file you want to convert (e.g., `input.pdf`) in the same directory as your Python script, or provide its full path.

2.  **Run the script:**

    ```python
    # Ensure this code block is in a Python file (e.g., pdf_converter.py)
    import fitz  # PyMuPDF
    from docx import Document

    def pdf_to_docx(pdf_path, docx_path):
        document = Document()
        pdf_document = fitz.open(pdf_path)
        
        for page in pdf_document:
            text = page.get_text()
            document.add_paragraph(text)

        document.save(docx_path)
        pdf_document.close()

    # Example usage - IMPORTANT: Change 'input.pdf' and 'output.docx' as needed
    pdf_to_docx('input.pdf', 'output.docx')
    ```

    Run this script:

    ```bash
    python pdf_converter.py
    ```

    A new Word document named `output.docx` will be created in the same directory, containing the extracted text from your PDF.

    **Note:** This converter extracts *text* only. It does not preserve formatting, images, or complex layouts from the original PDF.

-----

## Contributing

Contributions are welcome\! If you find a bug, have a feature request, or want to improve the code, please feel free to:

1.  Fork the repository.
2.  Create a new branch (`git checkout -b feature/your-feature-name`).
3.  Make your changes.
4.  Commit your changes (`git commit -m 'feat: Add new feature'`).
5.  Push to the branch (`git push origin feature/your-feature-name`).
6.  Open a Pull Request.

-----

## License

This project is open-sourced under the [MIT License](https://www.google.com/search?q=LICENSE).

-----