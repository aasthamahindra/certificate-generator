# Certificate Generator

This Python script generates certificates for participants based on the data in an Excel file and sends the certificates as PDF attachments via email. It uses Tkinter for the GUI, Pandas for handling Excel data, Pillow (PIL) for image processing, and smtplib for sending emails.

## Installation

1. **Clone the repository:**

    ```bash
    git clone https://github.com/aasthamahindra/certificate-generator.git
    cd certificate-generator
    ```

2. **Install dependencies:**

    ```bash
    pip install -r requirements.txt
    ```

3. **Update script:**

    - Replace placeholder values in the script with correct paths, filenames, and credentials.
    - Enable "Less secure app access" on your Gmail account if using Gmail for sending emails.

## Usage

1. **Ensure Excel file:**

    - Make sure you have an Excel file (`details.xlsx`) with participant data.
    - Columns: Name, Date of Joining, E-mail.

2. **Run the script:**

    ```bash
    python main.py
    ```

3. **Generate Certificates:**

    - Click the "Generate Certificate" button to generate certificates for participants who joined the program over 365 days ago.
    - Certificates will be saved as PDFs and sent via email to the participants.

## Note

- Adjust the Excel file path in the script based on your project structure.
- Customize the email body and subject according to your requirements.
- Ensure proper configuration for sending emails using the specified email account.

---

**Feel free to customize the README to include more details specific to your project.** Add sections as needed, such as troubleshooting, known issues, or future improvements.
