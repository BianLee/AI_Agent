import json
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import base64
import requests
import os
from datetime import datetime
import re
import logging
from io import BytesIO

# Import pdfrw
from pdfrw import PdfReader, PdfWriter, PdfDict, PdfName, PdfObject

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

class EmailProcessor:
    def __init__(self, email_data):
        self.email_data = email_data
        self.attachments = []

    def parse_email(self):
        self.sender = self.email_data['sender']['email']
        self.sender_name = self.email_data['sender'].get('name', 'Unknown')
        self.subject = self.email_data['subject']
        self.body = self.extract_body(self.email_data['payload'])
        self.extract_attachments(self.email_data['payload'])
        return self.sender, self.subject, self.body, self.attachments

    def extract_body(self, payload):
        if 'parts' in payload:
            for part in payload['parts']:
                if part['mimeType'] == 'text/plain':
                    return part.get('content', '')
                elif 'parts' in part:
                    result = self.extract_body(part)
                    if result:
                        return result
        return ""

    def extract_attachments(self, payload):
        if 'parts' in payload:
            for part in payload['parts']:
                if part['mimeType'] == 'application/pdf':
                    content = part.get('content', '')
                    filename = part.get('filename', 'attachment.pdf')
                    if content and self.is_valid_base64(content):
                        pdf_content = base64.b64decode(content)
                        self.save_attachment(filename, pdf_content)
                    elif 'attachmentLink' in part:
                        pdf_content = self.download_attachment(part['attachmentLink'], filename)
                    else:
                        pdf_content = None
                    if pdf_content:
                        self.attachments.append({
                            'filename': filename,
                            'content': pdf_content
                        })
                elif 'parts' in part:
                    self.extract_attachments(part)

    def is_valid_base64(self, s):
        try:
            return base64.b64encode(base64.b64decode(s)).decode() == s
        except Exception:
            return False

    def save_attachment(self, filename, content):
        try:
            with open(filename, 'wb') as f:
                f.write(content)
            logging.info(f"Attachment '{filename}' saved.")
        except Exception as e:
            logging.error(f"Error saving attachment '{filename}': {e}")

    def download_attachment(self, url, filename):
        try:
            logging.info(f"Downloading attachment from: {url}")
            response = requests.get(url)
            response.raise_for_status()
            self.save_attachment(filename, response.content)
            return response.content
        except requests.exceptions.RequestException as e:
            logging.error(f"Error downloading attachment: {e}")
            return None


class PDFProcessor:
    def __init__(self, pdf_path=None):
        self.pdf_path = pdf_path
        if self.pdf_path:
            logging.info(f"Initializing PDFProcessor with file: {self.pdf_path}")
            if not os.path.exists(self.pdf_path):
                logging.warning(f"File {self.pdf_path} does not exist.")

    def analyze_pdf(self):
        if not self.pdf_path or not os.path.exists(self.pdf_path):
            logging.error(f"No valid PDF file found at {self.pdf_path}. Cannot analyze.")
            return {}
        try:
            pdf = PdfReader(self.pdf_path)
            fields = {}
            for page in pdf.pages:
                annotations = page.get('/Annots')
                if annotations:
                    for annotation in annotations:
                        if annotation.get('/Subtype') == '/Widget' and annotation.get('/T'):
                            field_name = annotation['/T']
                            if isinstance(field_name, PdfObject):
                                field_name = field_name.decode() if hasattr(field_name, 'decode') else str(field_name)
                            fields[field_name] = ''
            if fields:
                logging.info(f"PDF fields found: {list(fields.keys())}")
            else:
                logging.info("No form fields found in PDF.")
            return fields
        except Exception as e:
            logging.error(f"Error analyzing PDF: {e}")
            return {}

    def fill_form(self, data, output_path='filled_form.pdf'):
        if not self.pdf_path or not os.path.exists(self.pdf_path):
            logging.error(f"No valid PDF file found at {self.pdf_path}. Cannot fill form.")
            return None
        try:
            template_pdf = PdfReader(self.pdf_path)
            for page in template_pdf.pages:
                annotations = page.get('/Annots')
                if annotations:
                    for annotation in annotations:
                        if annotation.get('/Subtype') == '/Widget' and annotation.get('/T'):
                            field_name = annotation['/T']
                            if isinstance(field_name, PdfObject):
                                field_name = field_name.decode() if hasattr(field_name, 'decode') else str(field_name)
                            if field_name in data:
                                self.update_field(annotation, data[field_name])
            
            PdfWriter().write(output_path, template_pdf)
            logging.info(f"Filled PDF saved as '{output_path}'.")
            return output_path
        except Exception as e:
            logging.error(f"Error filling PDF form: {e}")
            return None

    def update_field(self, annotation, value):
        annotation.update(PdfDict(V=value, AS=value))
        self.update_appearance_stream(annotation, value)

    def update_appearance_stream(self, annotation, value):
        font_size = 10
        font = "Helvetica"
        rect = annotation.get('/Rect', [0, 0, 100, 20])
        width = rect[2] - rect[0]
        height = rect[3] - rect[1]
        ap = PdfDict(
            N=PdfDict(
                Type=PdfName.XObject,
                Subtype=PdfName.Form,
                BBox=[0, 0, width, height],
                Resources=PdfDict(
                    Font=PdfDict(F1=PdfDict(Type=PdfName.Font, Subtype=PdfName.Type1, BaseFont=font))
                ),
                Matrix=[1, 0, 0, 1, 0, 0],
            )
        )
        ap.N.stream = f"""
        BT
        /F1 {font_size} Tf
        2 {height - font_size - 2} Td
        ({value}) Tj
        ET
        """
        annotation.update(PdfDict(AP=ap))

def fill_bill_of_sale(pdf_path, data, output_path):
    processor = PDFProcessor(pdf_path)
    form_fields = processor.analyze_pdf()
    
    # Map the data to the form fields
    form_data = {
        'IDENTIFICATION NUMBER': data.get('vin', ''),
        'YEAR MODEL': data.get('year', ''),
        'MAKE': data.get('make', ''),
        'LICENSE PLATE/CF #': data.get('license_plate', ''),
        'MOTORCYCLE ENGINE #': data.get('engine_number', ''),
        'PRINT SELLER\'S NAME[S]': data.get('seller_name', ''),
        'PRINT BUYER\'S NAME[S]': data.get('buyer_name', ''),
        'MO': data.get('sale_date_month', ''),
        'DAY': data.get('sale_date_day', ''),
        'YR': data.get('sale_date_year', ''),
        'SELLING PRICE': data.get('selling_price', ''),
        'GIFT VALUE': data.get('gift_value', ''),
        'PRINT NAME': data.get('seller_name', ''),  # Repeated for seller signature
        'DATE': data.get('sale_date', ''),
        'DL, ID OR DEALER #': data.get('seller_id', ''),
        'MAILING ADDRESS': data.get('seller_address', ''),
        'CITY': data.get('seller_city', ''),
        'STATE': data.get('seller_state', ''),
        'ZIP': data.get('seller_zip', ''),
        'DAYTIME PHONE #': data.get('seller_phone', ''),
        # Add buyer information fields here
    }
    
    return processor.fill_form(form_data, output_path)

class ResponseGenerator:
    def __init__(self, sender_email, recipient_email, subject, body, attachment_path):
        self.sender_email = sender_email
        self.recipient_email = recipient_email
        self.subject = subject
        self.body = body
        self.attachment_path = attachment_path
        if not os.path.exists(attachment_path):
            logging.warning(f"Warning: Attachment file {attachment_path} does not exist.")

    def send_email(self):
        msg = MIMEMultipart()
        msg['From'] = self.sender_email
        msg['To'] = self.recipient_email
        msg['Subject'] = f"Re: {self.subject}"

        msg.attach(MIMEText(self.body, 'plain'))

        if os.path.exists(self.attachment_path):
            with open(self.attachment_path, 'rb') as f:
                pdf_attachment = MIMEApplication(f.read(), _subtype="pdf")
                pdf_attachment.add_header('Content-Disposition', 'attachment', filename=os.path.basename(self.attachment_path))
                msg.attach(pdf_attachment)
            logging.info("Email includes a PDF attachment.")
        else:
            logging.warning(f"Attachment file {self.attachment_path} not found. Email will be sent without attachment.")

        logging.info(f"Would send email to {self.recipient_email} with subject '{msg['Subject']}'")
        logging.info(f"Email body:\n{self.body}")

        # Uncomment and configure the following code to send the email
        # smtp_server = smtplib.SMTP('smtp.gmail.com', 587)
        # smtp_server.starttls()
        # smtp_server.login(self.sender_email, "your_email_password")  # Use environment variables in production
        # smtp_server.send_message(msg)
        # smtp_server.quit()

def main(json_file_path):
    # Load JSON data
    try:
        with open(json_file_path, 'r') as file:
            email_data = json.load(file)
    except FileNotFoundError:
        logging.error(f"JSON file not found: {json_file_path}")
        return
    except json.JSONDecodeError:
        logging.error(f"Invalid JSON in file: {json_file_path}")
        return

    # Process email
    email_processor = EmailProcessor(email_data)
    sender, subject, body, attachments = email_processor.parse_email()

    # Process PDF
    if attachments:
        pdf_filename = attachments[0]['filename']
        pdf_processor = PDFProcessor(pdf_filename)
    else:
        logging.error("No PDF attachment found.")
        return

    form_fields = pdf_processor.analyze_pdf()
    if not form_fields:
        logging.error("Failed to analyze PDF. Exiting.")
        return

    # Extract information
    name_match = re.search(r'Cheers,\s*(\w+)', body)
    name = name_match.group(1) if name_match else email_processor.sender_name or "Unknown"

    # Fill form
    form_data = {
        "Name": name,
        "Date": datetime.now().strftime('%Y-%m-%d'),
        # Add other fields as needed
    }

    filled_pdf_path = pdf_processor.fill_form(form_data, output_path='filled_form.pdf')
    if not filled_pdf_path:
        logging.error("Failed to fill PDF form. Exiting.")
        return

    # Generate and send response
    response_body = f"""Hello {name},

Thank you for your email. I've filled out the form and attached it to this email.

Best regards,
AI Assistant"""

    response_generator = ResponseGenerator(
        "your_email@example.com",  # Replace with your email
        sender,
        subject,
        response_body,
        filled_pdf_path
    )
    response_generator.send_email()

if __name__ == "__main__":
    json_file_path = 'easy.json'  # Replace with your JSON file path
    main(json_file_path)