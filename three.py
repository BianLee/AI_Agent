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
import subprocess
import sys

# NLP and PDF processing imports
import spacy
from rapidfuzz import fuzz
from PyPDF2 import PdfReader
from fillpdf import fillpdfs

# Configure logging
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler()
    ]
)

# Ensure SpaCy model is installed
def install_spacy_model(model_name):
    try:
        spacy.load(model_name)
    except OSError:
        logging.info(f"Downloading SpaCy model: {model_name}")
        subprocess.check_call([sys.executable, "-m", "spacy", "download", model_name])

MODEL_NAME = "en_core_web_sm"
install_spacy_model(MODEL_NAME)
nlp = spacy.load(MODEL_NAME)

class EmailProcessor:
    def __init__(self, email_data):
        """
        Initialize the EmailProcessor with email data.
        
        :param email_data: Dictionary containing the email data.
        """
        self.email_data = email_data
        self.attachments = []

    def parse_email(self):
        """
        Parse the email to extract sender, subject, body, and attachments.
        
        :return: Tuple containing sender email, subject, body, and list of attachments.
        """
        self.sender = self.email_data['sender']['email']
        self.sender_name = self.email_data['sender'].get('name', 'Unknown')
        self.subject = self.email_data['subject']
        self.body = self.extract_body(self.email_data['payload'])
        self.extract_attachments(self.email_data['payload'])
        return self.sender, self.subject, self.body, self.attachments

    def extract_body(self, payload):
        """
        Extract the plain text body from the email payload.
        
        :param payload: Dictionary representing the email payload.
        :return: Extracted plain text body.
        """
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
        """
        Extract PDF attachments from the email payload.
        
        :param payload: Dictionary representing the email payload.
        """
        if 'parts' in payload:
            for part in payload['parts']:
                if part['mimeType'] == 'application/pdf':
                    content = part.get('content', '')
                    filename = part.get('filename', 'attachment.pdf')
                    
                    # Check if the content is base64-encoded
                    if content and self.is_valid_base64(content):
                        pdf_content = base64.b64decode(content)
                        self.save_attachment(filename, pdf_content)
                    # Check if there's an attachment link to download the PDF
                    elif 'attachmentLink' in part:
                        pdf_content = self.download_attachment(part['attachmentLink'], filename)
                    else:
                        pdf_content = None
                    
                    # If PDF content is obtained, add it to attachments
                    if pdf_content:
                        self.attachments.append({
                            'filename': filename,
                            'content': pdf_content
                        })
                # Recursively handle nested parts
                elif 'parts' in part:
                    self.extract_attachments(part)

    def is_valid_base64(self, s):
        """
        Validate if a string is valid base64.
        
        :param s: String to validate.
        :return: Boolean indicating if the string is valid base64.
        """
        try:
            return base64.b64encode(base64.b64decode(s)).decode() == s
        except Exception:
            return False

    def save_attachment(self, filename, content):
        """
        Save the attachment content to a file.
        
        :param filename: Name of the file to save.
        :param content: Binary content of the file.
        """
        try:
            with open(filename, 'wb') as f:
                f.write(content)
            logging.info(f"Attachment '{filename}' saved successfully.")
        except Exception as e:
            logging.error(f"Error saving attachment '{filename}': {e}")

    def download_attachment(self, url, filename):
        """
        Download the attachment from a given URL and save it.
        
        :param url: URL to download the attachment from.
        :param filename: Name of the file to save.
        :return: Binary content of the downloaded file or None if failed.
        """
        try:
            logging.info(f"Downloading attachment from: {url}")
            response = requests.get(url)
            response.raise_for_status()
            self.save_attachment(filename, response.content)
            return response.content
        except requests.exceptions.RequestException as e:
            logging.error(f"Error downloading attachment from {url}: {e}")
            return None
# EmailProcessor class remains the same

class NLPPDFProcessor:
    def __init__(self, pdf_path=None):
        """
        Initialize the NLPPDFProcessor with the path to the PDF.
        
        :param pdf_path: Path to the PDF file.
        """
        self.pdf_path = pdf_path
        if self.pdf_path:
            logging.info(f"Initializing NLPPDFProcessor with file: {self.pdf_path}")
            if not os.path.exists(self.pdf_path):
                logging.warning(f"File {self.pdf_path} does not exist.")

    def analyze_pdf(self):
        """
        Analyze the PDF and extract form field names using PyPDF2.
        Returns a dictionary with field names as keys.
        
        :return: Dictionary of PDF form fields.
        """
        if not self.pdf_path or not os.path.exists(self.pdf_path):
            logging.error(f"No valid PDF file found at {self.pdf_path}. Cannot analyze.")
            return {}
        try:
            pdf = PdfReader(self.pdf_path)
            fields = {}
            if "/AcroForm" in pdf.trailer["/Root"]:
                form = pdf.trailer["/Root"]["/AcroForm"]
                if "/Fields" in form:
                    for field in form["/Fields"]:
                        field_obj = field.get_object()
                        field_name = field_obj.get("/T")
                        if field_name:
                            # Handle byte strings if necessary
                            if isinstance(field_name, bytes):
                                field_name = field_name.decode('utf-8')
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
        """
        Fill the PDF form using fillpdf library.
        
        :param data: Dictionary containing data to fill into the PDF.
        :param output_path: Path to save the filled PDF.
        :return: Path to the filled PDF if successful, else None.
        """
        if not self.pdf_path or not os.path.exists(self.pdf_path):
            logging.error(f"No valid PDF file found at {self.pdf_path}. Cannot fill form.")
            return None
        try:
            # Analyze the PDF to get form fields
            form_fields = self.analyze_pdf()
            if not form_fields:
                logging.warning("No form fields to fill.")
                return None

            updates = {}

            # Match data keys to PDF form fields using NLP and fuzzy matching
            for field_name in form_fields.keys():
                best_match = self.find_best_match(field_name, data)
                if best_match:
                    logging.debug(f"Matching field '{field_name}' with data key '{best_match}'")
                    updates[field_name] = data[best_match]

            if updates:
                logging.info(f"Updating the following fields: {updates}")
                # Use fillpdf to fill the PDF
                fillpdfs.write_fillable_pdf(self.pdf_path, output_path, updates)
                logging.info(f"Filled PDF saved as '{output_path}'.")
                return output_path
            else:
                logging.warning("No matching fields found to update.")
                return None
        except Exception as e:
            logging.error(f"Error filling PDF form: {e}")
            return None

    def find_best_match(self, field_name, data):
        """
        Find the best matching key in data for the given field_name.
        
        :param field_name: The name of the PDF form field.
        :param data: Dictionary containing data to fill.
        :return: Best matching key from data if score > threshold, else None.
        """
        field_doc = nlp(field_name.lower())
        best_match = None
        best_score = 0

        for key in data.keys():
            key_doc = nlp(key.lower())
            score = fuzz.token_sort_ratio(field_name.lower(), key.lower())

            # Boost score if there are matching entities
            field_entities = [ent.label_ for ent in field_doc.ents]
            key_entities = [ent.label_ for ent in key_doc.ents]
            if set(field_entities) & set(key_entities):
                score += 10
                logging.debug(f"Boosted score for field '{field_name}' and key '{key}' due to matching entities.")

            logging.debug(f"Matching field '{field_name}' with key '{key}': Score = {score}")

            if score > best_score:
                best_score = score
                best_match = key

        if best_match:
            logging.info(f"Best match for field '{field_name}' is '{best_match}' with score {best_score}")
        else:
            logging.info(f"No suitable match found for field '{field_name}'")

        return best_match if best_score > 40 else None  # Adjust threshold as needed
class ResponseGenerator:
    def __init__(self, sender_email, recipient_email, subject, body, attachment_path):
        """
        Initialize the ResponseGenerator with email details.
        
        :param sender_email: Sender's email address.
        :param recipient_email: Recipient's email address.
        :param subject: Subject of the email.
        :param body: Body of the email.
        :param attachment_path: Path to the attachment file.
        """
        self.sender_email = sender_email
        self.recipient_email = recipient_email
        self.subject = subject
        self.body = body
        self.attachment_path = attachment_path
        if not os.path.exists(attachment_path):
            logging.warning(f"Warning: Attachment file {attachment_path} does not exist.")

    def send_email(self):
        """
        Compose and send the email with the attachment.
        """
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
def call_gemini_api(prompt, temperature=0.2, max_tokens=500):
    """
    Function to interact with Gemini API.

    :param prompt: The prompt to send to the LLM.
    :param temperature: Sampling temperature.
    :param max_tokens: Maximum number of tokens to generate.
    :return: Generated text from Gemini.
    """
    api_url = "https://api.gemini.com/v1/chat/completions"  # Replace with actual Gemini API endpoint
    headers = {
        "Authorization": f"Bearer {os.getenv('GEMINI_API_KEY')}",
        "Content-Type": "application/json"
    }
    payload = {
        "model": "gemini-pro",  # Replace with the actual model name if different
        "messages": [
            {"role": "system", "content": "You are a helpful assistant for extracting information from emails to fill out forms."},
            {"role": "user", "content": prompt}
        ],
        "temperature": temperature,
        "max_tokens": max_tokens
    }

    try:
        response = requests.post(api_url, headers=headers, json=payload)
        response.raise_for_status()
        data = response.json()
        generated_text = data['choices'][0]['message']['content']
        return generated_text
    except requests.exceptions.RequestException as e:
        logging.error(f"Error calling Gemini API: {e}")
        return ""

def update_json_with_gemini(email_body, json_file_path):
    """
    Update JSON data using Gemini API based on email content.

    :param email_body: The body of the email.
    :param json_file_path: Path to the JSON file containing form data.
    :return: Updated JSON data.
    """
    # Load existing JSON data
    try:
        with open(json_file_path, 'r') as file:
            existing_data = json.load(file)
    except FileNotFoundError:
        existing_data = {}

    # Prepare prompt for Gemini
    prompt = f"""
    Given the following email content:

    {email_body}

    And the existing form data:

    {json.dumps(existing_data, indent=2)}

    Please update the form data with any new or changed information from the email. 
    Provide the result as a JSON object. If no updates are needed, return the original data.
    """

    # Call Gemini API
    response = call_gemini_api(prompt)

    try:
        updated_data = json.loads(response)
        # Save updated data back to file
        with open(json_file_path, 'w') as file:
            json.dump(updated_data, file, indent=2)
        logging.info(f"Updated JSON data saved to {json_file_path}")
        return updated_data
    except json.JSONDecodeError:
        logging.error("Failed to parse Gemini API response as JSON")
        return existing_data

def fill_form_with_nlp(pdf_path, data, output_path):
    """
    Function to fill PDF form using NLP-based field matching.
    
    :param pdf_path: Path to the input PDF form.
    :param data: Dictionary containing data to fill.
    :param output_path: Path to save the filled PDF.
    :return: Path to the filled PDF if successful, else None.
    """
    processor = NLPPDFProcessor(pdf_path)
    return processor.fill_form(data, output_path)

def main(json_file_path, email_json_path):
    """
    Main function to orchestrate email processing and PDF form filling.
    
    :param json_file_path: Path to the JSON file containing form data.
    :param email_json_path: Path to the JSON file containing email data.
    """
    # Load email data
    try:
        with open(email_json_path, 'r') as file:
            email_data = json.load(file)
        logging.info(f"Loaded email data from '{email_json_path}'.")
    except FileNotFoundError:
        logging.error(f"Email JSON file not found: {email_json_path}")
        return
    except json.JSONDecodeError:
        logging.error(f"Invalid JSON format in file: {email_json_path}")
        return

    # Process email
    email_processor = EmailProcessor(email_data)
    sender, subject, body, attachments = email_processor.parse_email()

    # Update JSON data using Gemini
    updated_data = update_json_with_gemini(body, json_file_path)

    # Download attachments
    downloaded_files = [attachment['filename'] for attachment in attachments]
    if not downloaded_files:
        logging.error("No PDF attachments to process. Exiting.")
        return

    # Process each downloaded PDF
    for pdf_filename in downloaded_files:
        filled_pdf_path = fill_form_with_nlp(pdf_filename, updated_data, output_path=f'filled_{pdf_filename}')
        if filled_pdf_path:
            logging.info(f"Successfully filled the PDF form: {filled_pdf_path}")

            # Generate and send response email
            response_body = f"""Hello,

Thank you for your email. I've filled out the form based on the information provided and attached it to this email.

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
        else:
            logging.error(f"Failed to fill the PDF form for: {pdf_filename}")

if __name__ == "__main__":
    json_file_path = 'form_data.json'  # Path to store and update form data
    email_json_path = 'easy.json'  # Path to email data JSON
    main(json_file_path, email_json_path)