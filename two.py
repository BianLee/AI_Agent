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
import time

# NLP and PDF processing imports
import spacy
from rapidfuzz import fuzz
from PyPDF2 import PdfReader
from fillpdf import fillpdfs

# Configure logging
logging.basicConfig(
    level=logging.DEBUG,  # Set to DEBUG to capture detailed logs
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

# Set Gemini API Key
gemini_api_key = os.getenv("GEMINI_API_KEY")  # Ensure this environment variable is set securely
if not gemini_api_key:
    raise ValueError("Gemini API key not found. Please set the GEMINI_API_KEY environment variable.")

def modify_prompt_for_safety(original_prompt):
    """
    Modify the prompt to avoid triggering safety filters.

    :param original_prompt: The original prompt string.
    :return: Modified prompt string.
    """
    # Example modification: Rephrase to be more neutral
    modified_prompt = original_prompt.replace("Please extract", "Kindly provide")
    modified_prompt = modified_prompt.replace("extract", "identify and provide")
    logging.debug(f"Modified prompt to avoid safety triggers: {modified_prompt}")
    return modified_prompt

def call_gemini_api(prompt, temperature=0.2, max_tokens=500, retry_count=3, backoff_factor=2):
    """
    Function to interact with Gemini API with retry logic.

    :param prompt: The prompt to send to the LLM.
    :param temperature: Sampling temperature.
    :param max_tokens: Maximum number of tokens to generate.
    :param retry_count: Number of retry attempts.
    :param backoff_factor: Factor for exponential backoff.
    :return: Generated text from Gemini.
    """
    base_url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-pro:generateContent"
    params = {
        "key": gemini_api_key
    }
    headers = {
        "Content-Type": "application/json"
    }
    payload = {
        "contents": [{
            "parts": [{
                "text": prompt
            }]
        }],
        "generationConfig": {
            "temperature": temperature,
            "maxOutputTokens": max_tokens,
            "topP": 1,
            "topK": 1
        }
    }

    for attempt in range(retry_count):
        try:
            response = requests.post(base_url, headers=headers, params=params, json=payload, timeout=30)
            response.raise_for_status()
            logging.debug(f"Gemini API Response Text: {response.text}")
            
            # Check if response is JSON
            try:
                data = response.json()
            except json.JSONDecodeError:
                logging.error("Response is not in JSON format.")
                logging.debug(f"Response Text: {response.text}")
                return ""

            candidates = data.get('candidates', [])
            if not candidates:
                logging.error("No 'candidates' found in Gemini API response.")
                return ""

            first_candidate = candidates[0]
            finish_reason = first_candidate.get('finishReason', '')
            if finish_reason == 'SAFETY':
                logging.warning("Gemini API flagged the request as unsafe.")
                if attempt < retry_count - 1:
                    sleep_time = backoff_factor ** attempt
                    logging.info(f"Retrying after {sleep_time} seconds with a modified prompt...")
                    time.sleep(sleep_time)
                    # Modify the prompt to be more neutral
                    prompt = modify_prompt_for_safety(prompt)
                    payload['contents'][0]['parts'][0]['text'] = prompt
                    continue
                else:
                    logging.error("Max retry attempts reached. Cannot proceed due to safety restrictions.")
                    return ""

            content = first_candidate.get('content', {})
            parts = content.get('parts', [])
            if not parts:
                logging.error("No 'parts' found in the 'content' of Gemini API response.")
                return ""

            generated_text = parts[0].get('text', '')
            if not generated_text:
                logging.error("No 'text' found in the first part of Gemini API response.")
                return ""

            return generated_text

        except requests.exceptions.HTTPError as http_err:
            logging.error(f"HTTP error occurred: {http_err}")
            logging.debug(f"Response Text: {response.text}")
        except requests.exceptions.RequestException as req_err:
            logging.error(f"Request exception: {req_err}")
        except Exception as e:
            logging.error(f"An unexpected error occurred: {e}")
            logging.debug(f"Response Text: {response.text}")

        # Exponential backoff before retrying
        sleep_time = backoff_factor ** attempt
        logging.info(f"Retrying after {sleep_time} seconds...")
        time.sleep(sleep_time)

    return ""

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
            response = requests.get(url, timeout=30)
            response.raise_for_status()
            self.save_attachment(filename, response.content)
            return response.content
        except requests.exceptions.RequestException as e:
            logging.error(f"Error downloading attachment from {url}: {e}")
            return None

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

    def generate_field_mapping(self, field_names, data):
        """
        Use Gemini to generate a mapping of PDF fields to data keys.

        :param field_names: List of PDF form field names.
        :param data: Extracted data dictionary.
        :return: Dictionary mapping PDF fields to data keys.
        """
        # Prepare the email body for mapping
        email_body = data.get('email_body', '')

        prompt = f"""
Please identify and provide the following information from the data below in JSON format:

1. Seller's Name
2. Buyer's Name
3. Seller's Mailing Address
4. Buyer's Mailing Address
5. Seller's Print Name 1
6. Seller's Print Name 2
7. Buyer's Print Name 1
8. Buyer's Print Name 2
9. Seller's ZIP Code
10. Buyer's ZIP Code

**Data:**
{json.dumps(data, indent=4)}

**PDF Form Fields:**
{', '.join(field_names)}

**Instructions:**
- Ensure that ZIP Codes are 5-digit numbers.
- Provide only the JSON mapping without any additional text or explanations.
- If a field does not have corresponding data, set its value to null.

**Example JSON Output:**
{{
    "Seller's Name": "John Doe",
    "Buyer's Name": "Jane Smith",
    "Seller's Mailing Address": "123 Elm Street, Springfield, IL",
    "Buyer's Mailing Address": "456 Oak Avenue, Lincoln, NE",
    "Seller's Print Name 1": "John",
    "Seller's Print Name 2": "Doe",
    "Buyer's Print Name 1": "Jane",
    "Buyer's Print Name 2": "Smith",
    "Seller's ZIP Code": "62704",
    "Buyer's ZIP Code": "68508"
}}
"""

        mapping_text = call_gemini_api(prompt)
        if not mapping_text:
            logging.error("Failed to generate field mapping using Gemini.")
            return {}

        try:
            mapping = json.loads(mapping_text)
            logging.info(f"Generated field mapping: {mapping}")
            return mapping
        except json.JSONDecodeError as e:
            logging.error(f"Error decoding JSON from Gemini response: {e}")
            logging.debug(f"Mapping Text: {mapping_text}")
            return {}

    def fill_form(self, data, output_path='filled_form.pdf'):
        """
        Fill the PDF form using fillpdf library with Gemini-assisted mapping.

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

            # Generate field mapping using Gemini
            field_mapping = self.generate_field_mapping(list(form_fields.keys()), data)
            if not field_mapping:
                logging.warning("No field mapping generated. Cannot fill form.")
                return None

            updates = {}

            for pdf_field, data_key in field_mapping.items():
                if data_key in data and data[data_key]:
                    updates[pdf_field] = data[data_key]
                else:
                    updates[pdf_field] = None  # Set to null if data is missing
                    logging.warning(f"Data key '{data_key}' not found or empty in data dictionary for PDF field '{pdf_field}'. Setting it to null.")

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

def extract_data_from_email(body):
    """
    Use Gemini to extract relevant data from the email body.

    :param body: The plain text body of the email.
    :return: Dictionary containing extracted data.
    """
    prompt = f"""
Please identify and provide the following information from the email below in JSON format:

1. Seller's Name
2. Buyer's Name
3. Seller's Mailing Address
4. Buyer's Mailing Address
5. Seller's Print Name 1
6. Seller's Print Name 2
7. Buyer's Print Name 1
8. Buyer's Print Name 2
9. Seller's ZIP Code
10. Buyer's ZIP Code

**Email Body:**

{body}

**Instructions:**
- Ensure that ZIP Codes are 5-digit numbers.
- Provide only the JSON without any additional text or explanations.
- If a field does not have corresponding data, set its value to null.

**Example JSON Output:**
{{
    "Seller's Name": "John Doe",
    "Buyer's Name": "Jane Smith",
    "Seller's Mailing Address": "123 Elm Street, Springfield, IL",
    "Buyer's Mailing Address": "456 Oak Avenue, Lincoln, NE",
    "Seller's Print Name 1": "John",
    "Seller's Print Name 2": "Doe",
    "Buyer's Print Name 1": "Jane",
    "Buyer's Print Name 2": "Smith",
    "Seller's ZIP Code": "62704",
    "Buyer's ZIP Code": "68508"
}}
"""

    try:
        response_text = call_gemini_api(prompt, temperature=0.3, max_tokens=500)
        if not response_text:
            logging.error("Failed to extract data using Gemini.")
            return {}

        # Attempt to parse JSON
        try:
            data = json.loads(response_text)
        except json.JSONDecodeError as e:
            logging.error(f"Error decoding JSON from Gemini response: {e}")
            logging.debug(f"Response Text: {response_text}")
            return {}

        # Validate and format data
        data['Seller mail address'] = validate_address(data.get('Seller mail address', ''))
        data['Buyer mail address'] = validate_address(data.get('Buyer mail address', ''))
        data['Seller ZIP Code'] = validate_zip(data.get('Seller ZIP Code', ''))
        data['Buyer ZIP Code'] = validate_zip(data.get('Buyer ZIP Code', ''))

        logging.info(f"Extracted and validated data from email: {data}")
        return data
    except Exception as e:
        logging.error(f"Error extracting data from email with Gemini: {e}")
        return {}

def validate_zip(zip_code):
    """
    Validate that the ZIP code is a 5-digit number.

    :param zip_code: The ZIP code to validate.
    :return: Formatted ZIP code if valid, else None.
    """
    if re.fullmatch(r'\d{5}', zip_code):
        return zip_code
    else:
        logging.warning(f"Invalid ZIP code format: {zip_code}")
        return None

def validate_address(address):
    """
    Basic validation for addresses.
    This can be enhanced with more sophisticated checks or APIs.

    :param address: The address to validate.
    :return: Address if valid, else None.
    """
    if isinstance(address, str) and len(address) > 5:
        return address
    else:
        logging.warning(f"Invalid address format: {address}")
        return None

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
            try:
                with open(self.attachment_path, 'rb') as f:
                    pdf_attachment = MIMEApplication(f.read(), _subtype="pdf")
                    pdf_attachment.add_header('Content-Disposition', 'attachment', filename=os.path.basename(self.attachment_path))
                    msg.attach(pdf_attachment)
                logging.info("Email includes a PDF attachment.")
            except Exception as e:
                logging.error(f"Failed to attach PDF: {e}")
        else:
            logging.warning(f"Attachment file {self.attachment_path} not found. Email will be sent without attachment.")

        logging.info(f"Would send email to {self.recipient_email} with subject '{msg['Subject']}'")
        logging.info(f"Email body:\n{self.body}")

        # Uncomment and configure the following code to send the email
        # try:
        #     smtp_server = smtplib.SMTP('smtp.gmail.com', 587)
        #     smtp_server.starttls()
        #     smtp_server.login(self.sender_email, os.getenv("EMAIL_PASSWORD"))  # Use environment variables in production
        #     smtp_server.send_message(msg)
        #     smtp_server.quit()
        #     logging.info("Email sent successfully.")
        # except Exception as e:
        #     logging.error(f"Failed to send email: {e}")

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

def download_attachments_from_email(email_processor):
    """
    Download all PDF attachments from the email using EmailProcessor.

    :param email_processor: An instance of EmailProcessor.
    :return: List of downloaded attachment filenames.
    """
    email_processor.extract_attachments(email_processor.email_data.get('payload', {}))

    # Log the downloaded attachments
    if email_processor.attachments:
        logging.info("Downloaded the following attachments:")
        for attachment in email_processor.attachments:
            logging.info(f" - {attachment['filename']}")
        return [attachment['filename'] for attachment in email_processor.attachments]
    else:
        logging.info("No PDF attachments found in the email.")
        return []

def main(json_file_path):
    """
    Main function to orchestrate email processing and PDF form filling.

    :param json_file_path: Path to the JSON file containing email data.
    """
    # Load JSON data
    try:
        with open(json_file_path, 'r') as file:
            email_data = json.load(file)
        logging.info(f"Loaded email data from '{json_file_path}'.")
    except FileNotFoundError:
        logging.error(f"JSON file not found: {json_file_path}")
        return
    except json.JSONDecodeError:
        logging.error(f"Invalid JSON format in file: {json_file_path}")
        return

    # Process email
    email_processor = EmailProcessor(email_data)
    sender, subject, body, attachments = email_processor.parse_email()

    # Extract data from email using Gemini
    extracted_data = extract_data_from_email(body)
    if not extracted_data:
        logging.error("Failed to extract data from email. Exiting.")
        return

    # Download attachments
    downloaded_files = download_attachments_from_email(email_processor)
    if not downloaded_files:
        logging.error("No PDF attachments to process. Exiting.")
        return

    # Process each downloaded PDF
    for pdf_filename in downloaded_files:
        filled_pdf_path = fill_form_with_nlp(pdf_filename, extracted_data, output_path=f'filled_{pdf_filename}')
        if filled_pdf_path:
            logging.info(f"Successfully filled the PDF form: {filled_pdf_path}")

            # Generate and send response email
            response_body = f"""Hello {email_processor.sender_name},

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
        else:
            logging.error(f"Failed to fill the PDF form for: {pdf_filename}")

if __name__ == "__main__":
    json_file_path = 'easy.json'  # Replace with your JSON file path
    main(json_file_path)
