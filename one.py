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
            logging.info(f"Attachment '{filename}' saved successfully.")
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
            logging.error(f"Error downloading attachment from {url}: {e}")
            return None

class NLPPDFProcessor:
    def __init__(self, pdf_path=None):
        self.pdf_path = pdf_path
        if self.pdf_path:
            logging.info(f"Initializing NLPPDFProcessor with file: {self.pdf_path}")
            if not os.path.exists(self.pdf_path):
                logging.warning(f"File {self.pdf_path} does not exist.")

    def analyze_pdf(self):
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
        if not self.pdf_path or not os.path.exists(self.pdf_path):
            logging.error(f"No valid PDF file found at {self.pdf_path}. Cannot fill form.")
            return None
        try:
            form_fields = self.analyze_pdf()
            if not form_fields:
                logging.warning("No form fields to fill.")
                return None

            updates = {}
            print(list(form_fields.keys()), list(data.keys()))
            matched_fields = self.match_fields_with_gemini(list(form_fields.keys()), list(data.keys()))

            for pdf_field, data_key in matched_fields.items():
                if data_key in data:
                    updates[pdf_field] = data[data_key]

            if updates:
                logging.info(f"Updating the following fields: {updates}")
                fillpdfs.write_fillable_pdf(self.pdf_path, output_path, updates)
                logging.info(f"Filled PDF saved as '{output_path}'.")
                return output_path
            else:
                logging.warning("No matching fields found to update.")
                return None
        except Exception as e:
            logging.error(f"Error filling PDF form: {e}")
            return None

    def match_fields_with_gemini(self, pdf_fields, data_keys):
        prompt = f"""
        I have a PDF form with the following fields:
        {json.dumps(pdf_fields, indent=2)}

        And I have data with the following keys:
        {json.dumps(data_keys, indent=2)}

        Please match each PDF form field with the most appropriate data key from the data provided.
        Return the result as a JSON object where each key is a PDF form field and its value is the corresponding data key.
        If there's no suitable match for a field, use null.

        **Only output the JSON object and nothing else. Do not include any Markdown or code block formatting.**
        """

        response = call_gemini_api(prompt)

        # Write the full response to a separate JSON file
        self.write_response_to_file(response)

        # Strip code block markers if present
        cleaned_response = self.strip_code_blocks(response)

        # Write the cleaned response for verification
        # self.write_cleaned_response_to_file(cleaned_response)

        try:
            # Attempt to parse the cleaned response directly
            matched_fields = json.loads(cleaned_response)
            logging.info("Field matching completed using Gemini API.")
            return matched_fields
        except json.JSONDecodeError:
            logging.error("Failed to parse Gemini API response as JSON after stripping code blocks.")

            # Attempt to extract JSON from the response using regex
            json_match = re.search(r'\{.*\}', cleaned_response, re.DOTALL)
            if json_match:
                try:
                    matched_fields = json.loads(json_match.group())
                    logging.info("Field matching completed using Gemini API after extracting JSON.")
                    return matched_fields
                except json.JSONDecodeError:
                    logging.error("Extracted JSON from Gemini API response is invalid.")

            logging.info("Falling back to fuzzy matching.")
            return self.fallback_fuzzy_matching(pdf_fields, data_keys)

    def strip_code_blocks(self, text):
        """
        Extracts JSON content from a code block if present.
        Removes ```json and ``` markers.
        """
        pattern = r'```json\s*\n?(.*?)\n?```'
        match = re.search(pattern, text, re.DOTALL)
        if match:
            json_text = match.group(1).strip()
            logging.debug("Code block markers found and removed.")
            return json_text
        else:
            logging.debug("No code block markers found.")
            return text.strip()


    def write_response_to_file(self, response):
        """
        Writes the raw Gemini API response to a JSON file with a timestamp.
        """
        try:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"gemini_response_{timestamp}.json"
            file_path = os.path.join(API_RESPONSES_DIR, filename)

            with open(file_path, 'w', encoding='utf-8') as f:
                f.write(response)

            logging.info(f"Gemini API response written to '{file_path}'.")
        except Exception as e:
            logging.error(f"Failed to write Gemini API response to file: {e}")

    def write_cleaned_response_to_file(self, cleaned_response):
        """
        Writes the cleaned Gemini API response to a separate JSON file for verification.
        """
        try:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"gemini_cleaned_response_{timestamp}.json"
            file_path = os.path.join(API_RESPONSES_DIR, filename)

            with open(file_path, 'w', encoding='utf-8') as f:
                f.write(cleaned_response)

            logging.info(f"Cleaned Gemini API response written to '{file_path}'.")
        except Exception as e:
            logging.error(f"Failed to write cleaned Gemini API response to file: {e}")



    def fallback_fuzzy_matching(self, pdf_fields, data_keys):
        matched_fields = {}
        for field in pdf_fields:
            best_match = max(data_keys, key=lambda x: fuzz.ratio(field.lower(), x.lower()))
            if fuzz.ratio(field.lower(), best_match.lower()) > 70:
                matched_fields[field] = best_match
            else:
                matched_fields[field] = None
        return matched_fields

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

import os
import json
import requests
import logging

def call_gemini_api(prompt, temperature=0.2, max_output_tokens=500):
    api_key = os.environ.get('GEMINI_API_KEY')
    if not api_key:
        logging.error("GEMINI_API_KEY environment variable is not set")
        return ""

    api_url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-pro:generateContent"
    
    headers = {
        "Content-Type": "application/json"
    }
    
    data = {
        "contents": [{"parts":[{"text": prompt}]}],
        "generationConfig": {
            "temperature": temperature,
            "maxOutputTokens": max_output_tokens,
        }
    }
    
    params = {
        "key": api_key
    }

    try:
        response = requests.post(api_url, headers=headers, params=params, json=data)
        response.raise_for_status()
        
        # Log request details without exposing the full API key
        logging.info(f"API call to {api_url} successful. Status code: {response.status_code}")
        
        result = response.json()
        
        if 'candidates' in result and result['candidates']:
            generated_text = result['candidates'][0]['content']['parts'][0]['text']
            # Log a snippet of the response for debugging
            logging.debug(f"API response snippet: {generated_text[:100]}...")
            return generated_text
        else:
            logging.error(f"Unexpected response structure: {result}")
            return ""
    except requests.exceptions.RequestException as e:
        logging.error(f"Error calling Gemini API: {e}")
        if hasattr(e, 'response') and e.response is not None:
            logging.error(f"Response content: {e.response.text}")
        return ""


def download_attachments_from_email(email_processor):
    email_processor.extract_attachments(email_processor.email_data.get('payload', {}))

    if email_processor.attachments:
        logging.info("Downloaded the following attachments:")
        for attachment in email_processor.attachments:
            logging.info(f" - {attachment['filename']}")
        return [attachment['filename'] for attachment in email_processor.attachments]
    else:
        logging.info("No PDF attachments found in the email.")
        return []

def main(json_file_path):
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

    email_processor = EmailProcessor(email_data)
    sender, subject, body, attachments = email_processor.parse_email()

    downloaded_files = download_attachments_from_email(email_processor)
    if not downloaded_files:
        logging.error("No PDF attachments to process. Exiting.")
        return
    try:
        with open('data.json', 'r') as file:
            user_data = json.load(file)
        logging.info("Loaded form data from 'data.json'.")
    except FileNotFoundError:
        logging.error("Form data file 'data.json' not found.")
        return
    except json.JSONDecodeError:
        logging.error("Invalid JSON format in 'data.json'.")
        return


    data = user_data
    for pdf_filename in downloaded_files:
        processor = NLPPDFProcessor(pdf_filename)
        filled_pdf_path = processor.fill_form(data, output_path=f'filled_{pdf_filename}')
        if filled_pdf_path:
            logging.info(f"Successfully filled the PDF form: {filled_pdf_path}")

            response_body = f"""Hello,

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
