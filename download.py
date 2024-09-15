import json
import base64
import requests
import os
import logging

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

class EmailProcessor:
    def __init__(self, email_data):
        """
        Initialize the EmailProcessor with email data.
        
        :param email_data: Dictionary containing the email data.
        """
        self.email_data = email_data
        self.attachments = []

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

def download_attachments_from_email(json_file_path):
    """
    Load email data from a JSON file and download all PDF attachments.
    
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

    # Initialize EmailProcessor and extract attachments
    email_processor = EmailProcessor(email_data)
    payload = email_data.get('payload', {})
    email_processor.extract_attachments(payload)

    # Log the downloaded attachments
    if email_processor.attachments:
        logging.info("Downloaded the following attachments:")
        for attachment in email_processor.attachments:
            logging.info(f" - {attachment['filename']}")
    else:
        logging.info("No PDF attachments found in the email.")

if __name__ == "__main__":
    json_file_path = 'easy.json'  # Replace with your actual JSON file path
    download_attachments_from_email(json_file_path)
