import spacy
from rapidfuzz import fuzz
import logging
import os
from PyPDF2 import PdfReader
from fillpdf import fillpdfs
import subprocess
import sys
import requests  

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

class NLPPDFProcessor:
    def __init__(self, pdf_path=None):
        self.pdf_path = pdf_path
        if self.pdf_path:
            logging.info(f"Initializing NLPPDFProcessor with file: {self.pdf_path}")
            if not os.path.exists(self.pdf_path):
                logging.warning(f"File {self.pdf_path} does not exist.")

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

    def save_attachment(self, filename, content):
        try:
            with open(filename, 'wb') as file:
                file.write(content)
            logging.info(f"Attachment saved as: {filename}")
        except IOError as e:
            logging.error(f"Error saving attachment: {e}")



    def analyze_pdf(self):
        """
        Analyze the PDF and extract form field names using PyPDF2.
        Returns a dictionary with field names as keys.
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

            # Match data keys to PDF form fields
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

# Example usage
if __name__ == "__main__":
    pdf_path = 'easy-pdf.pdf'        # Path to your fillable PDF form
    output_path = 'filled_form.pdf'  # Desired output path

    data = {
        "Print seller's name": 'John Doe',
        "Printed Buyer's name": 'Jane Smith',
        "Seller mail address": '123 Seller St, SellerCity, SC',
        "Buyer mail address": '456 Buyer Ave, BuyerCity, BC',
        "Seller print name 1": 'John Doe',
        "Seller print name 2": 'John Doe',
        "Buyer print name 1": 'Jane Smith',
        "Buyer print name 2": 'Jane Smith',
        # Add more key-value pairs as needed based on detected PDF fields
    }

    filled_pdf_path = fill_form_with_nlp(pdf_path, data, output_path)
    if filled_pdf_path:
        logging.info(f"Successfully filled the PDF form: {filled_pdf_path}")
    else:
        logging.error("Failed to fill the PDF form.")
