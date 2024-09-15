from fillpdf import fillpdfs
import logging

# Configure logging to display information
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

def list_form_fields(pdf_path):
    try:
        form_fields = fillpdfs.get_form_fields(pdf_path)
        if form_fields:
            logging.info("PDF Form Fields:")
            for field in form_fields:
                logging.info(f" - {field}")
        else:
            logging.info("No form fields found in the PDF.")
    except Exception as e:
        logging.error(f"Error extracting form fields: {e}")

# Example usage
if __name__ == "__main__":
    pdf_path = 'easy-pdf.pdf'  # Replace with your PDF file path
    list_form_fields(pdf_path)
