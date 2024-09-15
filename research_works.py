from fillpdf import fillpdfs
import logging
import os

# Configure logging
logging.basicConfig(
    level=logging.INFO,  # Set to INFO or DEBUG for more detailed logs
    format='%(asctime)s - %(levelname)s - %(message)s'
)

def fill_single_field_with_fillpdf(input_pdf, output_pdf, field_name, field_value):
    if not os.path.exists(input_pdf):
        logging.error(f"Input PDF '{input_pdf}' does not exist.")
        return False

    data_dict = {
        field_name: field_value
    }

    try:
        fillpdfs.write_fillable_pdf(input_pdf, output_pdf, data_dict)
        logging.info(f"Filled PDF saved as '{output_pdf}'.")
        return True
    except Exception as e:
        logging.error(f"Error filling PDF form: {e}")
        return False

# Example usage
if __name__ == "__main__":
    input_pdf = 'easy-pdf.pdf'          # Path to your fillable PDF form
    output_pdf = 'filled_form_fillpdf.pdf'  # Desired output path
    field_name = "Print seller's name"  # Exact field name
    field_value = 'John Doe'            # Value to insert

    success = fill_single_field_with_fillpdf(input_pdf, output_pdf, field_name, field_value)
    if success:
        logging.info("PDF form filled successfully using fillpdf.")
    else:
        logging.error("Failed to fill PDF form using fillpdf.")
