import pdfrw

# Path to input PDF file (the form you want to fill out)
input_pdf_path = 'form_template.pdf'
# Path to save the filled-out PDF
output_pdf_path = 'filled_form.pdf'

# Data to fill in the PDF form
form_data = {
    'Name': 'John Doe',
    'Date of Birth': '01/01/1990',
    'Address': '123 Main St, Springfield, USA',
    'Phone Number': '(555) 123-4567',
    'Email': 'johndoe@example.com',
}

# Function to update PDF form fields
def fill_pdf(input_pdf, output_pdf, data):
    template_pdf = pdfrw.PdfReader(input_pdf)
    annotations = template_pdf.pages[0]['/Annots']
    
    for annotation in annotations:
        field = annotation.getObject()
        field_name = field['/T'][1:-1]  # Strip leading and trailing parentheses
        
        if field_name in data.keys():
            field.update(
                pdfrw.PdfDict(V='{}'.format(data[field_name]), AS='{}'.format(data[field_name]))
            )
    
    pdfrw.PdfWriter(output_pdf, trailer=template_pdf).write()

# Run the function to fill the PDF
fill_pdf(input_pdf_path, output_pdf_path, form_data)

print(f"Form filled and saved as {output_pdf_path}")
