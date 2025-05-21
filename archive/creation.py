# import PyPDF2

# def extract_text_from_pdf(pdf_path):
#     try:
#         with open(pdf_path, 'rb') as pdf_file:
#             reader = PyPDF2.PdfReader(pdf_file)
#             text = ''
#             for page in reader.pages:
#                 text += page.extract_text()

#                 return text

#     except Exception as e:
#         return f"An error occurred: {e}"

# pdf_path = r'C:\Users\jc_it\Documents\python\Mboi Extraction\PH4_74LCX74MTCX-IPN_74LCX74MTCX-IPN-ASY_ASY.pdf'

# pdf_text = extract_text_from_pdf(pdf_path)
# print(pdf_text)



import pdfplumber

def extract_text_with_pdfplumber(pdf_path):
    try:
        with pdfplumber.open(pdf_path) as pdf:
            text = ''
            for page in pdf.pages:
                text += page.extract_text()
            return text
    except Exception as e:
        return f"An error occurred: {e}"

pdf_path = r'C:\Users\jc_it\Documents\python\Mboi Extraction\PH4_74LCX74MTCX-IPN_74LCX74MTCX-IPN-ASY_ASY.pdf'
pdf_text = extract_text_with_pdfplumber(pdf_path)
print(pdf_text)