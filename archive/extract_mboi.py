import fitz  # PyMuPDF

def extract_text_from_pdf(pdf_path):
    doc = fitz.open(pdf_path)
    text = ""
    for page in doc:
        text += page.get_text("text") + "\n"  # Extract text from each page
    return text

if __name__ == "__main__":
    pdf_file = "C:\\Users\\jc_it\\Documents\python\\Mboi Extraction\\PH4_74VHC02MTCX-IPN_74VHC02MTCX-IPN-ASY_ASY.pdf"
    # pdf_file = "C:\\Users\\jc_it\\Desktop\\ONSEMI Dumping (Green_Honey)\\MBOI\\PH4_MC74HC595ADR2G_MC74HC595ADR2G-ASY_ASY_Ver1.pdf"
    extracted_text = extract_text_from_pdf(pdf_file)
    
    with open("extracted_mboi.txt", "w", encoding="utf-8") as file:
        file.write(extracted_text)
     
    print("Extraction complete! Check 'extracted_mboi.txt' for results.")


# PH4_MC74HCU04ADTR2G_99MC74HCU04ADTG_ASY_ver.1.pdf