import fitz  # PyMuPDF
import re

def extract_text_from_pdf(pdf_path):
    doc = fitz.open(pdf_path)
    text = ""
    for page in doc:
        text += page.get_text("text") + "\n"  # Extract text from each page
    return text

def process_extracted_text(text):
    # Regular expressions to extract the desired fields
    mboi_pattern = r"(Manufacturing Bill of Information \(MBOI\))"
    part_number_pattern = r"Part Number\s*([A-Za-z0-9\-]+)"
    part_type_pattern = r"Part Type\s*([A-Za-z\s]+)"
    part_version_pattern = r"Part Version\s*([\d]+)"
    part_version_date_pattern = r"Part Version Date-Time\s*([\d\-]+ [A-Za-z]+ [\d:]+ \(GMT [\+\-\d]+\))"
    
    # Search for the patterns in the text
    mboi_match = re.search(mboi_pattern, text)
    part_number_match = re.search(part_number_pattern, text)
    part_type_match = re.search(part_type_pattern, text)
    part_version_match = re.search(part_version_pattern, text)
    part_version_date_match = re.search(part_version_date_pattern, text)
    
    # Extract and format the information
    result = ""
    
    if mboi_match:
        result += mboi_match.group(1) + "\n"
    
    if part_number_match:
        result += f"Part Number: {part_number_match.group(1)}\n"
    
    if part_type_match:
        result += f"Part Type: {part_type_match.group(1)}\n"
    
    if part_version_match:
        result += f"Part Version: {part_version_match.group(1)}\n"
    
    if part_version_date_match:
        result += f"Part Version Date-Time: {part_version_date_match.group(1)}\n"
    
    return result

if __name__ == "__main__":
    # pdf_file = "C:\\Users\\jc_it\\Desktop\\ONSEMI Dumping (Green_Honey)\\MBOI\\PH4_MC74HCU04ADTR2G_99MC74HCU04ADTG_ASY_ver.1.pdf"
    pdf_file = "C:\\Users\\jc_it\\Documents\python\\Mboi Extraction\\PH4_74VHC02MTCX-IPN_74VHC02MTCX-IPN-ASY_ASY.pdf"

    extracted_text = extract_text_from_pdf(pdf_file)
    
    formatted_text = process_extracted_text(extracted_text)
    
    with open("formatted_mboi.txt", "w", encoding="utf-8") as file:
        file.write(formatted_text)
    
    print("Extraction complete! Check 'formatted_mboi.txt' for results.")
