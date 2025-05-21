import warnings
warnings.filterwarnings("ignore", message="CropBox missing")  # Suppress CropBox warnings
import pdfplumber

def extract_mboi_data(pdf_path):
    print(f"Extracting data from: {pdf_path}\n")
    
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            
            # Check if this page contains the MBOI section
            if "Manufacturing Bill of Information (MBOI)" in text:
                print("----- Manufacturing Bill of Information (MBOI) -----")
                
                # Extract all tables on this page
                tables = page.extract_tables()
                
                # The MBOI table is usually the first one after the header
                if tables:
                    mboi_table = tables[0]  # First table = MBOI data
                    for row in mboi_table:
                        # Clean empty cells and print row
                        clean_row = [str(cell).strip() for cell in row if cell]
                        print(" | ".join(clean_row))
                break  # Exit after processing MBOI

# Your PDF file path
pdf_file = "C:\\Users\\jc_it\\Documents\\python\\Mboi Extraction\\PH4_74LCX244MTCX-IPN_74LCX244MTCX-IPN-ASY_ASY.pdf"

# Run extraction
extract_mboi_data(pdf_file)