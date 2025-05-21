import os
import re
import shutil
from openpyxl import Workbook
from datetime import datetime
from SqlHandler import connect_to_database  # Import the database connection function

def parse_notepad_files(input_folder, backup_folder):
    # Connect to the database
    conn = connect_to_database()
    if not conn:
        print("Exiting program due to database connection failure.")
        return  # Exit if the database connection fails

    cursor = conn.cursor()

    # Check if there are any files matching the pattern
    notepad_files = [f for f in os.listdir(input_folder) if f.startswith("notepad") and f.endswith(".txt")]
    if not notepad_files:
        print("No notepad files found. Exiting program.")
        return  # Exit the function if no files are found

    # Create a new workbook for all extracted data
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Extracted Data"

    # Write headers
    sheet.append([
        "Orderable Part Section", "Part Number", "Description", "Primary Die", "Piece Parts",
        "Part NumberF", "DescriptionF", "Part NumberE", "DescriptionE",
        "Part NumberW", "DescriptionW", "Part NumberM", "DescriptionM",
        "Revision History (Date)", "Revision History (Details)", "Package Kit"
    ])

    # Get the current date and time for file naming
    now = datetime.now()
    process_date = now.strftime("%m%d%Y")
    timestamp = now.strftime("%m%d%S%Y")  # Format for the output file name

    # Process all files matching the pattern
    for file_name in notepad_files:
        file_path = os.path.join(input_folder, file_name)
        with open(file_path, 'r', encoding='utf-8') as file:
            content = file.read()

        # Clean up special characters
        content = re.sub(r'[^\x00-\x7F]+', ' ', content)

        # Extract Orderable Part Section
        orderable_part_section_match = re.search(r"Orderable Part Section\s+-\s+([^\n]+)", content)
        orderable_part_section = orderable_part_section_match.group(1).strip() if orderable_part_section_match else "Orderable Part Section Not Found"

        # Extract Part Number and Description
        part_number_desc_match = re.search(r"Part Number \| Desc\s+([^\t\n]+)\s+\|\s+([^\t\n]+)", content)
        if part_number_desc_match:
            part_number = part_number_desc_match.group(1).strip()
            description = part_number_desc_match.group(2).strip()
        else:
            part_number = "Part Number Not Found"
            description = "Description Not Found"

        # Extract all Alternate sections dynamically
        alternates = re.findall(r"(Alternate \d{3}).*?(BOI 67-Bonding Diagram|$)", content, re.DOTALL)
        if not alternates:
            print("No Alternate sections found. Skipping file.")
            continue

        for alternate, _ in alternates:
            print(f"Processing {alternate}...")

            # Extract BOM Components under the current Alternate
            bom_components = re.search(rf"{alternate}.*?BOI 67-Bonding Diagram", content, re.DOTALL)
            if bom_components:
                bom_section = bom_components.group(0)

                # Extract all Primary Die entries
                primary_dies = re.findall(r"Primary Die\s+([^\t\n]+)", bom_section)
                if not primary_dies:
                    print(f"No Primary Dies found in {alternate}.")
                    continue

                # Extract all Package Kits dynamically
                package_kits = re.findall(r"Package Kit\s+([^\s]+)\s+PKT", bom_section)
                if not package_kits:
                    print(f"No Package Kits found in {alternate}.")
                    continue

                # Loop through each Primary Die and associate it with all Package Kits
                for primary_die in primary_dies:
                    for package_kit in package_kits:
                        # Extract Piece Parts specific to this Package Kit
                        package_kit_section = re.search(
                            rf"Package Kit\s+{package_kit}.*?Piece Parts\s+(.*?)\s+(Package Kit|BOI|$)",
                            bom_section, re.DOTALL
                        )
                        if package_kit_section:
                            piece_parts_section = package_kit_section.group(1)
                            # Extract individual piece parts (including S or - in identifiers)
                            piece_parts = re.findall(r"(\d{10,}[S|-]?[A-Z0-9]*)\s+([A-Z]+)", piece_parts_section)
                        else:
                            piece_parts = []

                        # Sort Piece Parts in the order FRAME, EPXY, WIRE, MOLD
                        piece_parts_dict = {desc: num for num, desc in piece_parts}
                        sorted_piece_parts = [
                            (piece_parts_dict.get("FRAME", None), "FRAME"),
                            (piece_parts_dict.get("EPXY", None), "EPXY"),
                            (piece_parts_dict.get("WIRE", None), "WIRE"),
                            (piece_parts_dict.get("MOLD", None), "MOLD")
                        ]

                        # Extract Revision History (latest entry)
                        revision_history_match = re.search(
                            r"Revision History\s+Rev\s+Rev Date\s+Details\s+(.*?)(?=\n\n|\Z)", content, re.DOTALL
                        )
                        if revision_history_match:
                            revision_history_section = revision_history_match.group(1).strip()

                            # Extract all revisions
                            revisions = re.findall(
                                r"(\d+)\s+(\d{4}-[A-Z]{3}-\d{2} [\d:]+)\s+(.*?)(?=\n\d+|\Z)", revision_history_section, re.DOTALL
                            )
                            if revisions:
                                # Get the latest revision (highest Rev number)
                                latest_revision = max(revisions, key=lambda x: int(x[0]))
                                revision_date = latest_revision[1].strip()  # Rev Date
                                revision_details = latest_revision[2].strip()  # Details
                            else:
                                revision_date = "Date Not Found"
                                revision_details = "Details Not Found"
                        else:
                            revision_date = "Date Not Found"
                            revision_details = "Details Not Found"

                        # Write data in a single row
                        row = [
                            orderable_part_section,  # Orderable Part Section
                            part_number,  # Part Number (dynamic)
                            description,  # Description (dynamic)
                            primary_die,  # Primary Die
                            "",  # Placeholder for "Piece Parts" column
                            sorted_piece_parts[0][0], "FRAME",  # Part NumberF, DescriptionF
                            sorted_piece_parts[1][0], "EPXY",  # Part NumberE, DescriptionE
                            sorted_piece_parts[2][0], "WIRE",  # Part NumberW, DescriptionW
                            sorted_piece_parts[3][0], "MOLD",  # Part NumberM, DescriptionM
                            revision_date,  # Revision History (Date)
                            revision_details,  # Revision History (Details)
                            package_kit  # Package Kit
                        ]
                        sheet.append(row)

                        try:
                            # Check if the record already exists
                            cursor.execute("""
                                SELECT COUNT(*) 
                                FROM cst_onsemi_MBOIExtractedData 
                                WHERE [Part Number] = ? AND [Package Kit] = ? AND [Primary Die] = ?
                            """, (part_number, package_kit, primary_die))
                            exists = cursor.fetchone()[0]

                            if exists > 0:
                                print(f"Record already exists for Part Number: {part_number}, Package Kit: {package_kit}, Primary Die: {primary_die}. Skipping insertion.")
                            else:
                                # Insert the record if it doesn't exist
                                cursor.execute("""
                                    INSERT INTO cst_onsemi_MBOIExtractedData (
                                        [Orderable Part Section], [Part Number], [Description], [Primary Die],
                                        [Piece Parts], [Part NumberF], [DescriptionF], [Part NumberE], [DescriptionE],
                                        [Part NumberW], [DescriptionW], [Part NumberM], [DescriptionM],
                                        [Revision History (Date)], [Revision History (Details)], [Package Kit]
                                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                                """, row)
                                conn.commit()
                                print(f"Inserted record for Part Number: {part_number}, Package Kit: {package_kit}, Primary Die: {primary_die}.")
                        except Exception as e:
                            print(f"Error inserting data into database: {e}")
            else:
                print(f"BOM Components not found for {alternate}. Skipping.")

        # Move the processed file to the backup folder with a new name
        if not os.path.exists(backup_folder):
            os.makedirs(backup_folder)
        new_file_name = f"process_{process_date}_{file_name}"
        shutil.move(file_path, os.path.join(backup_folder, new_file_name))

    # Save the Excel file with the new naming format
    output_file = os.path.join(input_folder, f"{timestamp}_extracted_data.xlsx")
    workbook.save(output_file)
    print(f"Data has been written to {output_file}")

    # Close the database connection
    conn.close()