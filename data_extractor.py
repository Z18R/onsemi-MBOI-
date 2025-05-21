from Data import parse_notepad_files  
from SqlHandler import get_folder_path_by_id  

try:
    input_folder = get_folder_path_by_id(3)  
    backup_folder = get_folder_path_by_id(4)  
except Exception as e:
    print(f"Error fetching folder paths: {e}")
    exit(1)  

parse_notepad_files(input_folder, backup_folder)