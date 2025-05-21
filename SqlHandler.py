import pyodbc

def connect_to_database():
    """Establish a connection to the SQL Server database."""
    connection_string = (
        "Driver={SQL Server};"
        "Server=MSDynamics-DB\\AXDB;"
        "Database=MES_ATEC;"
        "UID=sa;"
        "PWD=p@ssw0rd;"
        "TrustServerCertificate=Yes;"
    )
    try:
        conn = pyodbc.connect(connection_string)
        print("Database connection established.")
        return conn
    except pyodbc.Error as e:
        print(f"Error connecting to database: {e}")
        return None

def get_folder_path_by_id(folder_id):
    """Fetch the folder path for a specific ID from the database."""
    conn = connect_to_database()
    if not conn:
        raise Exception("Failed to connect to the database.")

    cursor = conn.cursor()
    try:
        # Query to fetch the folder path for the given ID
        cursor.execute("""
            SELECT TOP 1 folderPath 
            FROM TBL_BEPP_MATRIX
            WHERE ID = ? AND ACTIVE = 1
        """, folder_id)
        row = cursor.fetchone()
        if row:
            return row[0]  # Return the folder path
        else:
            raise Exception(f"No active folder path found for ID = {folder_id}.")
    finally:
        conn.close()