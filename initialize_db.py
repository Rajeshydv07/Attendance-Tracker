import os
import pandas as pd

def initialize_database():
    file_path = "attendance.xlsx"
    
    # Check if file already exists to avoid overwriting existing data
    if os.path.exists(file_path):
        print(f"'{file_path}' already exists. Skipping initialization.")
        return
        
    print(f"Creating a new attendance database: {file_path}")
    
    # Columns matching the tracking needs
    # Column 1: Roll Number
    # Column 2: Student Email
    # Column 3: CI (Computer Intelligence) - Leaves Taken
    # Column 4: Python - Leaves Taken
    # Column 5: DM (Data Mining) - Leaves Taken
    data = {
        "Roll Number": [1, 2, 3, 4, 5],
        "Student Email": [
            "student1@example.com",
            "student2@example.com",
            "student3@example.com",
            "student4@example.com",
            "student5@example.com"
        ],
        "CI": [0, 1, 2, 0, 1],
        "Python": [0, 1, 2, 1, 2],
        "DM": [0, 1, 2, 2, 0]
    }
    
    df = pd.DataFrame(data)
    
    # Save using pandas and openpyxl
    df.to_excel(file_path, index=False, sheet_name="Sheet1")
    print("Database initialized successfully with mock student data!")

if __name__ == "__main__":
    initialize_database()
