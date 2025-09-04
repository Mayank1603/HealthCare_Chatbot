import pandas as pd
import sys
import json
import re
from docx import Document
from PIL import Image
import pytesseract
import PyPDF2

# Function to extract text from a PDF file
def read_pdf(file_path):
    text = ""
    try:
        with open(file_path, 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            for page in reader.pages:
                text += page.extract_text()
    except Exception as e:
        return f"Error reading PDF file: {e}"
    return text

# Function to extract text from a Word file
def read_word(file_path):
    text = ""
    try:
        doc = Document(file_path)
        for paragraph in doc.paragraphs:
            text += paragraph.text + "\n"
    except Exception as e:
        return f"Error reading Word file: {e}"
    return text

# Function to extract text from an image file
def read_image(file_path):
    text = ""
    try:
        image = Image.open(file_path)
        text = pytesseract.image_to_string(image)
    except Exception as e:
        return f"Error reading image file: {e}"
    return text

# Main function to handle PDF, Word, and image files
def read_file(file_path):
    if file_path.endswith('.pdf'):
        return read_pdf(file_path)
    elif file_path.endswith('.docx'):
        return read_word(file_path)
    elif file_path.endswith(('.png', '.jpg', '.jpeg', '.bmp', '.tiff')):
        return read_image(file_path)
    else:
        return "Unsupported file format. Please provide a PDF, Word, or image file."

# Function to extract numeric value (ignores ranges containing '-')
def extract_numeric_value(value):
    try:
        # Check if it's a range (contains '-')
        if '-' in value:
            return None
        match = re.search(r"[-+]?\d*\.\d+|\d+", value)
        if match:
            return float(match.group())
    except Exception as e:
        pass
    return None

# Function to categorize rows based on the first row containing column names (Test and Result only)
# Function to categorize rows based on the first row containing column names (Test and Result only)
def categorize_rows_based_on_columns(extracted_data):
    rows = extracted_data.splitlines()

    # Print the raw extracted data for debugging
    print(f"Extracted Data:\n{extracted_data}\n")
    
    # Find the row containing 'Test' which should be the header row
    header_row_index = None
    for i, row in enumerate(rows):
        if 'test' in row.lower():
            header_row_index = i
            break
    
    # If no header row is found, return an error
    if header_row_index is None:
        return "Error: Couldn't find the 'Test' column in the data."

    # Extract column names from the identified header row
    column_names = rows[header_row_index].split()  # Split the header row into columns
    
    # Print the column names for debugging
    print(f"Column Names: {column_names}\n")

    # Identify the index of "Test" and "Result" columns
    test_col_index = None
    result_col_index = None
    normal_col_index = None
    range_col_index = None
    for idx, col_name in enumerate(column_names):
        if 'test' in col_name.lower():
            test_col_index = idx
        elif 'result' in col_name.lower():
            result_col_index = idx
        elif 'normal' in col_name.lower():
            normal_col_index = idx
        elif 'range' in col_name.lower():
            range_col_index = idx
    
    # If any column is not found, return an error message
    if test_col_index is None or result_col_index is None or normal_col_index is None or range_col_index is None:
        return "Error: Couldn't find 'Test', 'Normal', 'Range' or 'Result' columns in the data."

    # Categorize the data starting from the row after the header row
    categorized_data = []
    for row in rows[header_row_index + 1:]:  # Start from the row after the header
        columns = row.split()  # Split each row into columns
        
        if len(columns) > max(test_col_index, result_col_index, normal_col_index, range_col_index):  # Ensure enough columns
            row_data = {
                "Test": columns[test_col_index],
                "Normal": columns[normal_col_index],
                "Range": columns[range_col_index],
                "Result": columns[result_col_index]
            }
            categorized_data.append(row_data)
    
    return categorized_data


# Function to write the extracted data into an Excel file
def write_to_excel(data, output_file):
    try:
        # Convert data to a pandas DataFrame
        df = pd.DataFrame(data)
        
        # Write the DataFrame to an Excel file
        df.to_excel(output_file, index=False)
        
        # Print the DataFrame to check its content
        print(f"Data written to {output_file} successfully!")
        print(df.head())  # Print the first few rows to inspect
        return f"Data successfully written to {output_file}"
    except Exception as e:
        return f"Error writing to Excel: {e}"

# Function to read back the Excel file and extract Test and Result columns
def extract_from_excel(excel_file):
    try:
        # Load the Excel file into a pandas DataFrame
        df = pd.read_excel(excel_file)
        
        # Print the entire DataFrame to check its contents
        print(f"Data extracted from {excel_file}:")
        print(df.head())  # Print the first few rows to inspect
        
        # Extract the Test and Result columns
        extracted_data = df[['Test', 'Result']]
        
        return extracted_data
    except Exception as e:
        return f"Error reading from Excel: {e}"

# Main function to process the uploaded file
if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(json.dumps({"error": "Please provide the file path as an argument."}, indent=4))
        sys.exit(1)

    file_path = sys.argv[1]
    output_file = 'extracted_medical_report_data.xlsx'  # Specify the output Excel file name

    # Extract text from the provided file
    content = read_file(file_path)

    if content.startswith("Error"):
        print(json.dumps({"error": content}, indent=4))
        sys.exit(1)

    # Categorize rows into a structured format based on the Test and Result columns
    categorized_data = categorize_rows_based_on_columns(content)

    if isinstance(categorized_data, str):  # Check if an error occurred while categorizing
        print(json.dumps({"error": categorized_data}, indent=4))
        sys.exit(1)

    # Write the categorized data to an Excel file
    result_message = write_to_excel(categorized_data, output_file)
    print(result_message)

    # Now, extract the Test and Result columns back from the Excel file
    extracted_data = extract_from_excel(output_file)

    if isinstance(extracted_data, pd.DataFrame):
        print("Extracted Data:")
        print(extracted_data)
    else:
        print(json.dumps({"error": extracted_data}, indent=4))
