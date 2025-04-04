import pandas as pd
import re
from pathlib import Path


def convert_hpspool_to_excel(input_file, output_file):
    """
    Convert an HPSPOOL file to Excel format.

    Args:
        input_file (str): Path to the input HPSPOOL file
        output_file (str): Path to save the output Excel file
    """
    # Convert strings to Path objects
    input_path = Path(input_file)
    output_path = Path(output_file)

    # Ensure output file has .xlsx extension
    if output_path.suffix.lower() != '.xlsx':
        output_path = output_path.with_suffix('.xlsx')
        print(f"Adding .xlsx extension to output file: {output_path}")

    # Check if input file exists
    if not input_path.exists():
        raise FileNotFoundError(f"Input file not found: {input_path}")

    # Read the HPSPOOL file
    with open(input_path, 'r', encoding='utf-8', errors='ignore') as f:
        lines = f.readlines()

    # Remove form feed characters and other printer control sequences
    cleaned_lines = []
    for line in lines:
        # Remove form feed and other control characters
        cleaned_line = re.sub(r'\x0c|\x1b\[[0-9;]*[a-zA-Z]', '', line)
        if cleaned_line.strip():  # Only keep non-empty lines
            cleaned_lines.append(cleaned_line.strip())

    # Parse the data into a structured format
    data = []
    current_row = []

    for line in cleaned_lines:
        # Split line by whitespace (adjust the splitting logic based on your file format)
        fields = [field for field in line.split() if field.strip()]

        # Check if this is a new row or continuation
        if fields and not line.startswith(' '):
            if current_row:
                data.append(current_row)
            current_row = fields
        else:
            current_row.extend(fields)

    # Add the last row if it exists
    if current_row:
        data.append(current_row)

    # Check if we have any data
    if not data:
        raise ValueError("No data found in input file")

    # Convert to DataFrame
    df = pd.DataFrame(data)

    # If the first row contains headers, use it as column names
    if len(df) > 0:
        df.columns = [f'Column_{i + 1}' for i in range(len(df.columns))]

    # Create output directory if it doesn't exist
    output_path.parent.mkdir(parents=True, exist_ok=True)

    # Save to Excel with explicit engine specification
    df.to_excel(output_path, index=False, engine='openpyxl')

    return df


# Example usage
if __name__ == "__main__":
    try:
        # Make sure your input file path is correct
        input_path = r"C:\Users\**\ETAT DES PROVISIONS MATHEMATIQUES.hpspool"

        # The output path must end with .xlsx
        output_path = r"C:\Users\**\ETAT DES PROVISIONS MATHEMATIQUES.xlsx"

        df = convert_hpspool_to_excel(input_path, output_path)
        print(f"Successfully converted {input_path} to {output_path}")
        print(f"Number of rows processed: {len(df)}")

    except Exception as e:
        print(f"Error converting file: {str(e)}")
