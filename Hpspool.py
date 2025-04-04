import os
from pathlib import Path


def view_hpspool_file(file_path):
    """
    Display the contents of an HPSPOOL file in a readable format.

    Args:
        file_path (str): Path to the HPSPOOL file
    """
    # Convert to Path object for better path handling
    path = Path(file_path)

    # Check if file exists
    if not path.exists():
        raise FileNotFoundError(f"File not found: {file_path}")

    print(f"\nContents of {path.name}:")
    print("-" * 80)

    try:
        # Read and display the file contents
        with open(path, 'r', encoding='utf-8', errors='ignore') as f:
            lines = f.readlines()

        # Process and display each line
        for i, line in enumerate(lines, 1):
            # Remove form feed characters and other control sequences
            clean_line = line.strip('\f\n\r')

            # Skip empty lines
            if clean_line:
                # Add line numbers for reference
                print(f"{i:4d} | {clean_line}")

    except Exception as e:
        print(f"Error reading file: {str(e)}")

    print("-" * 80)
    print(f"Total lines: {len(lines)}")


# Example usage
if __name__ == "__main__":
    # Example with raw string path
    file_path = r"C:\Users\**\ETAT DES PROVISIONS MATHEMATIQUES.hpspool"

    try:
        view_hpspool_file(file_path)
    except Exception as e:
        print(f"Error: {str(e)}")
