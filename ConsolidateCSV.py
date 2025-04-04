import os
import pandas as pd
import glob
from typing import Optional


def parse_accounting_file(file: str, encoding: str = 'cp1252') -> Optional[pd.DataFrame]:
    """
    Parses accounting integration files with the specific structure:
    Date, Reference, Account codes, Description, Amount, etc.

    Args:
        file (str): Path to CSV file
        encoding (str): File encoding (default: cp1252 for French Windows files)

    Returns:
        Optional[pd.DataFrame]: Parsed DataFrame
    """
    try:
        # Define the column structure based on the sample provided
        columns = [
            'DATE', 'SPACE1', 'REFERENCE', 'AFFAIRE', 'COMPTE', 'CODE',
            'DESCRIPTION', 'PERIODE', 'SPACE2', 'CODE_JOURNAL',
            'NUM_PIECE', 'MONTANT', 'CODE_TIERS', 'NUM_OPERATION',
            'TYPE_MVT', 'STATUT'
        ]

        # Read the file with fixed-width or delimited format
        # Adjust the separator based on your actual file format
        df = pd.read_csv(file, encoding=encoding, names=columns,
                         delimiter=' ', skipinitialspace=True)

        # Clean up the data
        # Remove any empty columns
        df = df.drop(columns=[col for col in df.columns if col.startswith('SPACE')])

        # Convert date column to datetime
        df['DATE'] = pd.to_datetime(df['DATE'], format='%d/%m/%Y', errors='coerce')

        # Clean amount column (remove any thousand separators and convert to numeric)
        df['MONTANT'] = df['MONTANT'].str.replace(',', '').astype(float)

        # Add source file information
        df['SOURCE_FILE'] = os.path.basename(file)
        df['TRIMESTRE'] = df['SOURCE_FILE'].str.extract(r'(\dT\d{2})')[0]

        return df

    except Exception as e:
        print(f"Error parsing {file}: {str(e)}")
        return None


def consolidate_accounting_files(path: str) -> pd.DataFrame:
    """
    Consolidates multiple accounting integration files from a directory.

    Args:
        path (str): Directory path containing the files

    Returns:
        pd.DataFrame: Consolidated DataFrame
    """
    original_dir = os.getcwd()
    try:
        os.chdir(path)
        consolidated_df = pd.DataFrame()

        # Get integration files only
        files = [f for f in glob.glob('*IntegrationComptable.csv')]

        if not files:
            raise ValueError(f"No integration files found in {path}")

        print(f"Found {len(files)} integration files")

        for file in files:
            print(f"Processing {file}...")
            current_df = parse_accounting_file(file)

            if current_df is None:
                continue

            if consolidated_df.empty:
                consolidated_df = current_df
                print(f"First file processed: {file}")
            else:
                consolidated_df = pd.concat([consolidated_df, current_df],
                                            ignore_index=True)

            print(f"Successfully processed {file}")

        return consolidated_df

    finally:
        os.chdir(original_dir)


def save_consolidated_file(consolidated_df: pd.DataFrame, output_path: str,
                           filename: str = "consolidated_accounting.csv") -> None:
    """
    Saves the consolidated DataFrame with proper formatting.

    Args:
        consolidated_df (pd.DataFrame): Consolidated DataFrame
        output_path (str): Path to save the output file
        filename (str): Name of the output file
    """
    if not consolidated_df.empty:
        # Sort by date and reference
        consolidated_df = consolidated_df.sort_values(['DATE', 'REFERENCE'])

        full_path = os.path.join(output_path, filename)
        consolidated_df.to_csv(full_path, index=False, encoding='cp1252')

        print(f"Consolidated file saved to {full_path}")
        print(f"Total number of entries: {len(consolidated_df)}")
        print("\nSample of consolidated data:")
        print(consolidated_df.head())
    else:
        print("No data to save")


if __name__ == "__main__":
    input_path = r"C:\Users\**\Fichiers intégration"
    output_path = r"C:\Users\**\Fichiers intégration\PQTH"

    try:
        result_df = consolidate_accounting_files(input_path)
        save_consolidated_file(result_df, output_path)
    except Exception as e:
        print(f"Error: {str(e)}")
        raise
