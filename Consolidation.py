import os
import pandas as pd
import glob

def consolidate_excel_files(path):
    """
    Consolidates multiple Excel files from a directory into a single DataFrame,
    matching columns across files regardless of their order.

    Args:
        path (str): Directory path containing Excel files.

    Returns:
        pd.DataFrame: Consolidated DataFrame.
    """
    os.chdir(path)

    # Initialize empty DataFrame
    consolidated_df = pd.DataFrame()

    # Get all Excel files
    excel_files = glob.glob('*.xlsx')

    if not excel_files:
        raise ValueError(f"No Excel files found in {path}")

    print(f"Found {len(excel_files)} Excel files")

    # Process each file
    for file in excel_files:
        try:
            print(f"Processing {file}...")
            current_df = pd.read_excel(file)

            # Add source file column
            current_df['Source_File'] = file

            if consolidated_df.empty:
                consolidated_df = current_df
                print(f"First file processed: {file}")
            else:
                # Identify common columns
                common_cols = list(set(consolidated_df.columns) & set(current_df.columns))

                if not common_cols:
                    print(f"Warning: No matching columns found in {file}")
                    continue

                # Ensure missing columns in current_df are filled with None
                for col in consolidated_df.columns:
                    if col not in current_df.columns:
                        current_df[col] = None

                # Reorder columns to match consolidated_df
                current_df = current_df[consolidated_df.columns]

                # Append to consolidated DataFrame
                consolidated_df = pd.concat([consolidated_df, current_df], ignore_index=True)

            print(f"Successfully processed {file}")

        except Exception as e:
            print(f"Error processing {file}: {str(e)}")
            continue

    return consolidated_df


def save_consolidated_file(consolidated_df, output_path, filename="consolidated.xlsx"):
    """
    Saves the consolidated DataFrame to an Excel file.

    Args:
        consolidated_df (pd.DataFrame): Consolidated DataFrame.
        output_path (str): Path to save the output file.
        filename (str): Name of the output file.
    """
    if not consolidated_df.empty:
        full_path = os.path.join(output_path, filename)
        consolidated_df.to_excel(full_path, index=False)
        print(f"Consolidated file saved to {full_path}")
    else:
        print("No data to save")


if __name__ == "__main__":
    input_path = r"C:\Users\**\Bases réass\path"
    output_path = r"C:\Users\**\Bases réass\path"

    try:
        # Consolidate files
        result_df = consolidate_excel_files(input_path)

        # Save consolidated file
        save_consolidated_file(result_df, output_path)

    except Exception as e:
        print(f"Error: {str(e)}")
