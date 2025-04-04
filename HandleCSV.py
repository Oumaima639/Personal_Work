import pandas as pd
import glob
import os
import datetime

# Constants
log_file = "log_file.txt"
path = r"C:\Users\**\Données actuariat"

# Set working directory at the start
try:
    os.chdir(path)
    print(f"Working directory set to: {os.getcwd()}")
except Exception as e:
    print(f"Error setting working directory: {str(e)}")
    exit(1)


def extract_from_csv(file_to_process):
    try:
        print(f"Reading file: {file_to_process}")
        dataframe = pd.read_csv(file_to_process, encoding='ISO-8859-1')
        print(f"Successfully read {len(dataframe)} rows")
        return dataframe
    except Exception as e:
        print(f"Error reading file {file_to_process}: {str(e)}")
        return pd.DataFrame()


def extract():
    columns = ["SINISTRE", "DATSIN", "DATOUV", "CONT", "AG", "CAT",
               "USAGE", "RECOUR", "RESPS", "NATSIN", "CXP", "IDA_APL",
               "ProvGarantie", "ANNEEFICHIERRESERVE", "GARANTIE", "CLefCat",
               "Code branche", "Catégorie", "LIBELLE USAGE",
               "LibelleGarantieReserve", "DEFF"]

    extracted_data = pd.DataFrame(columns=columns)

    # List CSV files before processing
    csv_files = glob.glob("*.csv")
    print(f"Found {len(csv_files)} CSV files: {csv_files}")

    for csvfile in csv_files:
        df = extract_from_csv(csvfile)
        if not df.empty:
            extracted_data = pd.concat([extracted_data, df], ignore_index=True)

    print(f"Total rows extracted: {len(extracted_data)}")
    return extracted_data


def transform(data):
    """Reads the CLefCat n° and creates a new column containing the first digit that refers to the branch"""
    try:
        print(f"Starting transformation on {len(data)} rows")
        data['CLefCat'] = data.CLefCat.astype(str)
        data['Cbranche'] = data.CLefCat.str.extract(r'(\d)').fillna('')
        print(f"Transformation complete. Unique branch codes: {data['Cbranche'].unique()}")
        return data
    except Exception as e:
        print(f"Error during transformation: {str(e)}")
        raise


def load_data(transformed_data: pd.DataFrame):
    """Split data into Excel files by branch code and split large datasets into multiple sheets."""
    try:
        output_dir = "output_files"
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
            print(f"Created output directory: {output_dir}")

        max_rows_per_sheet = 1048576  # Excel's row limit
        branches = transformed_data['Cbranche'].unique()
        print(f"Processing {len(branches)} unique branches")

        for branch_code, branch_data in transformed_data.groupby('Cbranche'):
            if branch_code:  # Skip empty branch codes
                output_file = os.path.join(output_dir, f'branch_{branch_code}.xlsx')

                with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
                    for i in range(0, len(branch_data), max_rows_per_sheet):
                        part_data = branch_data[i:i + max_rows_per_sheet]
                        sheet_name = f'branch_{branch_code}_part_{i // max_rows_per_sheet + 1}'
                        print(f"Writing {len(part_data)} rows to {sheet_name} in {output_file}")
                        part_data.to_excel(writer, sheet_name=sheet_name, index=False)

                print(f"Data for branch {branch_code} saved to: {output_file}")

        print("Data split into separate workbooks and sheets successfully")
        print(f"Files are saved in: {os.path.abspath(output_dir)}")

    except Exception as e:
        print(f"Error saving files: {str(e)}")
        raise


def log_progress(message):
    try:
        timestamp_format = '%Y-%h-%d-%H:%M:%S'
        now = datetime.datetime.now()
        timestamp = now.strftime(timestamp_format)
        with open(log_file, "a") as f:
            f.write(f"{timestamp},{message}\n")
        print(message)  # Also print to console
    except Exception as e:
        print(f"Error logging progress: {str(e)}")


def main():
    try:
        log_progress("ETL Job Started")

        log_progress("Extract phase Started")
        extracted_data = extract()
        if extracted_data.empty:
            raise Exception("No data extracted")
        log_progress("Extract phase Ended")

        log_progress("Transform phase Started")
        transformed_data = transform(extracted_data)
        print("Transformed Data Shape:", transformed_data.shape)
        log_progress("Transform phase Ended")

        log_progress("Load phase Started")
        load_data(transformed_data)
        log_progress("Load phase Ended")

        log_progress("ETL Job Ended Successfully")

    except Exception as e:
        log_progress(f"ETL Job Failed: {str(e)}")
        raise


if __name__ == "__main__":
    main()
