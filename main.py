import pandas as pd
import argparse, sys, os, csv
from datetime import datetime


class ExcelTableExtractor:
    def __init__(self, file_path, sheet_name):
        self.file_path = file_path
        self.sheet_name = sheet_name if sheet_name else "Sheet1"  # Set default to "Sheet1" if empty
        self.df = pd.read_excel(file_path, sheet_name=sheet_name)

    def clean_dataframe(self):
        """Remove columns with only NaN, NULL, or NaT values."""
        self.df = self.df.dropna(axis=1, how='all')

    def find_keyword_positions(self, keyword):
        """Find the positions of a keyword in the DataFrame."""
        positions = []
        for col in self.df.columns:
            positions.extend(self.df[self.df[col] == keyword].index.tolist())
        return positions

    def extract_table(self, start_pos, end_pos):
        """Extract table from the DataFrame between start and end positions."""
        table = self.df.iloc[start_pos:end_pos+1]
        # Drop rows where all columns are NaN
        table = table.dropna(how='all')
        return table

    def find_rows_starting_with_total(self, keyword):
        """Find rows where any of the string columns start with a specific keyword."""
        string_columns = self.df.select_dtypes(include=[object, 'string'])
        return self.df[string_columns.apply(lambda x: x.str.startswith(keyword, na=False)).any(axis=1)]

    def extract_tables(self):
        # Keywords for the first and second table
        keywords_first_table = ("Pay Code", "Timecard Details")
        keywords_second_table = ("Date In", "Total")

        # Find positions of these keywords
        positions_first_table = [self.find_keyword_positions(key) for key in keywords_first_table]
        positions_second_table = [self.find_keyword_positions(key) for key in keywords_second_table]

        # Check if keywords are found in the DataFrame
        if not positions_first_table[0] or not positions_first_table[1]:
            raise ValueError("Keywords for the first table not found in the sheet.")
        if not positions_second_table[0] or not positions_second_table[1]:
            raise ValueError("Keywords for the second table not found in the sheet.")

        # Extract the first table
        start_pos_first, end_pos_first = min(positions_first_table[0]), max(positions_first_table[1])
        table_summary = self.extract_table(start_pos_first, end_pos_first)

        # Extract the second table
        start_pos_second, end_pos_second = min(positions_second_table[0]), max(positions_second_table[1])
        table_details = self.extract_table(start_pos_second, end_pos_second)

        # Extract rows starting with "Total"
        table_totals = self.find_rows_starting_with_total("Total")

        # Clean the tables by removing columns with only NaN, NULL, or NaT values
        table_summary = table_summary.dropna(axis=1, how='all')
        table_summary.columns = ['PAYCODE', 'HOURS']
        table_summary = table_summary.iloc[1:-2]

        # Add equivalence for pay codes
        # Define the equivalence mappings
        paycode_mappings = {
            '6th Consecutive Day Overtime': '6OT',
            'Overtime': 'OT',
            'Paid Time Off': 'PTO',
            'PAID UNION LUNCH': 'PUL',
            'Regular': 'REG'
        }

        # Add the equivalence column
        table_summary['EQUIV'] = table_summary['PAYCODE'].map(paycode_mappings)

        table_details = table_details.iloc[:, [0, 1, 5, 11, 15, 16, 21, 31]]
        table_details.columns = ['DAY', 'DATE', 'TIMESTAMP', 'HOURS', 'DAY_TOTALS', 'PAY_CODE', 'OUT_TYPE', 'WORKED_DEP']
        table_details = table_details.iloc[1:-1]

        table_totals = table_totals.dropna(axis=1, how='all')
        table_totals.columns = ['', 'HOURS_1', 'HOURS_2']
        table_totals['HOURS'] = table_totals['HOURS_1'].combine_first(table_totals['HOURS_2'])
        table_totals.drop(['HOURS_2', 'HOURS_1'], axis=1, inplace=True)

        return table_summary, table_details, table_totals

    def extract_additional_info(self):
        """Extract additional information like Company Code, Date Range, and File Number."""
        additional_info = {}

        # Iterate through the DataFrame rows and extract information
        for _, row in self.df.iterrows():
            row_string = ' '.join(row.dropna().astype(str))
            if 'Company Code:' in row_string:
                additional_info['Company Code'] = row_string.split('Company Code:')[1].split()[0]
            if 'Date Range:' in row_string:
                additional_info['Date Range'] = ' '.join(row_string.split('Date Range:')[1].split()[:3])
            if 'File Number:' in row_string:
                additional_info['File Number'] = int(float(row_string.split('File Number:')[1].split()[0]))

        return additional_info


def generate_timecard_csv(additional_info, table_summary, table_details, output_path):
    """Generate a Timecard header CSV """
    # Prepare data for CSV
    companyCode = additional_info.get('Company Code', '')
    fileNumber = f"000000{additional_info.get('File Number', '')}"[-6:]
    employeeId = f"{companyCode}{fileNumber}"

    # Extract and format dates
    date_range = additional_info.get('Date Range', '').split(' - ')
    date_from = datetime.strptime(date_range[0], '%m/%d/%Y').strftime('%Y-%m-%d') if len(date_range) > 0 else ''
    date_to = datetime.strptime(date_range[1], '%m/%d/%Y').strftime('%Y-%m-%d') if len(date_range) > 1 else ''
    timecard_label = f"{employeeId} - {date_from.replace('-', '/')} - {date_to.replace('-', '/')}"

    csv_data = {
        "company.companyCode": companyCode,
        "employee": employeeId,
        "dateFrom": date_from,
        "dateTo": date_to,
        "supervisor": "",  # Assuming supervisor's name is fixed
        "totalHs": table_summary['HOURS'].sum(),
        "timecardLabel": timecard_label
    }

    # Add summary data
    for i in range(len(table_summary)):
        row = table_summary.iloc[i]
        csv_data[f"sumary[{i}].paycode"] = row['EQUIV']
        csv_data[f"sumary[{i}].hours"] = row['HOURS']

    output_file = f"{output_path}/headers/Timecard_Header_{timecard_label.replace('/','-')}.csv"
    # Write to CSV
    with open(output_file, 'w', newline='') as csvfile:
        writer = csv.DictWriter(csvfile, fieldnames=csv_data.keys())
        writer.writeheader()
        writer.writerow(csv_data)

    # Generate timecard details CSV
    generate_timecard_details_csv(timecard_label, table_details, output_path)


def generate_timecard_details_csv(timecard_label, table_details, output_path):
    # Prepare data for CSV
    csv_rows = []
    notes = ""  # Initialize notes as an empty string

    for _, row in table_details.iterrows():
        try:
            # Format DATE and TIMESTAMP into datetime strings
            datetime_in = datetime.combine(row['DATE'], datetime.strptime(row['TIMESTAMP'].split(' - ')[0], '%I:%M %p').time()).strftime('%Y-%m-%d %H:%M')
            datetime_out = datetime.combine(row['DATE'], datetime.strptime(row['TIMESTAMP'].split(' - ')[1], '%I:%M %p').time()).strftime('%Y-%m-%d %H:%M')
        except Exception as e:
            # Handle the case where TIMESTAMP is not in the expected format
            print(f"Skipping row due to formatting error: {e}. Adding information as comment in the next row")
            notes = str(row.iloc[0])  # Save the first element of the row as a note if there's an error
            continue  # Skip this row

        # Handle nan values for dailyTotals and payCode
        daily_totals = row['DAY_TOTALS'] if pd.notna(row['DAY_TOTALS']) else 0
        pay_code = row['PAY_CODE'] if pd.notna(row['PAY_CODE']) else ''
        out_type = row['OUT_TYPE'] if pd.notna(row['OUT_TYPE']) else ''

        csv_row = {
            "timecard": timecard_label,
            "datetimeIn": datetime_in,
            "datetimeOut": datetime_out,
            "workedHours": row['HOURS'],
            "dailyTotals": daily_totals,
            "payCode": pay_code,
            "outType": out_type,
            "workedDepID": row['WORKED_DEP'],
            "notes": notes
        }
        notes = ""  # Clean notes

        csv_rows.append(csv_row)

    output_file = f"{output_path}/details/Timecard_Details_{timecard_label.replace('/', '-')}.csv"
    # Write to CSV
    with open(output_file, 'w', newline='') as csvfile:
        writer = csv.DictWriter(csvfile, fieldnames=csv_rows[0].keys())
        writer.writeheader()
        writer.writerows(csv_rows)


def main():
    try:
        # Define arguments
        parser = argparse.ArgumentParser(description="Extract tables from an Excel file.")
        parser.add_argument("file_path", help="Path to the Excel file")
        parser.add_argument("sheet_name", nargs='?', default="Sheet1",
                            help="Name of the sheet to extract tables from (default: Sheet1)")
        args = parser.parse_args()

        # Check if file exists
        if not os.path.exists(args.file_path):
            raise FileNotFoundError(f"The file '{args.file_path}' does not exist.")

        # Extract information from Sheets
        extractor = ExcelTableExtractor(args.file_path, args.sheet_name)
        table_summary, table_details, table_totals = extractor.extract_tables()

        # Extract additional information
        additional_info = extractor.extract_additional_info()
        print("----- Timecard Information: -----")
        print(additional_info)

        # Print the tables and total rows
        print("\n----- Summary -----")
        print(table_summary)
        print("\n----- Details -----")
        print(table_details)
        print("\nTotals")
        print(table_totals)

        # Generate CSV
        path = './generated_csv'  # Specify your desired output file path
        generate_timecard_csv(additional_info, table_summary, table_details, path)

    except FileNotFoundError as e:
        print(f"File Error: {e}")
    except ValueError as e:
        print(f"Value Error: {e}")
    except Exception as e:  # General exception for other errors, including invalid file format
        print(f"An error occurred: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
