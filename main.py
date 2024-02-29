import pandas as pd
import argparse, sys, os, csv
from datetime import datetime


class ExcelTableExtractor:
    def __init__(self, file_path, sheet_name):
        self.file_path = file_path
        self.sheet_name = sheet_name if sheet_name else "Sheet1"
        self.df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)

    def clean_dataframe(self):
        self.df = self.df.dropna(axis=1, how='all')

    def find_keyword_positions(self, df, keyword):
        positions = []
        for col in df.columns:
            positions.extend(df[df[col] == keyword].index.tolist())
        return positions

    def extract_tables(self, df=None):
        if df is None:
            df = self.df
        keywords_first_table = ("Pay Code", "Timecard Details")
        keywords_second_table = ("Date In", "Total")

        positions_first_table = [self.find_keyword_positions(df, key) for key in keywords_first_table]
        positions_second_table = [self.find_keyword_positions(df, key) for key in keywords_second_table]

        print(f"Positions of first table keywords: {positions_first_table}")
        print(f"Positions of second table keywords: {positions_second_table}")

        if not positions_first_table[0] or not positions_first_table[1]:
            print("Keywords for the first table not found.")
            raise ValueError("Keywords for the first table not found in the sheet.")
        if not positions_second_table[0] or not positions_second_table[1]:
            print("Keywords for the first table not found.")
            raise ValueError("Keywords for the second table not found in the sheet.")

        start_pos_first, end_pos_first = min(positions_first_table[0]), max(positions_first_table[1])
        table_summary = df.loc[start_pos_first:end_pos_first + 1].dropna(how='all')

        start_pos_second, end_pos_second = min(positions_second_table[0]), max(positions_second_table[1])
        table_details = df.loc[start_pos_second:end_pos_second + 1].dropna(how='all')

        # Clean the tables by removing columns with only NaN, NULL, or NaT values
        table_summary = table_summary.iloc[:, [0, 13]]
        table_summary.columns = ['PAYCODE', 'HOURS']
        table_summary = table_summary.iloc[1:-2]

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
        table_details.columns = ['DAY', 'DATE', 'TIMESTAMP', 'HOURS', 'DAY_TOTALS', 'PAY_CODE', 'OUT_TYPE',
                                 'WORKED_DEP']
        table_details = table_details.iloc[1:-1]

        return table_summary, table_details

    def find_rows_starting_with_total(self, df, keyword):
        """Find rows where any of the string columns start with a specific keyword."""
        string_columns = df.select_dtypes(include=[object, 'string'])
        return df[string_columns.apply(lambda x: x.str.startswith(keyword, na=False)).any(axis=1)]

    def extract_additional_info(self, df):
        """Extract additional information like Company Code, Date Range, and File Number."""
        additional_info = {}

        # Iterate through the DataFrame rows and extract information
        for _, row in df.iterrows():
            row_string = ' '.join(row.dropna().astype(str))
            if 'Company Code:' in row_string:
                additional_info['Company Code'] = row_string.split('Company Code:')[1].split()[0]
            if 'Date Range:' in row_string:
                additional_info['Date Range'] = ' '.join(row_string.split('Date Range:')[1].split()[:3])
            if 'File Number:' in row_string:
                additional_info['File Number'] = int(float(row_string.split('File Number:')[1].split()[0]))

        return additional_info

    def process_timecard_block(self, start, end):
        timecard_df = self.df.loc[start:end + 1]
        try:
            table_summary, table_details = self.extract_tables(timecard_df)
            table_totals = self.find_rows_starting_with_total(timecard_df, "Total")  # Find totals for this block
            table_totals = table_totals.dropna(axis=1, how='all')
            table_totals.columns = ['', 'HOURS_1', 'HOURS_2']
            table_totals['HOURS'] = table_totals['HOURS_1'].combine_first(table_totals['HOURS_2'])
            table_totals.drop(['HOURS_2', 'HOURS_1'], axis=1, inplace=True)

            additional_info = self.extract_additional_info(timecard_df)
            return table_summary, table_details, table_totals, additional_info
        except ValueError as e:
            print(f"Error processing block: {e}")
            return None, None, None, None  # or handle the error as needed


def generate_timecard_csv(additional_info, table_summary, table_details, output_path, block_index):
    """Generate a Timecard header CSV """
    try:
        # Prepare data for CSV
        companyCode = additional_info.get('Company Code', '')
        fileNumber = f"000000{additional_info.get('File Number', '')}"[-6:]
        employeeId = f"{companyCode}{fileNumber}"

        # Extract and format dates
        date_range = additional_info.get('Date Range', '').split(' - ')
        date_from = datetime.strptime(date_range[0], '%m/%d/%Y').strftime('%Y-%m-%d') if len(date_range) > 0 else ''
        date_to = datetime.strptime(date_range[1], '%m/%d/%Y').strftime('%Y-%m-%d') if len(date_range) > 1 else ''

        # Create a unique label for each timecard block
        timecard_label = f"{employeeId}_{date_from.replace('-', '')}_{date_to.replace('-', '')}_{fileNumber}_{block_index}"
        print(f"Generating timecard with Label {timecard_label}")
        csv_data = {
            "company": companyCode,
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
            csv_data[f"summary[{i}].paycode"] = row['EQUIV']
            csv_data[f"summary[{i}].hours"] = row['HOURS']

        output_file = f"{output_path}/headers/Timecard_Header_{timecard_label.replace('/', '-')}.csv"
        # Write to CSV
        with open(output_file, 'w', newline='') as csvfile:
            writer = csv.DictWriter(csvfile, fieldnames=csv_data.keys())
            writer.writeheader()
            writer.writerow(csv_data)

        # Generate timecard details CSV
        generate_timecard_details_csv(timecard_label, table_details, output_path)

    except Exception as e:
        print(f"Error generating CSV: {e}")


def generate_timecard_details_csv(timecard_label, table_details, output_path):
    # Prepare data for CSV
    csv_rows = []
    notes = ""  # Initialize notes as an empty string

    for _, row in table_details.iterrows():
        try:
            # Format DATE and TIMESTAMP into datetime strings
            datetime_in = datetime.combine(row['DATE'], datetime.strptime(row['TIMESTAMP'].split(' - ')[0],
                                                                          '%I:%M %p').time()).strftime('%Y-%m-%d %H:%M')
            datetime_out = datetime.combine(row['DATE'], datetime.strptime(row['TIMESTAMP'].split(' - ')[1],
                                                                           '%I:%M %p').time()).strftime(
                '%Y-%m-%d %H:%M')
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


def find_timecard_blocks(df):
    timecard_blocks = []
    start = None
    for i, row in df.iterrows():
        row_str = row.to_string()
        if "Timecard Detail Report with Signature:" in row_str:
            start = i
            # Check if the end marker is also in the same row
            if "Prepared On:" in row_str:
                end = i
                timecard_blocks.append((start, end))
                start = None  # Reset start for the next block
        elif start is not None and "Prepared On:" in row_str:
            end = i
            timecard_blocks.append((start, end))
            start = None  # Reset start for the next block

    return timecard_blocks


def main():
    try:
        parser = argparse.ArgumentParser(description="Extract tables from an Excel file.")
        parser.add_argument("file_path", help="Path to the Excel file")
        parser.add_argument("sheet_name", nargs='?', default="Sheet1", help="Name of the sheet to extract tables from")
        args = parser.parse_args()

        if not os.path.exists(args.file_path):
            raise FileNotFoundError(f"The file '{args.file_path}' does not exist.")

        extractor = ExcelTableExtractor(args.file_path, args.sheet_name)
        print(f"Extracting tables...\n{extractor.df.head()}")
        timecard_blocks = find_timecard_blocks(extractor.df)
        print(f"Found {len(timecard_blocks)} timecard blocks")
        print(timecard_blocks)

        for index, (start, end) in enumerate(timecard_blocks):
            print(f"\n\nProcessing block from {start} to {end}")
            table_summary, table_details, table_totals, additional_info = extractor.process_timecard_block(start, end)
            if table_summary is None:
                print(f"Skipping block from {start} to {end} due to errors")
                continue

            # Extract additional information
            # additional_info = extractor.extract_additional_info()
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
            path = './generated_csv'
            generate_timecard_csv(additional_info, table_summary, table_details, path, index)

    except FileNotFoundError as e:
        print(f"File Error: {e}")
    except ValueError as e:
        print(f"Value Error: {e}")
    except Exception as e:  # General exception for other errors, including invalid file format
        print(f"An error occurred: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
