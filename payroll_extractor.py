import pandas as pd
import sys, re, csv
from collections import defaultdict


class PayrollDataExtractor:
    def __init__(self, file_path, sheet_name):
        self.file_path = file_path
        self.sheet_name = sheet_name if sheet_name else "Sheet1"
        try:
            self.df = pd.read_excel(file_path, sheet_name=self.sheet_name, header=None)
        except (FileNotFoundError, IsADirectoryError, PermissionError) as e:
            print(f"Error opening file: {str(e)}")
            sys.exit(1)  # Exit the program
        except Exception as e:
            print(f"Unexpected error: {str(e)}")
            sys.exit(1)  # Exit the program

    def clean_dataframe(self):
        self.df.dropna(axis=0, how='all', inplace=True)
        self.df.dropna(axis=1, how='all', inplace=True)

    def extract_employee_data(self, start, end):
        payroll_df = self.df.loc[start:end + 1].copy()

        # Identifying rows that contain 'Associate ID' keyword
        payroll_df['is_associate'] = payroll_df[0].astype(str).str.contains('Associate ID')

        # Propagate the 'File Number' forward to get file number for all rows
        payroll_df['File Number'] = payroll_df.loc[payroll_df['is_associate'], 0].str.extract(r'File #: (\d+)',
                                                                                              expand=False)
        payroll_df['File Number'].ffill(inplace=True)  # Use ffill instead of fillna with method

        # Drop the column 'is_associate' as it's no longer needed
        payroll_df.drop(columns=['is_associate'], inplace=True)

        # Reordering columns to place 'File Number' first
        columns = list(payroll_df.columns)
        columns = [columns[-1]] + columns[:-1]
        payroll_df = payroll_df[columns]

        create_summary(payroll_df)

        return payroll_df


def create_summary(payroll_df):
    print("\n------ Summary ------")
    print(f"File Number: {payroll_df['File Number'].iloc[0]}")

    rate_str = payroll_df.iloc[0][0]
    match = re.search(r'Rate: (\d+\.\d+)', rate_str)
    rate = match.group(1) if match else "Not Found"
    print(f"Rate: {rate}")
    
    gross = payroll_df.iloc[0][8]
    print(f"Gross: {gross}")

    # Define Payroll Label
    payroll_label = f"EP{payroll_df['File Number'].iloc[0]} - 2023/10/02 - 2023/10/15"
    # Initialize CSV row data
    csv_data = {
        "company": "Case HM LLC",
        "employee": f"EP1{payroll_df['File Number'].iloc[0]}",
        "dateFrom": "2023-10-02",  # Assuming fixed
        "rate": rate,
        "gross": gross,
    }

    # Parse and sum the voluntary deductions
    voluntary_deductions = parse_and_sum_keyed_financial_values(payroll_df[11])
    # Iterate over the voluntary deductions and add them to the CSV data
    i = 0
    for key, total in voluntary_deductions.items():
        key = key.replace("\n", " ")
        print(f"Voluntary Deduction -- '{key}': {total}")
        csv_data[f"voluntaryDeductions[{i}].detail"] = key
        csv_data[f"voluntaryDeductions[{i}].amount"] = total
        i += 1

    # Parse and sum the voluntary deductions
    net_pay = parse_and_sum_keyed_financial_values(payroll_df[12])
    # Iterate over the voluntary deductions and add them to the CSV data
    for key, total in net_pay.items():
        key = key.replace("\n", " ")
        print(f"Net Pay -- '{key}': {total}")
        csv_data[f"netPay.detail"] = key
        csv_data[f"netPay.amount"] = total

    total_worked_hours = payroll_df.iloc[-1][1]
    total_worked_hours = extract_number_after_colon(total_worked_hours)
    print(f"Total Worked Hours: {total_worked_hours}")
    csv_data[f"totalHs"] = total_worked_hours

    # SUMMARY HOURS
    summary = dict()
    payroll_df[1] = pd.to_numeric(payroll_df[1], errors='coerce')
    regular_hours = payroll_df[1].sum()
    print(f"Total Regular Hours: {regular_hours}")

    payroll_df[4] = pd.to_numeric(payroll_df[4], errors='coerce')
    regular_earnings = payroll_df[4].sum()
    print(f"Total Regular Earnings: {regular_earnings}")
    summary["REG"] = {"hours": regular_hours, "total": regular_earnings}

    payroll_df[2] = pd.to_numeric(payroll_df[2], errors='coerce')
    overtime_hours = payroll_df[2].sum()
    print(f"Total Overtime Hours: {overtime_hours}")

    payroll_df[5] = pd.to_numeric(payroll_df[5], errors='coerce')
    overtime_earnings = payroll_df[5].sum()
    print(f"Total Overtime Earnings: {overtime_earnings}")
    summary["OT"] = {"hours": overtime_hours, "total": overtime_earnings}

    # Calculate the sums for each key in paycode hours
    paycode_hours_sums = parse_and_sum_values(payroll_df[3])
    for key, total in paycode_hours_sums.items():
        print(f"Total Hours for {key}: {total}")
        summary[key] = {"hours": total, "total": ""}

    # Calculate the sums for each key in paycode earnings
    paycode_earnings_sums = parse_and_sum_values(payroll_df[6])
    for key, total in paycode_earnings_sums.items():
        print(f"Total Earnings for {key}: {total}")
        if key not in summary:
            summary[key] = {}
        if "total" not in summary[key]:
            summary[key]["total"] = 0
        summary[key]["total"] = total

    i = 0
    for key, values in summary.items():
        csv_data[f"summary[{i}].paycode"] = key
        csv_data[f"summary[{i}].hours"] = values.get("hours", "")
        csv_data[f"summary[{i}].total"] = values.get("total", "")
        i += 1

    deductions = dict()
    # Calculate the sums for each key in federal taxes
    federal_taxes = parse_and_sum_values(payroll_df[9])
    for key, total in federal_taxes.items():
        print(f"Federal Tax rate {key}: {total}")
        deductions[key] = {"tax_type": "Federal", "total": total}


    # Calculate the sums for each key in local taxes
    local_taxes = parse_and_sum_values(payroll_df[10])
    for key, total in local_taxes.items():
        print(f"Local Tax rate {key}: {total}")
        deductions[key] = {"tax_type": "Local/State", "total": total}

    i = 0
    for key, values in deductions.items():
        csv_data[f"deductions[{i}].code"] = key
        csv_data[f"deductions[{i}].type"] = values["tax_type"]
        csv_data[f"deductions[{i}].rate"] = values["total"]
        i += 1

    csv_data[f"memos"] = ""
    csv_data[f"payrollLabel"] = payroll_label

    print("\n\n")


    # Write the CSV row
    with open(f"./generated_csv/payrolls/{payroll_label.replace('/', '-')}.csv", 'w', newline='') as csvfile:
        writer = csv.DictWriter(csvfile, fieldnames=csv_data.keys())
        writer.writeheader()
        writer.writerow(csv_data)


def extract_number_after_colon(s):
    # The pattern looks for any text followed by a colon, then spaces, and then a number
    match = re.search(r':\s*(\d+\.?\d*)', s)
    return float(match.group(1)) if match else None


def parse_and_sum_values(column):
    # Initialize a dictionary to hold the sums for each key
    sums = defaultdict(float)

    # Define a regex pattern to extract key-value pairs
    pattern = re.compile(r'(\b[A-Z0-9 ]+\b) (\d+\.\d+|\d+)')

    # Iterate over each cell in the column
    for cell in column:
        if not isinstance(cell, str):
            continue
        for match in pattern.finditer(cell):
            key, value = match.groups()
            sums[key] += float(value)

    return sums


def parse_and_sum_keyed_financial_values(column):
    sums = defaultdict(float)

    for cell in column:
        if not isinstance(cell, str):
            continue

        # Split the string into non-numeric and numeric parts
        parts = cell.rsplit(maxsplit=1)
        if len(parts) < 2:
            continue

        key, value_str = parts
        value_str = value_str.replace(',', '')  # Remove commas

        try:
            value = float(value_str)
            sums[key] += value
        except ValueError:
            # Handle the case where conversion to float fails
            continue

    return sums


def find_payroll_blocks(df):
    # Create a boolean series where True indicates the start of a new employee block
    new_block_starts = df[0].str.contains('Associate ID').fillna(False)

    # Find start indices of blocks
    start_indices = new_block_starts[new_block_starts].index

    # Generate tuples of (start, end) indices for each block
    payroll_blocks = [(start, end - 2) for start, end in zip(start_indices, start_indices[1:].append(pd.Index([len(df)])))]

    return payroll_blocks


def main():
    # update these lines with the actual path and sheet name or consider dynamic input
    file_path = '/home/mr/projects/labor-calculations-test/files/samples/quintuple-payroll.xlsx'
    sheet_name = '3_payrolls'

    payroll_extractor = PayrollDataExtractor(file_path, sheet_name)
    payroll_extractor.clean_dataframe()

    payroll_blocks = find_payroll_blocks(payroll_extractor.df)
    print(payroll_blocks)

    for index, (start, end) in enumerate(payroll_blocks):
        extracted_employee_data = payroll_extractor.extract_employee_data(start, end)
        print(extracted_employee_data)  # Display the first few rows


if __name__ == "__main__":
    main()
