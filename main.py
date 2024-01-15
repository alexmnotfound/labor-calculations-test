import pandas as pd
from datetime import datetime, date


class PayCalculator:
    def __init__(self, standard_rate, overtime_rate, doubletime_rate):
        self.standard_rate = standard_rate
        self.overtime_rate = overtime_rate
        self.doubletime_rate = doubletime_rate

    def calculate_hours(self, entry, exit):
        # Convert datetime.time to datetime.datetime for subtraction
        entry_datetime = datetime.combine(date.today(), entry)
        exit_datetime = datetime.combine(date.today(), exit)
        work_duration = exit_datetime - entry_datetime
        return work_duration.total_seconds() / 3600

    def calculate_daily_pay(self, entry_time, exit_time):
        hours_worked = self.calculate_hours(entry_time, exit_time)
        print(f"-- Hours worked: {round(hours_worked, 2)}")
        if hours_worked <= 8:
            return hours_worked * self.standard_rate
        elif hours_worked <= 12:
            print(f"-- Overtime: {round(hours_worked - 8, 2)}")
            return 8 * self.standard_rate + (hours_worked - 8) * self.overtime_rate
        else:
            return 8 * self.standard_rate + 4 * self.overtime_rate + (hours_worked - 12) * self.doubletime_rate

    def read_time_data(self, file_name):
        return pd.read_excel(file_name)

    def calculate_total_pay(self, file_name):
        total_pay = 0
        data = self.read_time_data(file_name)
        data.dropna(subset=["Entry Time", "Exit Time"], inplace=True)

        def calculate_row_pay(row):
            entry_time, exit_time = row['Entry Time'], row['Exit Time']
            day = row['Date'].replace("\n", " ")
            print(f"- Calculating Payment for {day}")
            return self.calculate_daily_pay(entry_time, exit_time)

        data['Daily Pay'] = data.apply(calculate_row_pay, axis=1)
        total_pay = data['Daily Pay'].sum()

        print("\n------------------------------\n")
        print("Summary:")

        for index, row in data.iterrows():
            day = row['Date'].replace("\n", " ")
            print(f"  {day}: {row['Daily Pay']:.2f} USD")

        print("\n------------------------------\n")
        print(f"Total Pay: {total_pay:.2f} USD")
        print("\n")


def main():
    # Constants
    standard_hourly_rate = 20  # Standard rate per hour
    overtime_rate = 1.5 * standard_hourly_rate  # 1.5 times the standard rate for overtime
    doubletime_rate = 2 * standard_hourly_rate  # Double rate for hours over 12 in a day

    # Show some info
    print("\n------- California labor calculations --------------\n")
    print("Standard rates:\n")
    print(f" Standard Hourly Rate:\n   {float(standard_hourly_rate)} USD")
    print(f" Overtime Rate (more than 8 hours pays 1.5 times the standard rate for overtime):\n   {float(overtime_rate)} USD")
    print(f" Doubletime Rate (more than 12 hours pays double rate):\n   {float(doubletime_rate)} USD\n")

    print(" Things to consider:\n"
          "  - Need to add weekends calculations?\n"
          "")
    print("\n------------------------------\n")

    # Example usage
    file_name = './files/timesheet.xlsx'  # Replace with your actual Excel file
    calculator = PayCalculator(standard_hourly_rate, overtime_rate, doubletime_rate)
    calculator.calculate_total_pay(file_name)


if __name__ == "__main__":
    main()
