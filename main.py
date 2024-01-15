import pandas as pd
from datetime import datetime, date


class Employee:
    def __init__(self, standard_rate):
        self.standard_rate = standard_rate

    def calculate_hours(self, entry, exit):
        entry_datetime = datetime.combine(date.today(), entry)
        exit_datetime = datetime.combine(date.today(), exit)
        work_duration = exit_datetime - entry_datetime
        return work_duration.total_seconds() / 3600

    def calculate_daily_pay(self, entry_time, exit_time):
        raise NotImplementedError("This method should be implemented by subclasses.")


class HourlyEmployee(Employee):
    def __init__(self, standard_rate, overtime_rate, doubletime_rate):
        super().__init__(standard_rate)
        self.overtime_rate = overtime_rate
        self.doubletime_rate = doubletime_rate

    def calculate_daily_pay(self, entry_time, exit_time):
        hours_worked = self.calculate_hours(entry_time, exit_time)
        print(f"  - Hours worked: {round(hours_worked, 2)}")

        if hours_worked <= 8:
            return hours_worked * self.standard_rate
        elif hours_worked <= 12:
            print(f"  - Overtime: {round(hours_worked - 8, 2)}")
            return 8 * self.standard_rate + (hours_worked - 8) * self.overtime_rate
        else:
            print(f"  - Doubletime: {round((hours_worked - 12), 2)}")
            return 8 * self.standard_rate + 4 * self.overtime_rate + (hours_worked - 12) * self.doubletime_rate


class PayCalculator:
    def __init__(self, employee):
        self.employee = employee

    def read_time_data(self, file_name):
        return pd.read_excel(file_name)

    def calculate_total_pay(self, file_name):
        total_pay = 0
        data = self.read_time_data(file_name)
        data.dropna(subset=["Entry Time", "Exit Time"], inplace=True)

        for index, row in data.iterrows():
            current_day = row['Date'].replace("\n", " ")
            print(f"- Calculating Payment for {current_day}")

            entry_time, exit_time = row['Entry Time'], row['Exit Time']
            daily_pay = self.employee.calculate_daily_pay(entry_time, exit_time)
            print(f"  - Total Payment {round(daily_pay, 2)} USD")

            total_pay += daily_pay
        return total_pay


def main():
    # Constants
    standard_hourly_rate = 20  # Standard rate per hour
    overtime_rate = 1.5 * standard_hourly_rate  # 1.5 times the standard rate for overtime
    doubletime_rate = 2 * standard_hourly_rate  # Double rate for hours over 12 in a day

    # Show some info
    print("\n---- California labor calculations -------\n")
    print("Standard rates:\n")
    print(f" Standard Hourly Rate:\n   {float(standard_hourly_rate)} USD")
    print(f" Overtime Rate (more than 8 hours pays 1.5 times the standard rate for overtime):\n   {float(overtime_rate)} USD")
    print(f" Doubletime Rate (more than 12 hours pays double rate):\n   {float(doubletime_rate)} USD\n")

    print(" Things to consider:\n"
          "  - Need to add weekends calculations?\n"
          "")
    print("\n---------------\n")

    # Initialize an HourlyEmployee instance
    hourly_employee = HourlyEmployee(standard_hourly_rate, overtime_rate, doubletime_rate)

    # Initialize PayCalculator with the employee instance
    calculator = PayCalculator(hourly_employee)

    # Example usage
    file_name = './files/timesheet.xlsx'
    total_pay = calculator.calculate_total_pay(file_name)

    print("\n---------------\n")
    print(f"Total Pay: {total_pay:.2f} USD")


if __name__ == "__main__":
    main()
