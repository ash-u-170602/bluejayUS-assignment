import pandas as pd
from datetime import timedelta

def analyze_work_schedule(file_path):
    # Read the Excel file into a Pandas DataFrame
    df = pd.read_excel(file_path)

    # Convert 'Time' and 'Time Out' columns to datetime objects
    df['Time'] = pd.to_datetime(df['Time'])
    df['Time Out'] = pd.to_datetime(df['Time Out'])

    # Sort the DataFrame by 'Employee Name' and 'Time'
    df = df.sort_values(by=['Employee Name', 'Time'])

    # Initialize variables for consecutive days and shift duration
    consecutive_days_count = 0
    total_shift_hours = 0

    # Iterate through rows in the DataFrame
    for index, row in df.iterrows():
        # Check if the current row is consecutive to the previous row
        if index > 0:
            time_diff = row['Time'] - df.iloc[index - 1]['Time Out']
            if time_diff == timedelta(days=1):
                consecutive_days_count += 1
            else:
                consecutive_days_count = 0

        # Check conditions for analysis
        if consecutive_days_count == 6:
            print(f"{row['Employee Name']} has worked for 7 consecutive days. Position: {row['Position ID']}")

        if index > 0:
            hours_between_shifts = (row['Time'] - df.iloc[index - 1]['Time Out']).seconds / 3600
            if 1 < hours_between_shifts < 10:
                print(f"{row['Employee Name']} has less than 10 hours between shifts but greater than 1 hour. Position: {row['Position ID']}")

        shift_duration = (row['Time Out'] - row['Time']).seconds / 3600
        if shift_duration > 14:
            print(f"{row['Employee Name']} has worked for more than 14 hours in a single shift. Position: {row['Position ID']}")

    print("Analysis completed.")

# Assuming the input file is named 'work_schedule.xlsx'
file_path = 'Assignment_Timecard.xlsx'
analyze_work_schedule(file_path)
