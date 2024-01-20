import pandas as pd

# Load the Excel file into a DataFrame
file_path = 'Assignment_Timecard.xlsx'  # Update with the actual file path
df = pd.read_excel(file_path)

# Convert the 'Time' and 'Time Out' columns to datetime objects
df['Time'] = pd.to_datetime(df['Time'], errors='coerce')
df['Time Out'] = pd.to_datetime(df['Time Out'], errors='coerce')

# Drop rows with missing values in the 'Employee Name' column
df = df.dropna(subset=['Employee Name'])

# Convert 'Timecard Hours (as Time)' column to numeric, handling errors by coercing to NaN
df['Timecard Hours (as Time)'] = pd.to_numeric(df['Timecard Hours (as Time)'], errors='coerce')

# Sort the DataFrame by 'Employee Name' and 'Time'
df = df.sort_values(by=['Employee Name', 'Time'])

# Function to calculate the time difference between consecutive shifts
def calculate_time_difference(group):
    return group['Time'].diff().fillna(pd.Timedelta(seconds=0))

# Function to check conditions and print results
def analyze_file(df):
    for name, group in df.groupby('Employee Name'):
        consecutive_days = group['Time'].diff().dt.days == 1
        less_than_10_hours = calculate_time_difference(group).between('0:01:00', '10:00:00')
        more_than_14_hours = group['Timecard Hours (as Time)'] > 14

        if any(consecutive_days):
            print(f"{name} has worked for 7 consecutive days.")

        if any(less_than_10_hours):
            print(f"{name} has less than 10 hours between shifts but greater than 1 hour.")

        if any(more_than_14_hours):
            print(f"{name} has worked for more than 14 hours in a single shift.")

# Call the function to analyze the file
analyze_file(df)
