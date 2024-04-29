import pandas as pd

# Reading Excel file
path = r"sample_data.xlsx"
df = pd.read_excel(path)

# Replacing the null salesman column with a name
df['SalesMan'].fillna('Rohan', inplace=True)


def convert_date(date):
    """Replacing all cells with dates from “DD/MM/YY” format to “DD/MM/YYYY” format"""
    try:
        return pd.to_datetime(date, format='%m/%d/%y').strftime('%m/%d/%Y')
    except ValueError:
        return date


df['OrderDate'] = df['OrderDate'].apply(convert_date)

# Writing Excel file
new_file_path = r"new_updated_data.xlsx"
df.to_excel(new_file_path, index=False)
