
import pandas as pd
from pathlib import Path
import os

# Read the Excel files
input_path = Path(input("Enter the path of your Excel file: "))
excel_file = pd.ExcelFile(input_path)

# Specify the sheet names
sheet_name_1 = ("GL_Code")
sheet_name_2 = ("epf_raw data")

# Read the sheets into DataFrames
df1 = excel_file.parse(sheet_name_1)
df2 = excel_file.parse(sheet_name_2)

# Hard-coded branch data
branch_data_values = {
    'Work Location': ['Delhi', 'Chennai', 'Pune', 'Mumbai', 'Bengaluru'],
    'Branch': ['01', '01', '05', '07', '03']
}
branch_data = pd.DataFrame(branch_data_values)

# Columns to sum for salary calculation
salary_columns_to_sum = [
    "Employer's PF Contribution","Employer's EPS Contribution"
]
salary_columns_to_sum2 = [
    "Employer's EDLI Contribution","Employer's PF Admin Charges","Employer's PF Contribution","Employer's EPS Contribution"
]

# Initialize an empty DataFrame to store the final result
result_df = pd.DataFrame(columns=["GL Code", "Debit", "Credit", "Branch", "Department", "Emp Code", "Dimension Exists","External Doc"])
external_doc_value = input("Enter the value for External Doc: ")

# Iterate through the unique values in 'Zoho Heads' column of df1
for value in df1['Zoho Heads'].unique():
    # Filter rows in df1 based on the current value
    rows_matching_value = df1[df1['Zoho Heads'] == value]

    if value == "EPF Employer Contribution":
        dimension_value = rows_matching_value['Dimension'].iloc[0]
        sign = rows_matching_value['Sign'].iloc[0]

        if dimension_value not in ['B', 'BD', 'BDE']:
            print(f"Unsupported Dimension: {dimension_value}")
            continue

        # Group and sum salary columns based on dimension
        grouped_data = df2.groupby(['Work Location', 'Department', 'Employee ID'][:len(dimension_value)])[
            salary_columns_to_sum].sum().reset_index()
        grouped_data = grouped_data.merge(branch_data, left_on='Work Location', right_on='Work Location')

        # Calculate the sum of salary components
        temp_df = pd.DataFrame({
            "GL Code": rows_matching_value['GL Code'].values[0],
            "Debit": grouped_data[salary_columns_to_sum].sum(axis=1),
            "Credit": 0,
            "Branch": grouped_data['Branch'],
            "Department": grouped_data['Department'] if 'Department' in grouped_data.columns else ' ',
            "Emp Code": grouped_data['Employee ID'] if 'Employee ID' in grouped_data.columns else ' ',
            "Dimension Exists": "Yes" if dimension_value else "No"
        })
    elif value == "EPF Contribution":
        dimension_value = rows_matching_value['Dimension'].iloc[0]
        sign = rows_matching_value['Sign'].iloc[0]

        if dimension_value not in ['B', 'BD', 'BDE']:
            print(f"Unsupported Dimension: {dimension_value}")
            continue

        # Group and sum salary columns based on dimension
        grouped_data = df2.groupby(['Work Location', 'Department', 'Employee ID'][:len(dimension_value)])[
            salary_columns_to_sum2].sum().reset_index()
        grouped_data = grouped_data.merge(branch_data, left_on='Work Location', right_on='Work Location')

        # Calculate the sum of salary components
        temp_df = pd.DataFrame({
            "GL Code": rows_matching_value['GL Code'].values[0],
            "Debit": 0,
            "Credit": grouped_data[salary_columns_to_sum2].sum(axis=1),
            "Branch": grouped_data['Branch'],
            "Department": grouped_data['Department'] if 'Department' in grouped_data.columns else ' ',
            "Emp Code": grouped_data['Employee ID'] if 'Employee ID' in grouped_data.columns else ' ',
            "Dimension Exists": "Yes" if dimension_value else "No"
        })
    

    else:
        if value not in df2.columns:
            continue

        dimension_value = rows_matching_value['Dimension'].iloc[0]
        sign = rows_matching_value['Sign'].iloc[0]

        if dimension_value not in ['B', 'BD', 'BDE']:
            print(f"Unsupported Dimension: {dimension_value}")
            continue

        # Group and sum data based on dimension
        group_by_columns = ['Work Location', 'Department', 'Employee ID'][:len(dimension_value)]
        grouped_data = df2.groupby(group_by_columns)[value].sum().reset_index()

        # Merge with branch data to get the branch information
        grouped_data = grouped_data.merge(branch_data, left_on='Work Location', right_on='Work Location')

        # Create a new DataFrame for the current value
        dimension_exists = "Yes" if dimension_value else "No"
        temp_df = pd.DataFrame({
            "GL Code": rows_matching_value['GL Code'].values[0],
            "Debit": 0,
            "Credit": 0,
            "Branch": grouped_data['Branch'],
            "Department": grouped_data['Department'] if 'Department' in grouped_data.columns else ' ',
            "Emp Code": grouped_data['Employee ID'] if 'Employee ID' in grouped_data.columns else ' ',
            "Dimension Exists": dimension_exists
        })

        # Update debit and credit based on the 'Sign' value
        if sign == 'Credit':
            temp_df["Credit"] = grouped_data[value]
        elif sign == 'Debit':
            temp_df["Debit"] = grouped_data[value]
        else:
            print(f"Unsupported Sign: {sign}")
            continue

    # Check if both "Debit" and "Credit" are zero, and skip if so
    if (temp_df["Debit"] == 0).all() and (temp_df["Credit"] == 0).all():
        continue

    # Append the temporary DataFrame to the result DataFrame
    temp_df["External Doc"] = external_doc_value  # Add External Doc column with user-provided value
    result_df = result_df._append(temp_df, ignore_index=True)

# Save the result to a new Excel file
DIR = os.path.dirname(input_path)
file = "epfOutput" + os.path.basename(input_path)
user_output = os.path.join(DIR, file)

filtered_result_df = result_df[(result_df['Debit'] != 0) | (result_df['Credit'] != 0)]
filtered_result_df.to_excel(user_output, index=False)
