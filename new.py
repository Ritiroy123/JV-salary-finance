import pandas as pd
from pathlib import Path
import os

def process_excel_file(input_path, sheet_name_1, sheet_name_2, salary_columns_to_sum, output_prefix):
    excel_file = pd.ExcelFile(input_path)

    df1 = excel_file.parse(sheet_name_1)
    df2 = excel_file.parse(sheet_name_2)

    branch_data_values = {
        'Work Location': ['Delhi', 'Chennai', 'Pune', 'Mumbai', 'Bengaluru'],
        'Branch': ['01', '01', '05', '07', '03']
    }
    branch_data = pd.DataFrame(branch_data_values)

    result_df = pd.DataFrame(columns=["GL Code", "Debit", "Credit", "External Doc", "Branch", "Department", "Emp Code",
                                      "Dimension Exists"])
    external_doc_value = input("Enter the value for External Doc: ")

    # Initialize grouped_data before the loop
    grouped_data = pd.DataFrame()

    for value in df1['Zoho Heads'].unique():
        rows_matching_value = df1[df1['Zoho Heads'] == value]

        if value in ["Deferred Paid", "Employee State Insurance"]:
            dimension_value = rows_matching_value['Dimension'].iloc[0]
            sign = rows_matching_value['Sign'].iloc[0]

            if dimension_value not in ['B', 'BD', 'BDE']:
                print(f"Unsupported Dimension: {dimension_value}")
                continue

            grouped_data = df2.groupby(['Work Location', 'Department', 'Employee ID'][:len(dimension_value)])[
                salary_columns_to_sum].sum().reset_index()
            grouped_data = grouped_data.merge(branch_data, left_on='Work Location', right_on='Work Location')

            temp_df = pd.DataFrame({
                "GL Code": rows_matching_value['GL Code'].values[0],
                "Debit": 0,
                "Credit": grouped_data[salary_columns_to_sum].sum(axis=1),
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

            group_by_columns = ['Work Location', 'Department', 'Employee ID'][:len(dimension_value)]
            grouped_data = df2.groupby(group_by_columns)[value].sum().reset_index()
            grouped_data = grouped_data.merge(branch_data, left_on='Work Location', right_on='Work Location')

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

            if sign == 'Credit':
                temp_df["Credit"] = grouped_data[value]
            elif sign == 'Debit':
                temp_df["Debit"] = grouped_data[value]
            else:
                print(f"Unsupported Sign: {sign}")
                continue

        if (temp_df["Debit"] == 0).all() and (temp_df["Credit"] == 0).all():
            continue

        temp_df["External Doc"] = external_doc_value
        result_df = result_df._append(temp_df, ignore_index=True)

    dir_path = os.path.dirname(input_path)
    output_file = f"{output_prefix}Output{os.path.basename(input_path)}"
    output_path = os.path.join(dir_path, output_file)

    filtered_result_df = result_df[(result_df['Debit'] != 0) | (result_df['Credit'] != 0)]
    filtered_result_df.to_excel(output_path, index=False)

# Example usage
input_path = Path(input("Enter the path of your Excel file: "))
process_excel_file(input_path, "GL_Code", "Deferred", ["Deferred Salary"], "deferred")

input_path = Path(input("Enter the path of your Excel file: "))
process_excel_file(input_path, "GL_Code", "ESI", ["ESI employer_contribution"], "esi")
