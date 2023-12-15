# # import pandas as pd


# # # df_rows = pd.read_excel('employee_salary_register_re (2).xlsx')


# # # df_columns = pd.read_excel('Salary GLs.xlsx')


# # #merged_data = pd.concat([file1, file2], ignore_index=True)
# # #merged_data.to_excel('modified_merged_file.xlsx', index=False)
# # import pandas as pd

# # # Load the first Excel file
# # file1 = pd.read_excel('employee_salary_register_re (2).xlsx')

# # # Load the second Excel file
# # file2 = pd.read_excel('Salary GLs.xlsx')
# # file3 = pd.read_excel('Sheet3.xlsx')
# # if(file2.sign ==credit)

# # #this is like group by branch department and employee id
# # if(file2.dimension==B then file1=work location which present in file3 work location==branch ||dimesion==D then file1==department||dimension==E then file1==employee id)
# # if(file2=='Regular salary'):
# #     add(file1.basic +hous rent allowance+Conveyance Allowance+PFP Program+Communication Allowance+Standard Special Allowance)
    
# # if(file2.row == file1.column)
# # add(file1.column)   
# # new combined file.to_excel(gl code,debit valuel,credit value,branch,department,emp-id) 



# # new_combined_file.to_excel('new_combined_excel_file.xlsx', index=False)



# # import pandas as pd

# # # Load the first Excel file
# # df1 = pd.read_excel('Salary GLs.xlsx')

# # # Load the second Excel file
# # df2 = pd.read_excel('employee_salary_register_re (2).xlsx')

# # # Load the third Excel file
# # df3 = pd.read_excel('Sheet3.xlsx')

# # # Combine the data from the three dataframes
# # combined_df = pd.concat([df1, df2, df3], ignore_index=True)

# # # Check if "Credit" and "Debit" columns exist, and assume a value of 0 if they don't
# # if "Credit" not in combined_df:
# #     combined_df["Credit"] = 0
# # if "Debit" not in combined_df:
# #     combined_df["Debit"] = 0

# # # Sum the "Debit" and "Credit" values based on grouping columns
# # result_df = combined_df.groupby(["GL Code", "Branch", "Department"]).agg({"Debit": "sum", "Credit": "sum"}).reset_index()

# # # Rename the columns to match the desired output
# # result_df.rename(columns={"GL Code": "GL Code", "Debit": "Debit", "Credit": "Credit", "Branch": "Branch", "Department": "Department"}, inplace=True)
# # result_df["Dimension Exists"] = "Yes"

# # # Save the result to an output Excel file
# # output_filename = "final_output.xlsx"
# # result_df.to_excel(output_filename, index=False)

# # print(f"Final data has been exported to {output_filename}")


# import pandas as pd

# # Read the first Excel file with "Zoho Heads"


# # comparison_values = df1['Zoho Heads'].values

# # for i in comparison_values:
# #     for col_name in df2.columns:
# #         if i==col_name:
# #          print()
# # Create a new DataFrame to store the sums


# #print(comparison_values)




# # if 'Regular Salary' in comparison_values:
# #         columns_to_sum = ['Basic', 'House Rent Allowance', 'Conveyance Allowance', 'PFP Program', 'Communication Allowance', 'Standard Special Allowance']
# #         print(columns_to_sum)
# #         regular_salary_sum = df2[columns_to_sum][:-3].sum().sum()
# #         print(f"Sum of specified columns in df2 for 'Regular Salary': {regular_salary_sum}")
        
# # # for value in comparison_values:
    
# #     if value in df2.columns:
# #         # Exclude the last row when calculating the sum
# #         column_values = df2[value][:-3]
# #         # print(column_values)
# #         column_sum = column_values.sum()
# #         print(f"Sum of '{value}' : {column_sum}")



# # Assuming you have the necessary data in df1 and df2

# # Create an empty DataFrame to store the grouped data



# # Create an empty DataFrame to store the grouped data
# # import pandas as pd

# # # Create an empty DataFrame to store the grouped data
# # grouped_data_df = pd.DataFrame(columns=['GL Code', 'Debit', 'Credit', 'Branch', 'Department', 'Emp Code', 'Dimension Exists'])

# # # Read the branch data from the third Excel file

# # for index, row in df1.iterrows():
# #     gl_code = row['GL Code']  # Assuming GL Code is the correct column name
# #     dimension_value = row['Dimension']

# #     if dimension_value == 'B':
# #         grouped_data = df2[:-3].groupby(['Work Location'])['Zoho Heads'].sum().reset_index()
# #     elif dimension_value == 'BD':
# #         grouped_data = df2[:-3].groupby(['Work Location', 'Department'])['Zoho Heads'].sum().reset_index()
# #         # Extract the Department from df2
# #         department = grouped_data['Department'].unique()
# #         if len(department) == 1:
# #             department = department[0]
# #         else:
# #             print(f"Multiple departments found for {gl_code}, skipping.")
# #             continue
# #     elif dimension_value == 'BDE':
# #         grouped_data = df2[:-3].groupby(['Work Location', 'Department', 'Employee Id'])['Zoho Heads'].sum().reset_index()
# #         # Extract the Department and Employee Id from df2
# #         department = grouped_data['Department'].unique()
# #         if len(department) == 1:
# #             department = department[0]
# #         else:
# #             print(f"Multiple departments found for {gl_code}, skipping.")
# #             continue
# #         emp_code = grouped_data['Employee Id'].unique()
# #         if len(emp_code) == 1:
# #             emp_code = emp_code[0]
# #         else:
# #             print(f"Multiple employee IDs found for {gl_code}, skipping.")
# #             continue
# #     else:
# #         print(f"Unsupported Dimension: {dimension_value}")
# #         continue

# #     # Determine Debit or Credit based on the sign of the value
# #     grouped_data['Debit'] = grouped_data['Zoho Heads'].apply(lambda x: x if x < 0 else 0)
# #     grouped_data['Credit'] = grouped_data['Zoho Heads'].apply(lambda x: -x if x > 0 else 0)

# #     # Merge with branch data
# #     grouped_data = pd.merge(grouped_data, branch_data, on='Work Location', how='left')

# #     grouped_data_df = pd.concat([grouped_data_df, grouped_data], ignore_index=True)

# #     # Fill other columns
# #     grouped_data_df['GL Code'] = gl_code
# #     grouped_data_df['Department'] = department
# #     grouped_data_df['Emp Code'] = emp_code
# #     grouped_data_df['Dimension Exists'] = 'Yes'

# # # Save the data to an Excel file
# # grouped_data_df = grouped_data_df[['GL Code', 'Debit', 'Credit', 'Branch', 'Department', 'Emp Code', 'Dimension Exists']]
# # grouped_data_df.to_excel("grouped_data.xlsx", index=False)


# # df1 = pd.read_excel("Salary GLs.xlsx")

# # # Read the second Excel file with employee data
# # df2 = pd.read_excel("employee_salary_register_re (2).xlsx")
# # branch_data = pd.read_excel("Sheet3.xlsx")

# # comparison_values = df1['Zoho Heads'].unique()
# # gl_code = df1['GL Code'].unique()


# # for value in comparison_values:
# #     if value in df2.columns:
# #         # Check the 'Dimension' value in df1
# #         dimension_value = df1[df1['Zoho Heads'] == value]['Dimension'].iloc[0]
        
# #         if dimension_value == 'B':
# #             # Group by 'Work Location'
# #             grouped_data = df2[:-3].groupby(['Work Location'])[value].sum()
# #         elif dimension_value == 'BD':
# #             # Group by 'Work Location' and 'Department'
# #             grouped_data = df2[:-3].groupby(['Work Location', 'Department'])[value].sum()
# #         elif dimension_value == 'BDE':
# #             # Group by 'Work Location', 'Department', and 'Employee Id'
# #             grouped_data = df2[:-3].groupby(['Work Location', 'Department', 'Employee ID'])[value].sum()
# #         else:
# #             print(f"Unsupported Dimension: {dimension_value}")
# #             continue
        
# # #         print(f"Grouped data for '{value}' ({dimension_value} Dimension):\n{grouped_data}")

# # import pandas as pd

# # # Read the first Excel file with GL data
# # df1 = pd.read_excel("Salary GLs.xlsx")

# # # Read the second Excel file with employee data
# # df2 = pd.read_excel("employee_salary_register_re (2).xlsx")

# # # Read the third Excel file with branch data
# # branch_data = pd.read_excel("Sheet3.xlsx")

# # # Initialize an empty DataFrame to store the final result
# # result_df = pd.DataFrame(columns=["GL Code", "Debit", "Credit", "Branch", "Department", "Emp Code", "Dimension Exists"])

# # comparison_values = df1['Zoho Heads'].unique()

# # # Iterate through the unique values in 'Zoho Heads' column of df1
# # for value in comparison_values:
# #     if value in df2.columns:
# #         # Check the 'Dimension' value in df1
# #         dimension_value = df1[df1['Zoho Heads'] == value]['Dimension'].iloc[0]

# #         # Determine whether the sum should be in credit or debit based on the 'Sign' column
# #         sign = df1[df1['Zoho Heads'] == value]['Sign'].iloc[0]

# #         if dimension_value == 'B':
# #             # Group by 'Work Location'
# #             grouped_data = df2.groupby(['Work Location'])[value].sum().reset_index()
# #         elif dimension_value == 'BD':
# #             # Group by 'Work Location' and 'Department'
# #             grouped_data = df2.groupby(['Work Location', 'Department'])[value].sum().reset_index()
# #         elif dimension_value == 'BDE':
# #             # Group by 'Work Location', 'Department', and 'Employee Id'
# #             grouped_data = df2.groupby(['Work Location', 'Department', 'Employee ID'])[value].sum().reset_index()
# #         else:
# #             print(f"Unsupported Dimension: {dimension_value}")
# #             continue

# #         # Merge with branch data to get the branch information
# #         grouped_data = grouped_data.merge(branch_data, left_on='Work Location', right_on='Work Location')

# #         # Create a new DataFrame for the current value
# #         dimension_exists = "Yes" if dimension_value else "No"
# #         temp_df = pd.DataFrame({
# #             "GL Code": df1[df1['Zoho Heads'] == value]['GL Code'].values[0],
# #             "Debit": 0,
# #             "Credit": 0,
# #             "Branch": grouped_data['Branch'],
# #             "Department": grouped_data['Department'] if 'Department' in grouped_data.columns else ' ',
# #             "Emp Code": grouped_data['Employee ID'] if 'Employee ID' in grouped_data.columns else ' ',
# #             "Dimension Exists": dimension_exists
# #         })

# #         # Update debit and credit based on the 'Sign' value
# #         if sign == 'Credit':
# #             temp_df["Credit"] = grouped_data[value]
# #         elif sign == 'Debit':
# #             temp_df["Debit"] = grouped_data[value]
# #         else:
# #             print(f"Unsupported Sign: {sign}")
# #             continue


# #         if not (temp_df["Debit"] == 0).all() or not (temp_df["Credit"] == 0).all():
# #             result_df = result_df._append(temp_df, ignore_index=True)

# # # Save the result to a new Excel file
# # # result_df.to_excel("output4.xlsx", index=False)
# # import pandas as pd

# # # Read the first Excel file with GL data
# # df1 = pd.read_excel("Salary GLs.xlsx")

# # # Read the second Excel file with employee data
# # df2 = pd.read_excel("employee_salary_register_re (2).xlsx")

# # # Read the third Excel file with branch data
# # branch_data = pd.read_excel("Sheet3.xlsx")
# # salary_columns_to_sum = ['Basic', 'House Rent Allowance', 'Conveyance Allowance', 'PFP Program', 'Communication Allowance', 'Standard Special Allowance']

# # # Initialize an empty DataFrame to store the final result
# # result_df = pd.DataFrame(columns=["GL Code", "Debit", "Credit", "Branch", "Department", "Emp Code", "Dimension Exists"])

# # comparison_values = df1['Zoho Heads'].unique()

# # # Iterate through the unique values in 'Zoho Heads' column of df1
# # for value in comparison_values:
# #     if value in df2.columns:
# #         # Check the 'Dimension' value in df1
# #         dimension_value = df1[df1['Zoho Heads'] == value]['Dimension'].iloc[0]

# #         # Determine whether the sum should be in credit or debit based on the 'Sign' column
# #         sign = df1[df1['Zoho Heads'] == value]['Sign'].iloc[0]

# #         if dimension_value == 'B':
# #             # Group by 'Work Location'
# #             grouped_data = df2.groupby(['Work Location'])[value].sum().reset_index()
# #         elif dimension_value == 'BD':
# #             # Group by 'Work Location' and 'Department'
# #             grouped_data = df2.groupby(['Work Location', 'Department'])[value].sum().reset_index()
# #         elif dimension_value == 'BDE':
# #             # Group by 'Work Location', 'Department', and 'Employee Id'
# #             grouped_data = df2.groupby(['Work Location', 'Department', 'Employee ID'])[value].sum().reset_index()
# #         else:
# #             print(f"Unsupported Dimension: {dimension_value}")
# #             continue

# #         # Merge with branch data to get the branch information
# #         grouped_data = grouped_data.merge(branch_data, left_on='Work Location', right_on='Work Location')

# #         # Create a new DataFrame for the current value
# #         dimension_exists = "Yes" if dimension_value else "No"
# #         temp_df = pd.DataFrame({
# #             "GL Code": df1[df1['Zoho Heads'] == value]['GL Code'].values[0],
# #             "Debit": 0,
# #             "Credit": 0,
# #             "Branch": grouped_data['Branch'],
# #             "Department": grouped_data['Department'] if 'Department' in grouped_data.columns else ' ',
# #             "Emp Code": grouped_data['Employee ID'] if 'Employee ID' in grouped_data.columns else ' ',
# #             "Dimension Exists": dimension_exists
# #         })

# #         # Update debit and credit based on the 'Sign' value
# #         if sign == 'Credit':
# #             temp_df["Credit"] = grouped_data[value]
# #         elif sign == 'Debit':
# #             temp_df["Debit"] = grouped_data[value]
# #         else:
# #             print(f"Unsupported Sign: {sign}")
# #             continue

# #         # Check if both "Debit" and "Credit" are zero, and skip if so
# #         if (temp_df["Debit"] == 0).all() and (temp_df["Credit"] == 0).all():
# #             continue

# #         # Append the temporary DataFrame to the result DataFrame
# #         result_df = result_df._append(temp_df, ignore_index=True)

# #     elif value=="Regular Salary":
# #         # Check the 'Dimension' value in df1
# #         dimension_value = df1[df1['Zoho Heads'] == value]['Dimension'].iloc[0]

# #         # Determine whether the sum should be in credit or debit based on the 'Sign' column
# #         sign = df1[df1['Zoho Heads'] == value]['Sign'].iloc[0]

# #         if dimension_value == 'B':
# #             # Group by 'Work Location'
# #             grouped_data = df2.groupby(['Work Location'])[salary_columns_to_sum].sum().reset_index()
# #         elif dimension_value == 'BD':
# #             # Group by 'Work Location' and 'Department'
# #             grouped_data = df2.groupby(['Work Location', 'Department'])[salary_columns_to_sum].sum().reset_index()
# #         elif dimension_value == 'BDE':
# #             grouped_data = df2.groupby(['Work Location', 'Department', 'Employee ID'])[salary_columns_to_sum].sum().reset_index()
# #         else:
# #             print(f"Unsupported Dimension: {dimension_value}")
# #             continue

# #         # Merge with branch data to get the branch information
# #         grouped_data = grouped_data.merge(branch_data, left_on='Work Location', right_on='Work Location')

# #         # Create a new DataFrame for the current value
# #         dimension_exists = "Yes" if dimension_value else "No"
# #         temp_df = pd.DataFrame({
# #             "GL Code": df1[df1['Zoho Heads'] == value]['GL Code'].values[0],
# #             "Debit": 0,
# #             "Credit": 0,
# #             "Branch": grouped_data['Branch'],
# #             "Department": grouped_data['Department'] if 'Department' in grouped_data.columns else ' ',
# #             "Emp Code": grouped_data['Employee ID'] if 'Employee ID' in grouped_data.columns else ' ',
# #             "Dimension Exists": dimension_exists
# #         })

# #         # Update debit and credit based on the 'Sign' value
# #         if sign == 'Credit':
# #             temp_df["Credit"] = grouped_data[salary_columns_to_sum]
# #         elif sign == 'Debit':
# #             temp_df["Debit"] = grouped_data[salary_columns_to_sum]

# #         else:
# #             print(f"Unsupported Sign: {sign}")
# #             continue


# #         # Check if both "Debit" and "Credit" are zero, and skip if so
# #         if (temp_df["Debit"] == 0).all() and (temp_df["Credit"] == 0).all():
# #             continue

# #         # Append the temporary DataFrame to the result DataFrame
# #         result_df = result_df._append(temp_df, ignore_index=True)



# # # Save the result to a new Excel file



# # # Save the result to a new Excel file
# # filtered_result_df = result_df[(result_df['Debit'] != 0) | (result_df['Credit'] != 0)]
# # filtered_result_df.to_excel("output2.xlsx", index=False)

# import pandas as pd
# from pathlib import Path
# import os

# # Read the Excel files
# input_1 = Path(input("Enter the path of your first file: "))
# input_2 = Path(input("Enter the path of your second file: "))
# input_3 = Path(input("Enter the path of your third file: "))



# df1 = pd.read_excel(input_1)
# df2 = pd.read_excel(input_2)
# branch_data = pd.read_excel(input_3)

# # Columns to sum for salary calculation
# salary_columns_to_sum = [
#     'Basic', 'House Rent Allowance', 'Conveyance Allowance',
#     'PFP Program', 'Communication Allowance', 'Standard Special Allowance'
# ]

# # Initialize an empty DataFrame to store the final result
# result_df = pd.DataFrame(columns=["GL Code", "Debit", "Credit", "Branch", "Department", "Emp Code", "Dimension Exists"])

# # Iterate through the unique values in 'Zoho Heads' column of df1
# for value in df1['Zoho Heads'].unique():
#     # Filter rows in df1 based on the current value
#     rows_matching_value = df1[df1['Zoho Heads'] == value]

#     if value == "Regular Salary":
#         dimension_value = rows_matching_value['Dimension'].iloc[0]
#         sign = rows_matching_value['Sign'].iloc[0]

#         if dimension_value not in ['B', 'BD', 'BDE']:
#             print(f"Unsupported Dimension: {dimension_value}")
#             continue

#         # Group and sum salary columns based on dimension
#         grouped_data = df2.groupby(['Work Location', 'Department', 'Employee ID'][:len(dimension_value)])[
#             salary_columns_to_sum].sum().reset_index()
#         grouped_data = grouped_data.merge(branch_data, left_on='Work Location', right_on='Work Location')

#         # Calculate the sum of salary components
#         temp_df = pd.DataFrame({
#             "GL Code": rows_matching_value['GL Code'].values[0],
#             "Debit": grouped_data[salary_columns_to_sum].sum(axis=1),
#             "Credit": 0,
#             "Branch": grouped_data['Branch'],
#             "Department": grouped_data['Department'] if 'Department' in grouped_data.columns else ' ',
#             "Emp Code": grouped_data['Employee ID'] if 'Employee ID' in grouped_data.columns else ' ',
#             "Dimension Exists": "Yes" if dimension_value else "No"
#         })

#     else:
#         if value not in df2.columns:
#             continue

#         dimension_value = rows_matching_value['Dimension'].iloc[0]
#         sign = rows_matching_value['Sign'].iloc[0]

#         if dimension_value not in ['B', 'BD', 'BDE']:
#             print(f"Unsupported Dimension: {dimension_value}")
#             continue

#         # Group and sum data based on dimension
#         group_by_columns = ['Work Location', 'Department', 'Employee ID'][:len(dimension_value)]
#         grouped_data = df2.groupby(group_by_columns)[value].sum().reset_index()

#         # Merge with branch data to get the branch information
#         grouped_data = grouped_data.merge(branch_data, left_on='Work Location', right_on='Work Location')

#         # Create a new DataFrame for the current value
#         dimension_exists = "Yes" if dimension_value else "No"
#         temp_df = pd.DataFrame({
#             "GL Code": rows_matching_value['GL Code'].values[0],
#             "Debit": 0,
#             "Credit": 0,
#             "Branch": grouped_data['Branch'],
#             "Department": grouped_data['Department'] if 'Department' in grouped_data.columns else ' ',
#             "Emp Code": grouped_data['Employee ID'] if 'Employee ID' in grouped_data.columns else ' ',
#             "Dimension Exists": dimension_exists
#         })

#         # Update debit and credit based on the 'Sign' value
#         if sign == 'Credit':
#             temp_df["Credit"] = grouped_data[value]
#         elif sign == 'Debit':
#             temp_df["Debit"] = grouped_data[value]
#         else:
#             print(f"Unsupported Sign: {sign}")
#             continue

#     # Check if both "Debit" and "Credit" are zero, and skip if so
#     if (temp_df["Debit"] == 0).all() and (temp_df["Credit"] == 0).all():
#         continue

#     # Append the temporary DataFrame to the result DataFrame
#     result_df = result_df._append(temp_df, ignore_index=True)

# # Save the result to a new Excel file
# DIR = os.path.dirname(input_1)
# file = "output"+ os.path.basename(input_1) 

# user_output = os.path.join(DIR, file)

# filtered_result_df = result_df[(result_df['Debit'] != 0) | (result_df['Credit'] != 0)]
# filtered_result_df.to_excel(user_output, index=False)


import pandas as pd
from pathlib import Path
import os

# Read the Excel files
input_path = Path(input("Enter the path of your Excel file: "))
excel_file = pd.ExcelFile(input_path)

# Specify the sheet names
sheet_name_1 = ("GL_Code")
sheet_name_2 = ("Salary")

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
    'Basic', 'House Rent Allowance', 'Conveyance Allowance',
    'PFP Program', 'Communication Allowance', 'Standard Special Allowance'
]

# Initialize an empty DataFrame to store the final result
result_df = pd.DataFrame(columns=["GL Code", "Debit", "Credit","External Doc", "Branch", "Department", "Emp Code", "Dimension Exists"])
external_doc_value = input("Enter the value for External Doc: ")

# Iterate through the unique values in 'Zoho Heads' column of df1
for value in df1['Zoho Heads'].unique():
    # Filter rows in df1 based on the current value
    rows_matching_value = df1[df1['Zoho Heads'] == value]

    if value == "Regular Salary":
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
    temp_df["External Doc"] = external_doc_value  
    result_df = result_df._append(temp_df, ignore_index=True)

# Save the result to a new Excel file
DIR = os.path.dirname(input_path)
file = "output2" + os.path.basename(input_path)
user_output = os.path.join(DIR, file)

filtered_result_df = result_df[(result_df['Debit'] != 0) | (result_df['Credit'] != 0)]
filtered_result_df.to_excel(user_output, index=False)


