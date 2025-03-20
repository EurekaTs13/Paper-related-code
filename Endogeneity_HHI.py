import pandas as pd

# Read Excel files
path1 = r"D:\excel\内生性问题修正矩阵.xlsx"
path2 = r"D:\excel\Firm_SystemRisk_Finance-20230101_modified.xlsx"
df1 = pd.read_excel(path1)
df2 = pd.read_excel(path2)

filename_column = df2.columns[0]
filename_list = df2[filename_column].to_list()  # Add parentheses
year_column = df2.columns[2]
year_list = df2[year_column].to_list()  # Change to year_column

filename_column_r = df1.columns[0]
filename_column_r_list = df1[filename_column_r].to_list()  # Add parentheses

txt_HHI = []

# Iterate through each row of data
for index1 in range(len(filename_list)):  # Change to filename_list
    print(index1)
    for index2 in range(len(filename_column_r_list)):  # Change to filename_column_r_list
        parts = str(filename_column_r_list[index2]).split('-')
        if str(filename_list[index1]) == parts[0] and int(parts[1]) + 2000 == year_list[index1]:
            HHI = 0
            for column in range(1, 61):
                HHI += (df1.iloc[index2, column] ** 2)
            txt_HHI.append(round(HHI, 4))

print(len(txt_HHI))
print(txt_HHI)

# Create a DataFrame for HHI values
txt_HHI_df = pd.DataFrame(txt_HHI, columns=['HHI'], index=filename_list)

# Save the results to an Excel file
output_excel_path = r"D:\excel\内生性问题HHI.xlsx"
with pd.ExcelWriter(output_excel_path, engine='xlsxwriter') as writer:
    txt_HHI_df.to_excel(writer)
print('done')