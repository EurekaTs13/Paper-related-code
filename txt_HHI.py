import pandas as pd

# Read Excel files
path1 = r"D:\excel\finance\topic_60_0.8_100_iter=5.xlsx"  # File with topic distributions
path2 = r"D:\excel\Firm_SystemRisk_Finance-20230101_modified.xlsx"  # File with firm and year information
df1 = pd.read_excel(path1)  # Load topic distribution data
df2 = pd.read_excel(path2)  # Load firm and year data

# Extract relevant columns from the second file
filename_column = df2.columns[0]  # Column containing filenames
filename_list = df2[filename_column].to_list()  # List of filenames
year_column = df2.columns[2]  # Column containing years
year_list = df2[year_column].to_list()  # List of years

# Extract filenames from the first file
filename_column_r = df1.columns[0]  # Column containing filenames in the first file
filename_column_r_list = df1[filename_column_r].to_list()  # List of filenames in the first file

# Initialize list to store HHI values
txt_HHI = []

# Iterate through each row of data to calculate HHI
for index1 in range(len(filename_list)):
    for index2 in range(len(filename_column_r_list)):
        # Split the filename to extract parts
        parts = str(filename_column_r_list[index2]).split('-')
        # Check if filename and year match
        if (str(filename_list[index1]) == parts[0] and
            int(parts[1]) + 2000 == year_list[index1]):
            HHI = 0  # Initialize HHI for the current file
            # Calculate HHI by summing squared topic proportions
            for column in range(1, 61):  # Columns 1 to 60 represent topics
                HHI += (df1.iloc[index2, column] ** 2)
            txt_HHI.append(round(HHI, 4))  # Append HHI rounded to 4 decimal places

# Validate output
print(f"Total HHI values calculated: {len(txt_HHI)}")
print("HHI values:", txt_HHI)

# Create a DataFrame to store HHI values
txt_HHI_df = pd.DataFrame(txt_HHI, columns=['HHI'], index=filename_list)

# Save the HHI values to an Excel file
output_excel_path = r"D:\excel\finance\txt_HHI_60.xlsx"
with pd.ExcelWriter(output_excel_path, engine='xlsxwriter') as writer:
    txt_HHI_df.to_excel(writer)
print('HHI calculation and saving completed successfully')