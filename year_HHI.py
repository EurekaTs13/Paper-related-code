import pandas as pd

# Initialize dictionaries
year_company_dict = {str(year).zfill(2): 0 for year in range(0, 21)}
year_sum_weight_dict = {str(year).zfill(2): [0 for _ in range(60)] for year in range(0, 21)}
year_aver_weight_dict = {str(year).zfill(2): [0 for _ in range(60)] for year in range(0, 21)}
year_HHI_dict = {str(year).zfill(2): 0 for year in range(0, 21)}

# Read the Excel file
path = r"D:\excel\finance\topic_60_0.8_100_iter=5.xlsx"
df = pd.read_excel(path)

# Iterate through each row of data
for row in range(0, 11400):
    info = str(df.iloc[row, 0]).split('-')
    year = info[1]
    year_company_dict[year] += 1
    for column in range(1, 81):  # Update year_sum_weight_dict
        year_sum_weight_dict[year][column - 1] += float(df.iloc[row, column])

# Calculate the average weight
for year in year_company_dict:
    if year_company_dict[year] > 0:
        year_aver_weight_dict[year] = [weight / year_company_dict[year] for weight in year_sum_weight_dict[year]]

# Calculate the annual HHI
for year in year_company_dict:
    for aver_weight in year_aver_weight_dict[year]:
        year_HHI_dict[year] += round(float(aver_weight) ** 2, 4)

print(year_company_dict)
print(year_sum_weight_dict)
print(year_aver_weight_dict)
print(year_HHI_dict)