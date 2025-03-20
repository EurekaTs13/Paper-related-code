import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.lines import Line2D

# Set global font to Times New Roman
plt.rcParams['font.family'] = 'Times New Roman'
# Ensure negative signs are displayed correctly
plt.rcParams['axes.unicode_minus'] = False

topic_excel_path = r"D:\excel\finance\topic_60_0.8_100_iter=5.xlsx"
finance_risk_excel_path = r"D:\excel\Firm_SystemRisk_modified_delete_year2000.xlsx"

def find_top3_values_in_row(df, row_index):
    # Get the data of the specified row
    row_data = df.iloc[row_index, 1:]

    # Sort and get the top 3 largest values and their corresponding column names
    sorted_row_data = row_data.sort_values(ascending=False).head(3)

    # Combine column names and their values into small arrays, with column names first and values second
    top3 = [[col, sorted_row_data[col]] for col in sorted_row_data.index]

    return top3

df1 = pd.read_excel(finance_risk_excel_path)
df2 = pd.read_excel(topic_excel_path)
cik_column = df1.columns[0]
year_column = df1.columns[2]
cik_list = df1[cik_column].to_list()
year_list = df1[year_column].to_list()
sic_column = df1.columns[5]
sic_list = df1[sic_column].to_list()
filename_column = df2.columns[0]
filename_list = df2[filename_column].to_list()

# Predefine lists for each industry
banking = []
insurance = []
asset_management = []

# Order of 60 topics (from product-oriented to service-oriented)
topic_order = [
    "Topic 48", "Topic 50", "Topic 21", "Topic 41", "Topic 47", "Topic 13", "Topic 0", "Topic 56", "Topic 43",
    "Topic 51", "Topic 45", "Topic 20", "Topic 25", "Topic 55", "Topic 57", "Topic 24", "Topic 28", "Topic 26",
    "Topic 49", "Topic 6", "Topic 38", "Topic 23", "Topic 10", "Topic 39", "Topic 22", "Topic 14", "Topic 42",
    "Topic 35", "Topic 12", "Topic 29", "Topic 11", "Topic 9", "Topic 16", "Topic 53", "Topic 37", "Topic 33",
    "Topic 58", "Topic 40", "Topic 3", "Topic 32", "Topic 36", "Topic 4", "Topic 46", "Topic 15", "Topic 8",
    "Topic 34", "Topic 52", "Topic 17", "Topic 7", "Topic 30", "Topic 5", "Topic 27", "Topic 31", "Topic 19",
    "Topic 54", "Topic 44", "Topic 2", "Topic 1", "Topic 18", "Topic 59"
]

print(len(topic_order))

# Determine the X values for each topic
topic_x_values = list(range(len(topic_order)))

# Plot the scatter plot
fig, ax = plt.subplots(figsize=(14, 8))

handles = [
    Line2D([0], [0], marker='o', color='w', markerfacecolor=(236/255, 187/255, 107/255), markersize=10, label='Banking'),
    Line2D([0], [0], marker='o', color='w', markerfacecolor=(107/255, 157/255, 170/255), markersize=10, label='Insurance'),
    Line2D([0], [0], marker='o', color='w', markerfacecolor=(144/255, 99/255, 159/255), markersize=10, label='Asset Management and Financial Services')
]

i = 0

for index1 in range(len(cik_list)):
    for index2 in range(len(filename_list)):
        filename = str(filename_list[index2])
        parts = filename.split('-')
        if parts[0] == str(cik_list[index1]) and int(parts[1]) + 2000 == int(year_list[index1]):
            sic = sic_list[index1]
            result = find_top3_values_in_row(df2, index2)

            if sic // 100 == 60 or sic // 100 == 61:  # Banking
                for col, val in result:
                    x = topic_order.index(str(col))  # Find the X value corresponding to the topic
                    ax.scatter(x, val, color=(236 / 255, 187 / 255, 107 / 255), alpha=0.7)
                banking.append(result)

            elif sic // 100 == 63 or sic // 100 == 64:  # Insurance
                for col, val in result:
                    x = topic_order.index(str(col))  # Find the X value corresponding to the topic
                    ax.scatter(x, val, color=(107 / 255, 157 / 255, 170 / 255), alpha=0.7)
                insurance.append(result)

            else:  # Asset Management
                for col, val in result:
                    x = topic_order.index(str(col))  # Find the X value corresponding to the topic
                    ax.scatter(x, val, color=(144 / 255, 99 / 255, 159 / 255), alpha=0.7)
                asset_management.append(result)

            i += 1
            print(i)

# Add labels, title, and legend
ax.set_xlabel('Topic', fontsize=16)
ax.set_ylabel('Value', fontsize=16)
ax.set_xticks(topic_x_values)
ax.set_xticklabels([f"Topic {int(str(i).split(' ')[1]) + 1}" for i in topic_order], fontsize=14, rotation=90)
ax.tick_params(axis='y', labelsize=14)
ax.legend(handles=handles, fontsize=12, title='', loc='upper left', frameon=False)

# Remove the title
ax.set_title('')

plt.tight_layout()
plt.show()

# Output the number of points in each category
print(f'Banking: {len(banking)}')
print(f'Insurance: {len(insurance)}')
print(f'Asset Management: {len(asset_management)}')