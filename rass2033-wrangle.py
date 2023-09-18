import pandas as pd

# Define the input and output file paths
input_excel_file = "/Users/femisokoya/Documents/GitHub/drink-drive-data-tables/ras2033.xlsx"
output_csv_file = "rass2033-result.csv"

# Read the Excel file and specify the worksheet
df = pd.read_excel(input_excel_file, sheet_name="Blood_alcohol_levels", header=4)  # Header row is row 5

# Remove the string '[note 3]' from all column headers and strip leading/trailing spaces
df.columns = df.columns.str.replace(r'\[note 3\]', '', regex=True).str.strip()

# Loop through all columns to remove leading/trailing spaces from values
for column in df.columns:
    if df[column].dtype == 'object':  # Only process object (string) columns
        df[column] = df[column].str.strip()

# Rename columns as specified
column_mapping = {
    "Collision year": "Coll_yr",
    "Road user type": "Rd_usr_type",
    "Countries": "Countries",
    "ONS Code": "Ons_code",
    "Under 10 mcg (percentage)": "10_mcg",
    "10 to 49 mcg (percentage)": "49_mcg",
    "50 to 79 mcg (percentage)": "79_mcg",
    "80 to 99 mcg (percentage)": "99_mcg",
    "100 to 149 mcg (percentage)": "149_mcg",
    "150 to 199 mcg (percentage)": "199_mcg",
    "200 or more mcg (percentage)": "200_mcg",
    "Sample size (known blood alcohol level)": "Samp_knwn",
    "Percentage over the limit 10pm to 3:59 am": "10pm_over",
    "Percentage over the limit 4am to 9:59 pm": "4am_over"
}

df = df.rename(columns=column_mapping)

# Map 'Countries' to ONS codes
ons_code_mapping = {
    "England and Wales": "E92000001",
    "Scotland": "S92000003"
}
df['Ons_code'] = df['Countries'].map(ons_code_mapping)

# Reorder columns to have 'Ons_code' after 'Countries'
df = df[['Countries', 'Ons_code'] + [col for col in df.columns if col != 'Countries' and col != 'Ons_code']]

# Save the resulting DataFrame to a CSV file
df.to_csv(output_csv_file, index=False)

print(f"Dataframe saved to {output_csv_file}")
