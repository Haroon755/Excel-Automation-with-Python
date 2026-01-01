import pandas as pd
import glob

# Saari Excel files read karo
files = glob.glob("*.xlsx")

all_data = []

for file in files:
    df = pd.read_excel(file)
    all_data.append(df)

# Sab files merge karo
merged_df = pd.concat(all_data, ignore_index=True)

# Missing Amount ko 0 se fill karo
merged_df["Amount"] = merged_df["Amount"].fillna(0)

# Duplicate rows remove
clean_df = merged_df.drop_duplicates()

# Calculations
total_amount = clean_df["Amount"].sum()
average_amount = clean_df["Amount"].mean()
total_records = len(clean_df)

# Summary dataframe
summary = pd.DataFrame({
    "Metric": ["Total Records", "Total Amount", "Average Amount"],
    "Value": [total_records, total_amount, round(average_amount, 2)]
})

# Output Excel with multiple sheets
with pd.ExcelWriter("final_report.xlsx", engine="openpyxl") as writer:
    clean_df.to_excel(writer, sheet_name="Cleaned Data", index=False)
    summary.to_excel(writer, sheet_name="Summary Report", index=False)

print("Automation Complete ✅")
print("Final file created: final_report.xlsx")
