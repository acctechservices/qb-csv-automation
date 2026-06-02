import pandas as pd

print("Loading QuickBooks data...")
df = pd.read_csv('qb_transactions.csv')

df['Date'] = pd.to_datetime(df['Date'])
df['Amount'] = pd.to_numeric(df['Amount'], errors='coerce')

print(f"Loaded {len(df)} rows. Starting audit checks...")

# --- CHECK 1: Duplicate Transactions ---
duplicates = df[df.duplicated(subset=['Date', 'Amount', 'Name/Vendor'], keep=False)]

# --- CHECK 2: Large Round Dollar Amounts ---
round_amounts = df[(df['Amount'] >= 1000) & (df['Amount'] % 100 == 0)]

# --- CHECK 3: Weekend Postings ---
df['Day_of_Week'] = df['Date'].dt.dayofweek
weekend_txns = df[df['Day_of_Week'].isin([5, 6])]

# --- Step 3: Export Findings into an Excel Report ---
print("Writing anomalies to Excel...")
with pd.ExcelWriter('QB_Audit_Report.xlsx') as writer:
    df.to_excel(writer, sheet_name='All Transactions', index=False)
    duplicates.to_excel(writer, sheet_name='Potential Duplicates', index=False)
    round_amounts.to_excel(writer, sheet_name='Large Round Amounts', index=False)
    weekend_txns.to_excel(writer, sheet_name='Weekend Postings', index=False)

print("\nSUCCESS! Your audit is complete.")
print("Check your folder for the new file: 'QB_Audit_Report.xlsx'")
