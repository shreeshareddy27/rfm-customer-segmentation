import pandas as pd
import numpy as np
from datetime import datetime

# Read the data
print("Reading data...")
df = pd.read_excel("Online Retail.xlsx")

# Remove rows with missing CustomerID
df = df[df['CustomerID'].notna()]

# Calculate Revenue
df['Revenue'] = df['Quantity'] * df['UnitPrice']

# Remove negative quantities (returns)
df = df[df['Quantity'] > 0]

# Convert InvoiceDate to datetime
df['InvoiceDate'] = pd.to_datetime(df['InvoiceDate'])

# Set analysis date (day after last transaction)
analysis_date = df['InvoiceDate'].max() + pd.Timedelta(days=1)
print(f"Analysis Date: {analysis_date}")

# Calculate RFM metrics
print("Calculating RFM...")
rfm = df.groupby('CustomerID').agg({
    'InvoiceDate': lambda x: (analysis_date - x.max()).days,  # Recency
    'InvoiceNo': 'nunique',  # Frequency
    'Revenue': 'sum'  # Monetary
})

rfm.columns = ['Recency', 'Frequency', 'Monetary']

# Create RFM scores (1-5, where 5 is best)
rfm['R_Score'] = pd.qcut(rfm['Recency'], 5, labels=[5,4,3,2,1])
rfm['F_Score'] = pd.qcut(rfm['Frequency'].rank(method='first'), 5, labels=[1,2,3,4,5])
rfm['M_Score'] = pd.qcut(rfm['Monetary'], 5, labels=[1,2,3,4,5])

# Calculate RFM Score
rfm['RFM_Score'] = rfm['R_Score'].astype(str) + rfm['F_Score'].astype(str) + rfm['M_Score'].astype(str)

# Create segments
def segment_customer(row):
    if row['R_Score'] >= 4 and row['F_Score'] >= 4:
        return 'Champions'
    elif row['R_Score'] >= 3 and row['F_Score'] >= 3:
        return 'Loyal Customers'
    elif row['R_Score'] >= 4:
        return 'Potential Loyalists'
    elif row['F_Score'] >= 4:
        return 'Cant Lose Them'
    elif row['R_Score'] <= 2 and row['F_Score'] <= 2:
        return 'Lost'
    elif row['R_Score'] <= 2:
        return 'At Risk'
    else:
        return 'Need Attention'

rfm['Segment'] = rfm.apply(segment_customer, axis=1)

# Reset index to make CustomerID a column
rfm = rfm.reset_index()

# Save to new Excel file
print("Saving results...")
rfm.to_excel("RFM_Analysis_Results.xlsx", index=False)

print(f"\nDone! Created file: RFM_Analysis_Results.xlsx")
print(f"Total customers analyzed: {len(rfm)}")
print(f"\nSegment distribution:")
print(rfm['Segment'].value_counts())

