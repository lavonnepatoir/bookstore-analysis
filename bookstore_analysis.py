import pandas as pd

# Loads the Excel file
df = pd.read_excel('bookstore_sales.xlsx')

# Displays the first few rows
print(df.head())

# Checks for missing values
print(df.isnull().sum())

df['Year Published'] = pd.to_datetime(df['Year Published'])

df_filtered = df[(df['Year Published'] >= '1997')]

total_sales = df['Sales'].sum()
average_sales = df['Sales'].mean()

print(f'Total Sales: ${total_sales}')
print(f'Average Sales per Book: ${average_sales}')

top_books = df.groupby('Book Title')['Sales'].sum().sort_values(ascending=False).head(10)
print(top_books)

import matplotlib.pyplot as plt
import seaborn as sns

plt.figure(figsize=(10, 6))
top_books.plot(kind='bar', title='Top-Selling Books')
plt.xlabel('Book Title')
plt.ylabel('Total Sales')
plt.xticks(rotation=45, ha='right')
plt.show()

with pd.ExcelWriter('bookstore_sales_report.xlsx') as writer:
    df.to_excel(writer, sheet_name='Raw Data')
    top_books.to_frame().to_excel(writer, sheet_name='Top Books')