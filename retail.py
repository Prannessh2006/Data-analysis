# Import necessary libraries
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from datetime import datetime

# Task 1: Load the dataset into a Pandas DataFrame
# Read the Excel file
df = pd.read_excel('Online Retail.xlsx')

# Display the first few rows to get an overview
print("First 5 rows of the dataset:")
print(df.head())

# Task 2: Data Cleaning
# Check for missing values
print("\nMissing values in each column:")
print(df.isnull().sum())

# Handle missing values
# Drop rows where CustomerID is missing since it's critical for customer analysis
df = df.dropna(subset=['CustomerID'])

# Remove rows with negative or zero quantities
df = df[df['Quantity'] > 0]

# Remove rows with negative or zero UnitPrice
df = df[df['UnitPrice'] > 0]

# Convert InvoiceDate to datetime if not already
df['InvoiceDate'] = pd.to_datetime(df['InvoiceDate'], errors='coerce')

# Remove any rows with invalid dates
df = df.dropna(subset=['InvoiceDate'])

# Check for duplicates and remove them
print("\nNumber of duplicate rows:", df.duplicated().sum())
df = df.drop_duplicates()

# Task 3: Basic Statistics
# Display basic statistics
print("\nBasic statistics of numerical columns:")
print(df.describe())

# Task 4: Data Visualization
# Create a figure for visualizations
plt.figure(figsize=(15, 10))

# Histogram of Quantity
plt.subplot(2, 2, 1)
sns.histplot(df['Quantity'], bins=50, kde=True)
plt.title('Distribution of Quantity')
plt.xlabel('Quantity')
plt.xlim(0, df['Quantity'].quantile(0.99))  # Limit to 99th percentile to handle outliers

# Histogram of UnitPrice
plt.subplot(2, 2, 2)
sns.histplot(df['UnitPrice'], bins=50, kde=True)
plt.title('Distribution of UnitPrice')
plt.xlabel('UnitPrice')
plt.xlim(0, df['UnitPrice'].quantile(0.99))  # Limit to 99th percentile

# Calculate total sales
df['TotalSales'] = df['Quantity'] * df['UnitPrice']

# Bar plot of top 10 countries by total sales
top_countries = df.groupby('Country')['TotalSales'].sum().sort_values(ascending=False).head(10)
plt.subplot(2, 2, 3)
top_countries.plot(kind='bar')
plt.title('Top 10 Countries by Total Sales')
plt.xlabel('Country')
plt.ylabel('Total Sales')
plt.xticks(rotation=45)

# Scatter plot of Quantity vs UnitPrice
plt.subplot(2, 2, 4)
sns.scatterplot(x='UnitPrice', y='Quantity', data=df)
plt.title('Quantity vs UnitPrice')
plt.xlabel('UnitPrice')
plt.ylabel('Quantity')
plt.xlim(0, df['UnitPrice'].quantile(0.99))
plt.ylim(0, df['Quantity'].quantile(0.99))

plt.tight_layout()
plt.show()

# Task 5: Analyze Sales Trends Over Time
# Extract month and day of the week
df['Month'] = df['InvoiceDate'].dt.month
df['DayOfWeek'] = df['InvoiceDate'].dt.day_name()

# Monthly sales trend
monthly_sales = df.groupby('Month')['TotalSales'].sum()
plt.figure(figsize=(10, 5))
monthly_sales.plot(kind='bar')
plt.title('Monthly Sales Trend')
plt.xlabel('Month')
plt.ylabel('Total Sales')
plt.xticks(rotation=0)
plt.show()

# Sales by day of the week
day_sales = df.groupby('DayOfWeek')['TotalSales'].sum().reindex(
    ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'])
plt.figure(figsize=(10, 5))
day_sales.plot(kind='bar')
plt.title('Sales by Day of the Week')
plt.xlabel('Day of the Week')
plt.ylabel('Total Sales')
plt.xticks(rotation=45)
plt.show()

# Task 6: Top-Selling Products and Countries
# Top 10 products by quantity sold
top_products = df.groupby('Description')['Quantity'].sum().sort_values(ascending=False).head(10)
print("\nTop 10 products by quantity sold:")
print(top_products)

# Top 10 countries by quantity sold
top_countries_qty = df.groupby('Country')['Quantity'].sum().sort_values(ascending=False).head(10)
print("\nTop 10 countries by quantity sold:")
print(top_countries_qty)

# Task 7: Identify Outliers
# Box plot for Quantity and UnitPrice
plt.figure(figsize=(12, 5))

plt.subplot(1, 2, 1)
sns.boxplot(y=df['Quantity'])
plt.title('Box Plot of Quantity')
plt.ylim(0, df['Quantity'].quantile(0.99))

plt.subplot(1, 2, 2)
sns.boxplot(y=df['UnitPrice'])
plt.title('Box Plot of UnitPrice')
plt.ylim(0, df['UnitPrice'].quantile(0.99))

plt.tight_layout()
plt.show()

# Identify outliers using IQR
Q1 = df[['Quantity', 'UnitPrice']].quantile(0.25)
Q3 = df[['Quantity', 'UnitPrice']].quantile(0.75)
IQR = Q3 - Q1
outliers = ((df[['Quantity', 'UnitPrice']] < (Q1 - 1.5 * IQR)) | 
            (df[['Quantity', 'UnitPrice']] > (Q3 + 1.5 * IQR))).any(axis=1)
print("\nNumber of rows with outliers:", outliers.sum())

# Task 8: Conclusions
print("\nSummary of Findings:")
print("- The dataset was cleaned by removing missing values, negative quantities/prices, and duplicates.")
print("- Sales are highest in certain months, indicating seasonal trends.")
print("- Specific days of the week show higher sales, suggesting targeted promotions.")
print("- Top-selling products and countries were identified, useful for inventory and marketing strategies.")
print("- Outliers were detected, which may represent bulk orders or errors needing further investigation.")
