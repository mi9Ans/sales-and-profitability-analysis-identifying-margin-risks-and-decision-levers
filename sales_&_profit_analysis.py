"""
Sales & Profitability Analysis
Author: Anshumaan Mishra

This script performs a structured analysis of sales, profit, discount,
and margin drivers using the Superstore dataset. The focus is on identifying
structural profitability risks and decision levers.
"""

import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import os

try:
    # GitHub / local project structure
    df = pd.read_excel("data/Sample - Superstore.xls")
except FileNotFoundError:
    # Google Colab upload path
    df = pd.read_excel("/content/Sample - Superstore.xls")

df['Order Date'] = pd.to_datetime(df['Order Date'])
df['Ship Date'] = pd.to_datetime(df['Ship Date'])
df['Delivery Time'] = (df['Ship Date'] - df['Order Date']).dt.days

total_sales = df['Sales'].sum()
total_profit = df['Profit'].sum()
average_discount = df['Discount'].mean() * 100
profit_margin = (total_profit / total_sales) * 100

print(f"Total Sales: ${total_sales:,.2f}")
print(f"Total Profit: ${total_profit:,.2f}")
print(f"Average Discount: {average_discount:.2f}%")
print(f"Profit Margin: {profit_margin:.2f}%")

regional_summary = (
    df.groupby('Region')
    .agg(Sales = ('Sales', 'sum'),
         Profit = ('Profit', 'sum'),
         Avg_Discount = ('Discount', 'mean'))
    .reset_index()
)

regional_summary['Profit_Margin'] = (
    (regional_summary['Profit'] / regional_summary['Sales']) * 100
)

regional_summary.plot(kind='line', x='Region', y='Profit_Margin',color='skyblue', marker='o')
plt.title('Profit Margin by Region')
plt.xlabel('Region')
plt.ylabel('Profit Margin (%)')
plt.grid(axis='y', linestyle='--', alpha=0.4)

plt.show()

category_summary = (
    df.groupby('Category')
      .agg(Sales=('Sales', 'sum'),
           Profit=('Profit', 'sum'),
           Avg_Discount=('Discount', 'mean'))
      .reset_index()
)
category_summary['Profit_Margin'] = (
    (category_summary['Profit'] / category_summary['Sales']) * 100
)

category_summary.plot(kind='bar', x='Category', y='Profit', color='royalblue')
plt.title('Total Profit by Category')
plt.xlabel('Category')
plt.ylabel('Total Profit')
plt.show()

furniture_df = df[df['Category'] == 'Furniture']

subcat_summary = (
    furniture_df.groupby('Sub-Category')
    .agg(Sales=('Sales', 'sum'),
         Profit=('Profit', 'sum'))
    .reset_index()
)

sns.scatterplot(data=subcat_summary, x='Sales', y='Profit',
                hue='Sub-Category', size='Profit', sizes=(50, 500),
                legend=False)

plt.title('Sales vs Profit by Sub-Category')
plt.xlabel('Sales')
plt.ylabel('Profit')
plt.tight_layout()
plt.show()

