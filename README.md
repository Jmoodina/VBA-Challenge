# Project Completion Report: Stock Analysis Script

## Overview:
The goal of this project was to create a VBA script that analyzes stock data for multiple years. The script calculates and outputs the yearly change, percentage change, and total stock volume for each stock. Additionally, it identifies the stocks with the greatest percentage increase, greatest percentage decrease, and greatest total volume. The script has been designed to run on every worksheet (representing each year's data) and uses conditional formatting to highlight positive and negative changes.

## Execution
1. Iterate Through All Stocks for One Year:
The VBA script employs a looping mechanism to traverse through each stock's data for a specific year.
Calculates and outputs the yearly change, percentage change, and total stock volume for each stock.

2. Identify "Greatest" Metrics:
The script determines, for each year, the stocks with the highest percentage increase, greatest percentage decrease, and greatest total volume.
Stores and outputs the relevant information for these exceptional cases.

3. Run on Every Worksheet:
Adjustments were made to the script to ensure its compatibility with each worksheet, representing data for different years.
The script iterates through all sheets, performing the analysis consistently for each.

4. Apply Conditional Formatting:
Conditional formatting enhances the visual representation of results.
Positive changes are highlighted in green, while negative changes are highlighted in red, providing a quick and clear overview.

## Results:

The script successfully analyzes stock data for every year, providing comprehensive information on each stock's performance.
The identified "Greatest" metrics offer insights into notable stock trends for each year.

![2018 stock solution screenshot](https://github.com/Jmoodina/VBA-Challenge/assets/141544196/cdbb4c77-5ba6-4eb8-8019-b03c22fe8119)

![2019 Stock solution Screenshot](https://github.com/Jmoodina/VBA-Challenge/assets/141544196/b8a5ba5e-8c36-46b6-b75a-f0befe77f445)

![2020 Stock solution Screenshot](https://github.com/Jmoodina/VBA-Challenge/assets/141544196/1749b8e2-1f43-4536-92de-828276347f02)

## Performance:

The script has been optimized to run efficiently on the provided dataset (alphabetical_testing.xlsx), completing the analysis in under 3 to 5 minutes.
Consistency Across Sheets:

The script ensures consistent execution on every sheet, maintaining the same analysis structure for each year's data.

## Final Note:

The completed VBA script simplifies and accelerates the process of stock data analysis. Users can now execute the script with a click of a button, saving time and ensuring uniformity in the analysis across multiple years.
