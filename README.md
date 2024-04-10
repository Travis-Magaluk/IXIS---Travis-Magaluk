# IXIS - Travis Magaluk
 Showcase of Coding Skills

# Website Performance Analysis

This project aims to analyze the performance of an online retailer's website using Google Analytics (GA) data. It involves processing two GA datasets containing basic ecommerce metrics to summarize website performance and provide insights to help the retailer understand their website’s performance.

## Project Scope

Imagine that you have been tasked with helping an online retailer measure and analyze their website’s performance. The retailer uses Google Analytics (GA) to track engagement and conversion on their website. You have been given two GA datasets containing basic ecommerce metrics and asked to summarize website performance and provide insights to help the retailer understand their website’s performance.

## Datasets

- `sessionCounts.csv`: Contains sessions, transactions, and quantity broken out by browser, device category, and date.
- `addsToCart.csv`: Contains adds to cart broken out by month.

## Analysis

The project involves the following steps:

1.**EDA Was Done but not uploaded**
2. **Data Cleaning**: The provided datasets are cleaned and processed to ensure consistency and accuracy in the analysis.
3. **Data Aggregation**: The data is aggregated to summarize website performance metrics such as sessions, transactions, quantity, and e-commerce conversion rate (ECR) by month, device category, and browser.
4. **Month-to-Month Comparison**: A comparison is made between the most recent two months' data to identify trends and changes in website performance metrics.
5. **Top Browsers Analysis**: The top 20 browsers are identified based on sessions, and their performance metrics are analyzed.
6. **Excel Output**: The aggregated data and analysis results are exported to an Excel file for further analysis and reporting.
7. **Generating Visualizations**: all_together.ipynb has the main code, but also includes code for generating some visualizations. 

## Usage

To run the analysis and generate the Excel output, follow these steps:

1. Clone the repository:

'''git clone https://github.com/Travis-Magaluk/website-performance-analysis.git'''

2. Navigate to the project directory

3. Install requirements.

4. Run 'travis_magaluk_coding_exercise.py' with the paths to the CSV files as command line arguments

'''python main.py sessionCounts.csv addsToCart.csv'''

5. a
