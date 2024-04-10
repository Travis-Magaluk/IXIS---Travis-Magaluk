import pandas as pd
import sys
import openpyxl


def clean_tables(session_cts, cart_adds):
    """
    Clean session count and cart adds dataframes by extracting month and year information,
    mapping month numbers to month names, and ensuring consistent ordering of month names.

    Args:
        session_cts (DataFrame): DataFrame containing session count data.
        cart_adds (DataFrame): DataFrame containing cart adds data.

    Returns:
        tuple: A tuple containing the cleaned session count and cart adds DataFrames.
    """

    # Define a function to map month numbers to month names
    def map_month(month_number):
        return pd.to_datetime(month_number, format='%m').strftime('%B')

    # Extract month and year information from session_cts dataframe
    session_cts['Month'] = session_cts['dim_date'].str.split('/').str[0]
    session_cts['Year'] = session_cts['dim_date'].str.split('/').str[2]
    session_cts['Year'] = '20' + session_cts['Year'].apply(lambda x: x.zfill(2))  # Ensure year has 4 digits
    session_cts['Year'] = session_cts['Year'].astype(int)  # Convert year to integer
    session_cts['month_name'] = session_cts['Month'].apply(map_month)  # Map month numbers to month names
    session_cts['dim_deviceCategory'] = session_cts['dim_deviceCategory'].str.title()  # Capitalize device categories

    # Extract month information from cart_adds dataframe
    cart_adds['dim_month'] = cart_adds['dim_month'].astype(str)
    cart_adds['month_name'] = cart_adds['dim_month'].apply(map_month)

    # Ensure consistent ordering of month names
    month_order = list(session_cts['month_name'].unique())

    # Apply categorical ordering to month_name column
    session_cts['month_name'] = pd.Categorical(session_cts['month_name'], categories=month_order, ordered=True)
    cart_adds['month_name'] = pd.Categorical(cart_adds['month_name'], categories=month_order, ordered=True)

    # Rename dim_year column to Year in cart_adds dataframe
    cart_adds.rename(mapper={'dim_year': 'Year'}, axis=1, inplace=True)

    return session_cts, cart_adds


def filter_last_months(df):
    """
    Filter the DataFrame to retain only the rows corresponding to the last two months.

    Args:
        df (DataFrame): DataFrame containing data with columns 'Year' and 'month_name'.

    Returns:
        DataFrame: Filtered DataFrame containing only the rows corresponding to the last two months.
    """
    # Sort the DataFrame by 'Year' and 'month_name' to ensure chronological order
    df = df.sort_values(by=['Year', 'month_name'])

    # Extract the last two unique month names
    last_months = df['month_name'].unique()[-3:]

    # Filter the DataFrame to retain only the rows corresponding to the last two months
    filtered_df = df[df['month_name'].isin(last_months)]

    return filtered_df


def create_month_device(md):
    """
    Create a new DataFrame by aggregating data based on month and device category,
    and calculating the ECR (ecommerce conversion rate) for each group.

    Note: Used provided dataset which was July 2012 - June 2013 Data.
    I ordered months in order based on year and month.
    In future would add functionality
    to pull last 12 months of data and order in months
    based the data no matter which month was current.

    Args:
        md (DataFrame): DataFrame containing data with columns 'month_name', 'dim_deviceCategory',
                        'sessions', 'transactions', and 'QTY'.

    Returns:
        DataFrame: A new DataFrame containing aggregated data for each month and device category,
                   along with the calculated ECR.
    """
    # Define aggregation functions for each column
    aggregation_functions = {
        'sessions': 'sum',
        'transactions': 'sum',
        'QTY': 'sum'
    }

    # Define the desired order of months
    month_order = ['July', 'August', 'September', 'October', 'November',
                   'December', 'January', 'February', 'March',
                   'April', 'May', 'June']

    # Convert 'month_name' column to categorical with specified order
    md['month_name'] = pd.Categorical(md['month_name'], categories=month_order, ordered=True)

    # Group data by 'month_name' and 'dim_deviceCategory', then aggregate using defined functions
    grouped_md = md.groupby(['month_name', 'dim_deviceCategory']).agg(aggregation_functions)

    # Calculate ECR (E-commerce Conversion Rate) by dividing 'transactions' by 'sessions'
    grouped_md['ECR'] = grouped_md['transactions'] / grouped_md['sessions']

    return grouped_md


def create_browser_device(md):
    """
    Create a new DataFrame by aggregating data based on month and device category,
    and calculating the ECR (e-commerce conversion rate) for each group.

    Note: Used provided dataset which was July 2012 - June 2013 Data.
    I ordered months in order based on year and month.
    In future would add functionality
    to pull last 12 months of data and order in months
    based the data no matter which month was current.

    Args:
        md (DataFrame): DataFrame containing data with columns 'dim_deviceCategory',
                        'sessions', 'transactions', and 'QTY'.

    Returns:
        DataFrame: A new DataFrame containing aggregated data for each month and device category,
                   along with the calculated ECR.
    """
    # Define aggregation functions for each column
    aggregation_functions = {
        'sessions': 'sum',
        'transactions': 'sum',
        'QTY': 'sum'
    }

    # Group data by 'month_name' and 'dim_deviceCategory', then aggregate using defined functions
    grouped_md = md.groupby(['dim_browser']).agg(aggregation_functions)

    # Calculate ECR (E-commerce Conversion Rate) by dividing 'transactions' by 'sessions'
    grouped_md['ECR'] = grouped_md['transactions'] / grouped_md['sessions']

    return grouped_md


def create_mtm_compare(session_counts, adds):
    """
    Create a DataFrame for month-to-month comparison of session counts and cart adds data.

    Next steps with this function would be to impliment a number of months feature,
    or select months and years to compare. This could be valuable for reoccuring reporting.

    Args:
        session_counts (DataFrame): DataFrame containing session count data with columns 'month_name',
                                     'sessions', 'transactions', and 'QTY'.
        adds (DataFrame): DataFrame containing cart adds data with columns 'month_name' and 'addsToCart'.

    Returns:
        DataFrame: A DataFrame containing month-to-month comparison metrics for session counts and cart adds.
    """
    # Filter session_counts to retain only the last three months
    last_three = filter_last_months(session_counts)

    # Define aggregation functions for session_counts
    aggregation_functions = {
        'sessions': 'sum',
        'transactions': 'sum',
        'QTY': 'sum'
    }

    # Group last_three by 'month_name' and aggregate session counts data
    last_three_grouped = last_three.groupby(['month_name']).agg(aggregation_functions)

    # Calculate ECR (E-commerce Conversion Rate) for last_three_grouped
    last_three_grouped['ECR'] = last_three_grouped['transactions'] / last_three_grouped['sessions']

    # Fill NaN values with 0 and remove rows with all 0 values
    last_three_grouped.fillna(0, inplace=True)
    last_three_grouped = last_three_grouped.loc[(last_three_grouped != 0).all(axis=1)]

    # Merge adds data with last_three_grouped on 'month_name'
    last_three_grouped = last_three_grouped.merge(adds[['month_name', 'addsToCart']], on=['month_name'], how='inner')

    # Calculate absolute and percent changes for each metric
    for col in ['sessions', 'transactions', 'QTY', 'ECR', 'addsToCart']:
        # Absolute change
        last_three_grouped[f'{col}_abs_change'] = last_three_grouped[col].diff()

        # Percent change
        last_three_grouped[f'{col}_pct_change'] = last_three_grouped[col].pct_change() * 100

    # Rearrange columns and set 'month_name' as index
    last_three_grouped = last_three_grouped[
        ['month_name', 'sessions', 'sessions_abs_change', 'sessions_pct_change',
         'addsToCart', 'addsToCart_abs_change', 'addsToCart_pct_change',
         'transactions', 'transactions_abs_change', 'transactions_pct_change',
         'QTY', 'QTY_abs_change', 'QTY_pct_change',
         'ECR', 'ECR_abs_change', 'ECR_pct_change']].set_index('month_name')

    # Transpose DataFrame for better readability
    last_three_grouped = last_three_grouped.transpose()

    return last_three_grouped


def create_month_device_totals(month_device):
    """
    Create a DataFrame with total values for each device category across all months.

    Args:
        month_device (DataFrame): DataFrame containing monthly data for different device categories.
                                Cols needed: month_name, dim_deviceCategory

    Returns:
        DataFrame: DataFrame with total values for each device category across all months, including a 'Total' row for each month.
    """
    # Reset index to ensure 'month_name' becomes a regular column
    month_device = month_device.reset_index()

    # Calculate totals for each month across all device categories
    totals = month_device.groupby('month_name').sum().reset_index()

    # Add a 'Total' category to identify total values
    totals['dim_deviceCategory'] = 'Total'

    # Concatenate original data with total rows
    month_device_with_totals = pd.concat([month_device, totals], ignore_index=False)

    # Define the order of device categories (including 'Total') for sorting
    order = ['Desktop', 'Mobile', 'Tablet', 'Total']

    # Convert 'dim_deviceCategory' to categorical and enforce the specified order
    month_device_with_totals['dim_deviceCategory'] = pd.Categorical(month_device_with_totals['dim_deviceCategory'],
                                                                    categories=order, ordered=True)

    # Sort the DataFrame by 'month_name' and 'dim_deviceCategory'
    month_device_with_totals = month_device_with_totals.sort_values(
        by=['month_name', 'dim_deviceCategory']).reset_index()

    # Remove the index column
    month_device_with_totals.drop('index', inplace=True, axis=1)

    return month_device_with_totals


def generate_excel(session_counts, adds_to_cart):

    # Reading in CSV Files
    sc = pd.read_csv(session_counts)
    adds = pd.read_csv(adds_to_cart)

    # Running the clean tables function on both tables
    sc, adds = clean_tables(sc, adds)

    # Creating a table aggregating the sessions data by Month * Device
    # and creating a ECR Metric
    month_device_agg = create_month_device(sc)

    # Creating a table aggregating data by browser and getting the top 20
    browser = create_browser_device(sc)
    browser = browser.sort_values('sessions', ascending=False)[:20]

    # Creating the month-to-month comparison table
    mtm_compare = create_mtm_compare(sc, adds)

    # Adding a totals row under each month for ease of viewing and editing for visualizations
    month_device_with_totals = create_month_device_totals(month_device_agg)

    # Writing the four tables to an Excel file
    with pd.ExcelWriter('website_agg.xlsx') as writer:
        # Write each DataFrame to a separate sheet
        month_device_agg.to_excel(writer, sheet_name='Month Device Agg', index=True)
        mtm_compare.to_excel(writer, sheet_name='Month to Month Comparison', index=True)
        browser.to_excel(writer, sheet_name='Top 20 Browsers', index=True)
        month_device_with_totals.to_excel(writer, sheet_name='Month Aggs with Total', index=True)


def main():
    try:
        # Check if the correct number of arguments are provided
        if len(sys.argv) != 3:
            raise ValueError("Usage: python main.py <session_counts_csv> <adds_to_cart_csv>")

        # Extract the command-line arguments
        session_counts = sys.argv[1]
        adds_to_cart = sys.argv[2]

        # Call the generate_excel function with the provided arguments
        generate_excel(session_counts, adds_to_cart)

        print("Excel file generated successfully.")

    except Exception as e:
        print("An error occurred:", e)


if __name__ == "__main__":
    main()