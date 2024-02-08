import yfinance as yf
import pandas as pd
import os
import xlsxwriter
import pdb

# Define a mapping from quarter strings to their end-of-quarter dates
quarter_mapping = {
    "20 Q1": "2020-03-31",
    "20 Q2": "2020-06-30",
    "20 Q3": "2020-09-30",
    "20 Q4": "2020-12-31",
    "21 Q1": "2021-03-31",
    "21 Q2": "2021-06-30",
    "21 Q3": "2021-09-30",
    "21 Q4": "2021-12-31",
    "22 Q1": "2022-03-31",
    "22 Q2": "2022-06-30",
    "22 Q3": "2022-09-30",
    "22 Q4": "2022-12-31",
    "23 Q1": "2023-03-31",
    "23 Q2": "2023-06-30",
    "23 Q3": "2023-09-30",
    "23 Q4": "2023-12-31",
}

# Mapping shortened metric strings to their yfinance API names
metric_mapping = {
    "revenue": "Total Revenue",
    "total expense": "Total Expenses",
    "earnings": "Net Income",
    "EPS": "Basic EPS",
    "operating income": "Operating Income",
    "net income": "Net Income",

    # Below values are part of the 'quarterly_incomestmt' or 'quarterly_income_stmt' attributes
    # "cash and equivalents": "Cash And Cash Equivalents",
    # "total assets": "Total Assets",  # You would need to find the exact name for this one
    # Add more mappings as needed
}

# Define the function
def GenerateMetrics(tickers_list, metrics_list, quarters_list, document_name):
    # Initialize a list to hold the data for all tickers
    all_data = []
    
    # Loop through each ticker
    for ticker in tickers_list:
        ticker_data = yf.Ticker(ticker)
        
        # Create a DataFrame to hold the data for this ticker
        ticker_df = pd.DataFrame(index=metrics_list, columns=quarters_list)
        
        # Fill the DataFrame with 'Ticker not found' if we can't get data
        if ticker_data.quarterly_financials is None:
            ticker_df.loc[:, :] = "Ticker not found"
        else:
            for quarter in quarters_list:
                quarter_end_date = quarter_mapping.get(quarter, "Quarter not found")
                if quarter_end_date == "Quarter not found":
                    ticker_df[quarter] = "Quarter not found"
                    continue

                for metric in metrics_list:
                    pdb.set_trace()
                    mapped_metric = metric_mapping.get(metric.lower(), "Metric not found")
                    if mapped_metric == "Metric not found":
                        ticker_df.loc[metric, quarter] = "Metric not found"
                    else:
                        value = ticker_data.quarterly_financials.get(quarter_end_date, {}).get(mapped_metric, "Metric not found")
                        ticker_df.loc[metric, quarter] = value

        # Add the ticker's data to the DataFrame
        ticker_df = ticker_df.reset_index()
        ticker_df.insert(0, 'Ticker', ticker)
        ticker_df.rename(columns={'index': 'Metric'}, inplace=True)
        
        # Append an empty row to the DataFrame for the separator
        separator = pd.DataFrame({col: '' for col in ticker_df.columns}, index=[len(ticker_df)])
        ticker_df = pd.concat([ticker_df, separator], ignore_index=True)
        
        # Add the ticker DataFrame with the empty row to the list
        all_data.append(ticker_df)

    # Combine all the individual DataFrames into one
    formatted_data = pd.concat(all_data, ignore_index=True)
    
    # Define the full file path
    directory = os.path.expanduser('~/Documents/Financial forecasting/Stock data')
    
    # Create the directory if it does not exist
    if not os.path.exists(directory):
        os.makedirs(directory)
    file_path = os.path.join(directory, document_name + '.xlsx')

    # Prepare the Excel writer with the defined file path
    writer = pd.ExcelWriter(file_path, engine='xlsxwriter')
    formatted_data.to_excel(writer, sheet_name='Sheet1', index=False)
    
    # Get the xlsxwriter workbook and worksheet objects
    workbook  = writer.book
    worksheet = writer.sheets['Sheet1']

    # Define a format for the grey horizontal bar
    format_grey_bar = workbook.add_format({'bg_color': '#D3D3D3'})
    
    # Calculate total columns for the grey bar span
    total_columns = len(formatted_data.columns)

    # Apply the format to the required rows
    for i in range(1,len(tickers_list)+1):
        # # The grey bar row is after each ticker's metrics, hence (len(metrics_list) + 1)
        # worksheet.set_row(i * (len(metrics_list)), None, format_grey_bar)

        # Apply the grey bar format to the range that matches the width of the data
        for col_num in range(total_columns):
            # Convert column number to Excel column letter
            col_letter = xlsxwriter.utility.xl_col_to_name(col_num)
            # Apply format to the cell that should be grey
            worksheet.write(f'{col_letter}{i * (len(metrics_list) + 1) + 1}', None, format_grey_bar)

    # Close the Pandas Excel writer and output the Excel file
    writer.close()

# Example usage
stocks_list = ["NVDA", "SMCI", "AMD"]
metrics_list = ["total expense", "earnings", "EPS"]
quarters_list = ["21 Q1", "23 Q2", "23 Q3", "23 Q4"]
document_name = "NVDA_SMCI_AMD"

GenerateMetrics(stocks_list, metrics_list, quarters_list, document_name)