# -*- coding: utf-8 -*-
"""
Created on Sat Jan 25 23:21:52 2025

@author: beniv
"""

from openpyxl.styles import numbers
import os
import pandas as pd
from datetime import datetime
import yaml 
skips=0
# Function to format date as DD/MM/YY
def format_date(date):
    """
    Formats a date string or datetime object as DD/MM/YY.
    If the date does not match the expected format, returns None.
    """
    if pd.isna(date):
        return None  # Skip rows with missing dates
    if isinstance(date, str):
        try:
            # Try parsing the date with the expected format
            date = datetime.strptime(date, "%Y-%m-%d %H:%M:%S")
        except ValueError:
            # If the format does not match, skip this row
            return None
    return date.strftime("%d/%m/%y")




# Main function
def main():
    # Get input for directory, column numbers, phrase, and classifications file, from params.yml
    with open('params.yml', 'r',encoding='utf-8') as file:
        params = yaml.safe_load(file)
    
    directory = params['going_in']['origin']  
    expense_name_col =  params['going_in']['expense_name_col']
    amount_col =  params['going_in']['amount_col']
    date_col = params['going_in']['date_col']
    phrase = params['going_in']['phrase']
    
    expense_summary_place= params['going_out']['save_place']
    summary_txt_name= params['going_out']['summary_txt_name']
    summary_xlsx_name= params['going_out']['summary_xlsx_name']

    # Load classifications from a TXT file if provided
    classifications_file = params['going_in']['expense_ctegories']
    classifications = load_classifications(classifications_file) 
    #
    try:
        # Process Excel files
        summary_df, detailed_df, category_expenses, classifications = process_excel_files(
            directory, expense_name_col, amount_col, date_col, phrase, classifications
        )

        # Round the numbers to the nearest whole number and convert to integers
        summary_df["Total Amount"] = summary_df["Total Amount"].round().astype(int)
        detailed_df["Amount"] = detailed_df["Amount"].round().astype(int)

        # Save results to Excel files
        
        with pd.ExcelWriter(os.path.join(expense_summary_place+summary_xlsx_name), engine='openpyxl') as writer:
            # Write the summary sheet
            summary_df.to_excel(writer, sheet_name="Summary", index=False)
            # Write the detailed breakdown sheet
            detailed_df.to_excel(writer, sheet_name="Detailed Breakdown", index=False)

            # Access the workbook and worksheet for formatting
            workbook = writer.book
            summary_sheet = writer.sheets["Summary"]
            detailed_sheet = writer.sheets["Detailed Breakdown"]

            # Apply custom number formatting to the "Total Amount" column in the Summary sheet
            for row in summary_sheet.iter_rows(min_row=2, min_col=2, max_col=2):  # Column B (Total Amount)
                for cell in row:
                    cell.number_format = '#,##0'  # Custom format: 1000 separator, no decimals

            # Apply custom number formatting to the "Amount" column in the Detailed Breakdown sheet
            for row in detailed_sheet.iter_rows(min_row=2, min_col=4, max_col=4):  # Column D (Amount)
                for cell in row:
                    cell.number_format = '#,##0'  # Custom format: 1000 separator, no decimals

        # Create a summary TXT file
        create_summary_txt(category_expenses, os.path.join(expense_summary_place+summary_txt_name))

        # Remove duplicate lines from the summary TXT file
        remove_duplicate_lines(os.path.join(expense_summary_place+summary_txt_name))

        print("\nExpense summary and detailed breakdown saved to 'expense_summary.xlsx'.")
        print("Expense classifications saved to 'expense_summary.txt'.")
        print("\n Number of lines skipped:",skips)
    except Exception as e:
        print(f"An error occurred: {e}")    
        
            
# Run the script
if __name__ == "__main__":
    main()
