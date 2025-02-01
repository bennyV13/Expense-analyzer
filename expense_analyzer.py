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

# Function to load classifications from a TXT file
def load_classifications(filename):
    """
    Loads classifications from a TXT file into a dictionary.
    """
    classifications = {}
    if os.path.exists(filename): #important
        with open(filename, "r", encoding="utf-8") as file:
            current_category = None
            for line in file:
                line = line.strip()
                if line.endswith(":"):  # Category line
                    current_category = line[:-1]  # Remove the colon
                elif line.startswith("  - ") or line.startswith("- "):  # Expense line
                    expense_name = line[4:] if line.startswith("  - ") else line[2:]
                    classifications[expense_name] = current_category
    return classifications

# Function to find the starting row after a specific phrase
def find_start_row(filepath, phrase):
    """
    Finds the starting row (two rows after the specified phrase) in an Excel file.
    """
    df = pd.read_excel(filepath, header=None)  # Read the file without headers
    for index, row in df.iterrows():
        if any(phrase in str(cell) for cell in row):  # Check if the phrase exists in the row
            return index + 2  # Start two rows after the phrase (index is zero-based)
    return 0  # Default to row 0 if the phrase is not found

# Function to classify expenses
def classify_expenses(df, expense_name_col, amount_col, date_col, classifications):
    """
    Classifies expenses based on user input and generates a summary and detailed breakdown.
    """
    category_expenses = {}  # Maps categories to lists of expense names
    results = []  # Stores detailed results
    summary = {}  # Stores summary of expenses by category
    last_expense = None  # Stores the last classified expense for "back" functionality
    global skips
    for index, row in df.iterrows():
        expense_name = row[expense_name_col - 1]  # Convert to 0-based index
        amount = row[amount_col - 1]              # Convert to 0-based index
        date = format_date(row[date_col - 1])     # Convert to 0-based index

        # Debugging: Print the row being processed
        print(f"\nProcessing Row {index + 1}:")
        print(f"Expense Name: {expense_name}")
        print(f"Amount: {amount}")
        print(f"Date: {date}")
        if index==263:
            print("Stop here")

        # Skip rows with missing or invalid dates
        if date is None:
            skips=skips+1
            print("Skipping row: Invalid date.")
            continue

        # Skip rows with missing expense names
        if pd.isna(expense_name):
            skips=skips+1
            print("Skipping row: Missing expense name.")
            continue

        # Check if the expense name has already been classified
        if expense_name in classifications:
            category = classifications[expense_name]
            print(f"Expense '{expense_name}' classified as '{category}'.")
        else:
            # Prompt the user to classify the expense
            while True:
                print(f"\nExpense: {expense_name}")
                print("Existing categories:", ", ".join(summary.keys()) if summary else "None")
                category = input("Enter an existing category, a new one, or 'back' to edit the last input: ").strip()

                if category.lower() == "back":
                    if last_expense:
                        expense_name = last_expense["Expense Name"]
                        amount = last_expense["Amount"]
                        date = last_expense["Date"]
                        results.pop()  # Remove the last result
                        if last_expense["Category"] in summary:
                            summary[last_expense["Category"]] -= last_expense["Amount"]
                            if summary[last_expense["Category"]] == 0:
                                del summary[last_expense["Category"]]
                    else:
                        print("No previous expense to edit.")
                    continue
                else:
                    break  # Use the category as entered by the user

            # Add the classification to the dictionary
            classifications[expense_name] = category
            print(f"Added classification: '{expense_name}' -> '{category}'.")

        # Add the expense name to the category_expenses dictionary
        if category not in category_expenses:
            category_expenses[category] = []
        category_expenses[category].append(expense_name)

        # Add the expense to the detailed results
        results.append({
            "Category": category,
            "Date": date,
            "Expense Name": expense_name,
            "Amount": amount
        })

        # Update the summary
        if category in summary:
            summary[category] += amount
        else:
            summary[category] = amount

        # Store the last classified expense for "back" functionality
        last_expense = results[-1]

    return summary, results, category_expenses, classifications

# Function to process a single Excel file
def process_single_file(filepath, expense_name_col, amount_col, date_col, phrase, classifications):
    """
    Processes a single Excel file and returns the summary and detailed breakdown.
    """
    try:
        # Find the starting row based on the phrase
        start_row = find_start_row(filepath, phrase)
        print(f"\nProcessing file: {filepath}")
        print(f"Start row: {start_row}")

        # Read the Excel file, skipping rows before the starting row
        df = pd.read_excel(filepath, skiprows=start_row)
        print(f"DataFrame Columns: {df.columns}")
        print(f"DataFrame Head:\n{df.head()}")

        # Ensure the required columns exist
        required_columns = [int(expense_name_col),int(amount_col), int(date_col)]
        if all(isinstance(col, int) and col <= len(df.columns) for col in required_columns):  # Check if column numbers are valid
            # Classify expenses
            summary, detailed_results, category_expenses, classifications = classify_expenses(
                df,
                expense_name_col,
                amount_col,
                date_col,
                classifications
            )
            return summary, detailed_results, category_expenses, classifications
        else:
            print(f"Missing required columns in file: {filepath}")
            return {}, [], {}, classifications
    except Exception as e:
        print(f"Error in process_single_file for file {filepath}: {e}")
        return {}, [], {}, classifications

# Function to process Excel files
def process_excel_files(directory, expense_name_col, amount_col, date_col, phrase, classifications):
    """
    Reads all Excel files from the specified directory, processes them separately, and aggregates the results.
    """
    all_summaries = []
    all_detailed_results = []
    all_category_expenses = {}

    # Read all Excel files in the directory
    for filename in os.listdir(directory):
        if filename.endswith(".xlsx"):
            filepath = os.path.join(directory, filename)
            print(f"\nProcessing file: {filename}")

            summary, detailed_results, category_expenses, classifications = process_single_file(
                filepath, expense_name_col, amount_col, date_col, phrase, classifications
            )

            # Aggregate results
            all_summaries.append(summary)
            all_detailed_results.extend(detailed_results)
            for category, expenses in category_expenses.items():
                if category not in all_category_expenses:
                    all_category_expenses[category] = []
                all_category_expenses[category].extend(expenses)

    if not all_summaries:
        raise ValueError("No valid Excel files found in the directory.")

    # Combine summaries
    combined_summary = {}
    for summary in all_summaries:
        for category, amount in summary.items():
            if category in combined_summary:
                combined_summary[category] += amount
            else:
                combined_summary[category] = amount

    # Convert results to DataFrames for easier saving
    summary_df = pd.DataFrame(list(combined_summary.items()), columns=["Category", "Total Amount"])
    detailed_df = pd.DataFrame(all_detailed_results)

    # Sort the detailed breakdown by Category and Expense Name
    detailed_df = detailed_df.sort_values(by=["Category", "Expense Name"])

    return summary_df, detailed_df, all_category_expenses, classifications

# Function to create a summary TXT file
def create_summary_txt(category_expenses, filename):
    """
    Creates a TXT file summarizing all expense names under their classifications,
    with an empty line separating each category.
    """
    with open(filename, "w", encoding="utf-8") as file:
        for category, expenses in category_expenses.items():
            file.write(f"{category}:\n")
            for expense in sorted(expenses):  # Convert to sorted list for consistent output
                file.write(f"  - {expense}\n")
            file.write("\n")  # Add an empty line after each category

def remove_duplicate_lines(filename):
    """
    Removes duplicate lines from a file and rewrites the file with unique lines,
    while preserving empty lines between categories.
    """
    if os.path.exists(filename):
        with open(filename, "r", encoding="utf-8") as file:
            lines = file.readlines()

        unique_lines = []
        seen_lines = set()
        for line in lines:
            # Keep empty lines as they are and ensure other lines are unique
            if line.strip() == "" or line not in seen_lines:
                unique_lines.append(line)
                seen_lines.add(line)

        with open(filename, "w", encoding="utf-8") as file:
            file.writelines(unique_lines)

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
    classifications = load_classifications(classifications_file) if classifications_file else {}
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
