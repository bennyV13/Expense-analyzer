import pandas as pd
import yaml


# Function to process a single Excel file
# Receives
    # filepath for the Excel file
    # expense_name_col, amount_col, date_col
    # classifications from yaml file
# Returns
    # updated classifications
    # sum of expenses by category
    # detailed breakdown of expenses 

    

def process_single_file(filepath, expense_name_col, amount_col, date_col, classifications):
    """
    Processes a single Excel file and returns the summary and detailed breakdown.
    """

    try:
        print(f"\nProcessing file: {filepath}")

        # Read the Excel file
        df = pd.read_excel(filepath, header=None)
        
        # Ensure the required columns exist
        required_columns = [int(expense_name_col),int(amount_col), int(date_col)]
        if all(isinstance(col, int) and col <= len(df.columns) for col in required_columns):  
            # Classify expenses
            detailed_results = []  # Stores detailed results
            summed_expenses = {}  # Stores summary of expenses by category
            # go throgh each row of the dataframe, check if the 
            # expense name has already been classified, 
            # if not, prompt the user to classify the expense
                # add new classification to the classifications dictionary
            # add the expense to the detailed results 
            # and update the summary
                 

            
            return summed_expenses, detailed_results, classifications
        else:
            print(f"Missing required columns in file: {filepath}")
            return {}, [], {}, classifications
    except Exception as e:
        print(f"Error in process_single_file for file {filepath}: {e}")
        return {}, [], {}, classifications

# Function to classify expenses
def classify_expenses(df, expense_name_col, amount_col, date_col, classifications):
    """
    Classifies expenses based on user input and generates a summary and detailed breakdown.
    """
    
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
