import yaml
import os

def load_expense_summary(filename):
    """
    Reads a categorized text file in Hebrew and returns a dictionary of categories with unique items.
    """
    categories = {}
    current_category = None

    if not os.path.exists(filename):
        print(f"File not found: {filename}")
        return categories  # Return empty if file doesn't exist

    with open(filename, "r", encoding="utf-8") as file:
        for line in file:
            line = line.strip()
            if not line:
                continue  # Skip empty lines

            if line.endswith(":"):  # Category line
                current_category = line[:-1]
                if current_category not in categories:
                    categories[current_category] = set()
            elif current_category and (line.startswith("- ") or line.startswith("  - ")):  # Item line
                item = line[2:].strip()  # Remove bullet point
                categories[current_category].add(item)  # Store in set to avoid duplicates

    return categories

def merge_expense_summaries(file1, file2, output_file):
    """
    Merges two categorized expense lists in Hebrew and removes duplicate items within each category.
    """
    # Load data from both files
    summary1 = load_expense_summary(file1)
    summary2 = load_expense_summary(file2)

    # Merge both summaries
    merged_summary = {}
    for category in set(summary1.keys()).union(summary2.keys()):
        merged_summary[category] = summary1.get(category, set()).union(summary2.get(category, set()))

    # Write the merged data back to a file
    with open(output_file, "w", encoding="utf-8") as out_file:
        for category, items in sorted(merged_summary.items(), key=lambda x: x[0], reverse=True):  # Sort categories (Hebrew-aware)
            out_file.write(f"{category}:\n")
            for item in sorted(items, key=lambda x: x[::-1]):  # Sort items correctly in Hebrew
                out_file.write(f"  - {item}\n")
            out_file.write("\n")  # Space between categories for readability

    print(f"✅ Merged file saved as '{output_file}'.")

def main():
    with open('params.yml', 'r',encoding='utf-8') as file:
        params = yaml.safe_load(file)
    file1=params['merging']['file1']
    file2=params['merging']['file2']
    output_file=params['merging']['output_file']
    merge_expense_summaries(file1, file2, output_file)

if __name__ == "__main__":
    main()