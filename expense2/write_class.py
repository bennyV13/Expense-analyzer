import yaml

# Function to create a summary TXT file
def create_summary_yaml(classifications, filename):
    # Write the data to a YAML file with UTF-8 encoding
    with open(filename, 'w', encoding='utf-8') as file:
        yaml.dump(classifications, file, allow_unicode=True, sort_keys=False)
