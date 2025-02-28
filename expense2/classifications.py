# Function to load classifications from a yml file
# Returns as dictionary 

import os
import yaml

def load_classifications(filename):
    """
    Loads classifications from a yml file into a dictionary.
    """
    classifications = {}
    if os.path.exists(filename): #important
        with open(filename, 'r', encoding='utf-8') as file:
            data=yaml.safe_load(file)
    return data


