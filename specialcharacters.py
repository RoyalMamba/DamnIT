import re
import os

def remove_special_characters(input_string):
    # Replace special characters with underscores
    clean_string = re.sub(r'[^a-zA-Z0-9]', '_', input_string)
    return clean_string

# Split image name into base name and extension
base_name, extension = os.path.splitext(image_name)

# Remove special characters from the base name
cleaned_base_name = remove_special_characters(base_name)

# Combine cleaned base name and extension
cleaned_image_name = cleaned_base_name + extension


