import os

file_path = r'c:\data_cleaning_tool\web\index.html'
with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

new_content = content.replace("'南華靜觀'", "'南華'")

with open(file_path, 'w', encoding='utf-8') as f:
    f.write(new_content)

print("Successfully updated index.html with UTF-8 encoding.")
