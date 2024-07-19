import pandas as pd

"""
 @Author: Micah Chen
 Date: 7/11/2024
 Description: This script reorders the 'Link ID' values in an Excel file based on a specific matching criteria.
 The goal is to ensure that the last three digits of one link ID match the first three digits of the next link ID.
 The updated DataFrame is saved to a new Excel file, and a change log is generated to record all modifications.
"""

print("Loading Excel file...")

# Load the uploaded Excel file
file_path = 'original.xlsx'
excel_data = pd.ExcelFile(file_path)

print("Excel file loaded successfully.")

# Load the sheet 'Sheet4' into a DataFrame
print("Loading sheet 'Sheet4' into DataFrame...")
df = excel_data.parse('Sheet4')
print("Sheet 'Sheet4' loaded successfully.")

# Extracting 'Link IDs' columns for processing
print("Extracting 'Link IDs' columns for processing...")
link_id_columns = df.filter(regex='^Unnamed').fillna('')
print("'Link IDs' columns extracted successfully.")

# Function to reorder link IDs by matching the last three digits
def reorder_link_ids(row):
    link_dict = {}
    head = row['Link ID'].split('_')[0]  # Use the start of 'Link ID' as head
    valid_links = [link for link in row if isinstance(link, str) and '_' in link]  # Filter out non-string and invalid links
    
    # Create a dictionary of link mappings
    for link in valid_links:
        parts = link.split('_')
        if len(parts) == 2:
            start, end = parts
            link_dict[start] = end
    
    # Generate the ordered list of links
    ordered_links = []
    current = head
    while current in link_dict:
        next_node = link_dict[current]
        ordered_links.append(f'{current}_{next_node}')
        current = next_node
    
    return ordered_links

print("Reordering link IDs...")

# Apply the reordering function to each row
df['Ordered Link IDs'] = df.apply(reorder_link_ids, axis=1)
print("Link IDs reordered successfully.")

# Initialize a list to hold the change logs
change_logs = []

print("Updating DataFrame with ordered links...")

# Update the DataFrame with the ordered links
for index, row in df.iterrows():
    ordered_links = row['Ordered Link IDs']
    for i, link in enumerate(ordered_links):
        col_name = df.columns[i + 6]  # Existing columns starting from the 7th column (index 6)
        if i + 6 < len(df.columns):
            original_link = df.at[index, col_name]
            if original_link != link and pd.notna(original_link):
                df.at[index, col_name] = link
                change_logs.append(f"Row {index} (Segment: {row['Segment']}): Original: {original_link}, Changed: {link}")

print("DataFrame updated successfully.")

# Drop the intermediate 'Ordered Link IDs' column
df = df.drop(columns=['Ordered Link IDs'])

# Save the updated DataFrame to a new Excel file
output_file_path = 'ordered.xlsx'
df.to_excel(output_file_path, index=False)
print(f"Updated DataFrame saved to {output_file_path}.")

# Write a summary report
change_log_path = 'change_log_fixed.txt'
with open(change_log_path, 'w') as log_file:
    log_file.write('\n'.join(change_logs))
    log_file.write(f"\nTotal changes made: {len(change_logs)}")
print(f"Change log written to {change_log_path}.")

# Output file paths for confirmation
print(f"Process completed successfully. Files saved: {output_file_path}, {change_log_path}")
output_file_path, change_log_path
