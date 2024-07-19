import pandas as pd

'''
 @Author: Micah Chen
 Date: 7/11/2024
 Description: This script loads an Excel file, reorders the link IDs based on the head node,
 updates the DataFrame with the ordered links, and saves the updated DataFrame to a new Excel file.
 It ensures the original column structure is maintained and only displays the changed link IDs.
'''

# Load the uploaded Excel file
file_path = 'xxx.xlsx'
excel_data = pd.ExcelFile(file_path)

# Load the sheet 'in' into a DataFrame
df = excel_data.parse('in')

# Extracting 'Link IDs' columns for processing
link_id_columns = df.filter(regex='^Unnamed').fillna('')

# Function to reorder link IDs based on the head node
def reorder_link_ids(row):
    link_dict = {}
    head = row['Link IDs'].split('_')[1]
    valid_links = [link for link in row if isinstance(link, str) and '_' in link]  # Filter out non-string and invalid links
    
    for link in valid_links:
        parts = link.split('_')
        if len(parts) == 2:
            start, end = parts
            link_dict[start] = end
    
    ordered_links = []
    current = head
    visited = set()  # To keep track of visited nodes and avoid infinite loops
    while current in link_dict and current not in visited:
        visited.add(current)
        next_node = link_dict[current]
        ordered_links.append(f'{current}_{next_node}')
        current = next_node
    
    return ordered_links

# Applying the function to each row
df['Ordered Link IDs'] = df.apply(reorder_link_ids, axis=1)

# Initialize a list to hold the change logs
change_logs = []

# Updating the DataFrame with the ordered links
for index, row in df.iterrows():
    ordered_links = row['Ordered Link IDs']
    for i, link in enumerate(ordered_links):
        col_name = df.columns[i + 6]  # Existing columns starting from the 7th column (index 6)
        original_link = df.at[index, col_name]
        if original_link != link and pd.notna(original_link):
            df.at[index, col_name] = link
            change_logs.append(f"Row {index} (Segment: {row['Segment']}): Original: {original_link}, Changed: {link}")

# Drop the intermediate 'Ordered Link IDs' column
df = df.drop(columns=['Ordered Link IDs'])

# Save the updated DataFrame to a new Excel file
output_file_path = 'ph61112.xlsx'
df.to_excel(output_file_path, index=False)

print("Updated file saved as:", output_file_path)

# Write a summary report
with open('change_log.txt', 'w') as log_file:
    log_file.write('\n'.join(change_logs))
    log_file.write(f"\nTotal changes made: {len(change_logs)}")

print("Change log written to change_log.txt")
