import pandas as pd

# Load the uploaded Excel file
file_path = 'original.xlsx'
excel_data = pd.ExcelFile(file_path)

# Load the sheet 'Sheet4' into a DataFrame
df = excel_data.parse('Sheet4')

# Extracting 'Link IDs' columns for processing
link_id_columns = df.filter(regex='^Unnamed').fillna('')

# Function to reorder link IDs by matching last three digits
def reorder_link_ids(row):
    link_dict = {}
    head = row['Link ID'].split('_')[0]  # Adjusted to use the start of 'Link ID' as head
    valid_links = [link for link in row if isinstance(link, str) and '_' in link]  # Filter out non-string and invalid links
    
    for link in valid_links:
        parts = link.split('_')
        if len(parts) == 2:
            start, end = parts
            link_dict[start] = end
    
    ordered_links = []
    current = head
    while current in link_dict:
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
        if i + 6 < len(df.columns):
            original_link = df.at[index, col_name]
            if original_link != link and pd.notna(original_link):
                df.at[index, col_name] = link
                change_logs.append(f"Row {index} (Segment: {row['Segment']}): Original: {original_link}, Changed: {link}")

# Drop the intermediate 'Ordered Link IDs' column
df = df.drop(columns=['Ordered Link IDs'])

# Save the updated DataFrame to a new Excel file
output_file_path = 'ordered.xlsx'
df.to_excel(output_file_path, index=False)

# Write a summary report
change_log_path = 'change_log_fixed.txt'
with open(change_log_path, 'w') as log_file:
    log_file.write('\n'.join(change_logs))
    log_file.write(f"\nTotal changes made: {len(change_logs)}")

output_file_path, change_log_path
