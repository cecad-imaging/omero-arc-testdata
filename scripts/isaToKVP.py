import pandas as pd
import sys

def read_first_isa_sheet(file_path):
    """Read the first sheet with a name starting with 'isa_'."""
    xls = pd.ExcelFile(file_path)
    for sheet_name in xls.sheet_names:
        if any(word in sheet_name.lower() for word in ['investigation', 'study', 'assay']):
            df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
            return df
    
    print("No sheets with 'investigation', 'study', or 'assay' found in the Excel file.")
    exit
        


def is_uppercase_with_spaces(s):
    return all(c.isupper() or c.isspace() for c in s)

def identify_sections_from_excel(excel):
    # Read the Excel file without usg the first row as headers
    df = excel
    # initialize variables
    start_row = None
    last_heading = None
    headings = []
    print('hi')
    print(df)
    # Identify the headings and their start rows
    for index, value in enumerate(df.iloc[:, 0]):
        cell_value = str(value)
        print(index,value)
        if is_uppercase_with_spaces(cell_value) and cell_value.strip():
            if start_row is not None:
                headings.append((last_heading, start_row, index))
                #print(last_heading, start_row, index)
            last_heading = cell_value
            start_row = index
            print(f"Found heading: {cell_value} at row {index}")
    if last_heading is not None:
        headings.append((last_heading, start_row, len(df)))  # Add the last heading

    return headings

# Define the revised function to process sections and save them to CSV files
def process_sections_to_csv(headings, target_hierarchy_level, df):
    # Process each section and save as CSV
    print("headings")
    print(headings[1])
    for heading, start, end in headings:
        # Iterate over each column startg from column 1 (sce column 0 is the heading)
        for col_index in range(1, df.shape[1]):  # df.shape[1] gives the number of columns
            # Check if all the values  the column for the current section are not NaN (not all none)
            if not df.iloc[start+1:end, col_index].isnull().all():
                # Select the data for the current section and column, transpose it, and set the first row as header
                data_to_save = df.iloc[start:end, [0, col_index]].transpose()
                data_to_save.columns = data_to_save.iloc[0]  # Set the first row as the header
                data_to_save = data_to_save.iloc[1:]  # Drop the first row sce it's now the header

                # Defe the CSV file name
                filename = f"{target_hierarchy_level}_{heading.strip().replace(' ', '_')}_{col_index}.csv"
                # Construct the full path to save the file

                
                # Save to CSV with the specified format and semicolon as separator
                data_to_save.to_csv(filename, sep=';', index=False, encoding='utf-8')
                print(f"Saved file: {filename}")

# Use the previously identified sections from the Excel file to generate CSV files
if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python script.py <file_path> <target_hierarchy_level>")
    else:
        file_path = sys.argv[1]
        target_hierarchy_level = sys.argv[2]
        # Read the first sheet with a name startg with 'isa_'
        print('hi')
        excel = read_first_isa_sheet(file_path)
        headings=identify_sections_from_excel(excel)
        process_sections_to_csv(headings,target_hierarchy_level,excel)


