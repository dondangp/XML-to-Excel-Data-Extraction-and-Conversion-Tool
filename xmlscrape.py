import os
import xml.etree.ElementTree as ET
import pandas as pd
import glob
import xlsxwriter

# Set the directory path
directory_path = "C:/Users/ddang/Desktop/xml/"

# Get a list of all XML files in the directory
xml_files = glob.glob(directory_path + "/*.xml")

# Create a list to store the data
data = []

# Iterate over each XML file
for xml_file in xml_files:
    # Parse the XML data file
    tree = ET.parse(xml_file)
    root = tree.getroot()

    # Retrieve the file name from the XML
    fileName = root.find('.//name').text
    print("Name of File:", fileName)

    # Iterate over each site and extract the information
    for site in root.findall(".//site"):
        name = site.find("name").text
        username_element = site.find(".//entry[@key='username']")
        downloadpath_element = site.find(".//entry[@key='download.folder']")
        host_element = site.find(".//entry[@key='host']")

        if username_element is not None:
            username = username_element.text
        else:
            username = "Not available"

        if downloadpath_element is not None:
            downloadpath = downloadpath_element.text
        else:
            downloadpath = "Not available"

        if host_element is not None:
            host = host_element.text
        else:
            host = "Not available"

        # Append the data to the list
        data.append({
            'Name': name,
            'Download Path': downloadpath,
            'Host': host,
            'Username': username
        })

# Create a DataFrame from the data
df = pd.DataFrame(data)

# Specify the output file path
output_file = os.path.join(directory_path, "Axway_Accounts.xlsx")

# Create a Pandas Excel writer using xlsxwriter as the engine
writer = pd.ExcelWriter(output_file, engine='xlsxwriter')

# Write the DataFrame to the Excel file
df.to_excel(writer, index=False, sheet_name='Sheet1')

# Get the xlsxwriter workbook and worksheet objects
workbook = writer.book
worksheet = writer.sheets['Sheet1']

# Iterate over the columns and adjust the widths based on content
for i, column in enumerate(df.columns):
    # Find the maximum length of the cells in the column
    column_width = max(df[column].astype(str).map(len).max(), len(column)) + 2
    # Set the width of the column
    worksheet.set_column(i, i, column_width)

# Save and close the workbook
writer.close()

print("Data exported to", output_file)
