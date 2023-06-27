# XML-to-Excel-Data-Extraction-and-Conversion-Tool
This program automates the process of extracting data from multiple XML files and organizing it into an Excel file. It iterates through a specified directory, 
locating XML files within it. For each XML file, it parses the data and retrieves relevant information such as site name, username, download path, and host. 
The extracted data is then stored in a DataFrame. The program dynamically adjusts the column widths in the Excel file to accommodate the data. 
Finally, it saves the DataFrame to an Excel file and provides a message confirming the successful export of the data.
Overall, this program streamlines the conversion of XML data into a structured Excel format, making it easier to analyze and work with the data.

Abilities:
It sets the directory path where the XML files are located.
It retrieves a list of all XML files in the specified directory.
It creates an empty list to store the extracted data.
It iterates over each XML file in the list.
For each XML file, it parses the XML data and retrieves the file name.
It iterates over each "site" element in the XML and extracts relevant information such as name, username, download path, and host.
It appends the extracted data to the list.
Once all XML files have been processed, it creates a DataFrame from the collected data.
It specifies the output file path for the Excel file.
It creates a Pandas Excel writer using the xlsxwriter engine.
It writes the DataFrame to the Excel file.
It retrieves the workbook and worksheet objects from xlsxwriter.
It iterates over the columns of the DataFrame and adjusts the column widths based on the content.
It saves and closes the workbook.
Finally, it prints a message indicating the successful export of the data to the Excel file.
