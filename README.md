# XML-Demo

This XML Demo was created to demonstrate a simple Python Script that scrapes data from an XML file, exports it to Excel, and executes a macro to set up a Pivot Table.

test_xml.xml = A dummy XML file that shows relevant information for different books.

xml_parser.py: This is the Python script. It reads from the XML file "test_xml.xml", and creates a readable dataframe using the pandas library. The formatted raw data table is exported to Excel, and the vba script is executed.

pivot_table_vba_script.vbs: this is a vbs script that can be executed through Python and runs the macro "pivotmacro.xlsm"

pivotmacro.xlsm: this macro creates a Pivot Table that shows the average price of a book in each genre from the example XML raw data.

cleaned_xml.xlsx = This is the resulting Excel file that is created after running the Python script. 

