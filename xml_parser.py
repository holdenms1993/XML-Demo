
import xml.etree.ElementTree as ET
import pandas as pd
import os

xml_file = 'test_xml.xml'

#read in xml file and parse
def read_xml_file(xml_file):
    element_tree = ET.parse(xml_file)
    root = element_tree.getroot()
    return root

#create a list of dictionaries to convert to pandas dataframes
def create_xml_dict(root):
    #create dictionary keys based of attributes for each book in XML file
    dict_keys = [x.tag for x in root[0]]

    #create book ID in dictionary
    dict_list = [child.attrib for child in root]

    #iterate through each child element and append values to dictionary
    for num, child in enumerate(root):
        for key in dict_keys:
            dict_list[num][key] = child.find(key).text

    return dict_list

def clean_df(xml_dict):
    df = pd.DataFrame(xml_dict)

    #replace newline and additional whitespace to clean the description field
    df = df.replace('\n', '', regex=True)
    df = df.replace('\s+',' ', regex=True)

    #rename columns for readability
    df = df.rename(columns={'id': 'ID', 'author': 'Author', 'title': 'Title', 'genre': 'Genre', 'price': 'Price', 'publish_date': 'Publish Date', 'description': 'Description'})

    #index datafram by book ID
    df = df.set_index('ID')
    df['Price'] = pd.to_numeric(df['Price'])
    return df

df = clean_df(create_xml_dict(read_xml_file(xml_file)))
df.to_excel('cleaned_xml.xlsx')

os.system('pivot_table_vba_script.vbs')
