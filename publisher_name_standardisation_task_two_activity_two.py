#import pandas package, fuzzywuzzy for string matching, and pip to install packages
import ensurepip
from gettext import install
import fuzzywuzzy
import pandas as pd
import pip
ensurepip
import process
 



#import excel spreadsheet and assign it a variable called metadata_extract as a data frame (table)
metadata_extract = pd.read_excel(r'C:\Users\46071956\Downloads\library_excel_data_cleaning_scripts\Outputs\metadata_extract_20260127_filtered_activity_two.xlsx')


#read and store the values within variable "Publisher" as a list
publisher_list = metadata_extract["Publisher"].tolist()

#print out the publisher list to see the values
