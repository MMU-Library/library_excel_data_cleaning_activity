#import pandas package and also file directory 
import os
import pandas as pd


#import the metadata from excel and assign it to the variable called metadata_extract as a data frame(table)
metadata_extract = pd.read_excel(r"C:\Users\46071956\Downloads\library_data_cleansing_activity_1\Inputs\metadata_extract_20260127.xlsx")


#display head of metadata_extract on screen to see the first few lines of the table
print(metadata_extract.head())
print(metadata_extract.columns)


#define what rows I want to keep  ("conference_item", "exhibition", "performance")
allowed_types= ("conference_item", "exhibition", "performance")


#create a filter to show me only rows that contain values conference_item AND exhibition AND performance within Column D ()
filtered_data=[metadata_extract["eprints_type"].isin(allowed_types)]


# get it to show me the filtered result 
print("below is the filtered data")

print(filtered_data)

#this only shows me whether rows meet the conditions true/false, i want it to FILTER it down. 
print("below is the final output")
filtered_data=metadata_extract[metadata_extract["eprints_type"].isin(allowed_types)]



#show the first few lines of rows that meet conditions and also show the number of total lines.
print(filtered_data.head())
print(len(filtered_data))


#export it back into an excel spreadsheet (note make sure to save before running as it was not recognising the new file path that I had pasted)
print("Saving new file with filtered data")

filtered_data.to_excel(r"C:\Users\46071956\Downloads\library_data_cleansing_activity_1\Outputs\metadata_extract_20260127_filtered.xlsx", index=False)