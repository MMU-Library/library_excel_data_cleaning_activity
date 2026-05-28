#import pandas as package
import pandas as pd


#import the metadata from excel and assign it to the variable called metadata_extract as a data frame(table)
metadata_extract = pd.read_excel('metadata_extract_20260127.xlsx')



#display head of metadata_extract to see the first few lines
print("here are the first few lines of the dataframe")

print(metadata_extract.head())
print(metadata_extract.columns)



#define which rows I would like to keep: article AND conference_item
allowed_types= ("article","conference_item")



#create a filter to show me only the rows that contain the values "article" and "conference_item" within column D(eprints_type)
filtered_data=[metadata_extract["eprints_type"].isin(allowed_types)]



#print out the filtered result
print("below is the filtered data")
print(filtered_data)



filtered_data=metadata_extract[metadata_extract["eprints_type"].isin(allowed_types)]
print(filtered_data.head)
print(len(filtered_data))



#export it back into an excel spreadsheet
filtered_data.to_excel:("metadata_extract_20260127_filtered_activity_two_cleaned.xlsx")


filtered_data.to_excel(r"C:\Users\46071956\Downloads\library_data_cleansing_activity_1\metadata_extract_20260127_filtered_activity_two.xlsx", index=False, engine="openpyxl")

print("complete")
