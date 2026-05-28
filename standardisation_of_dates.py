#import pandas package
import pandas as pd



#import excel spreadsheet and assign it a variable called metadata_extract as a data frame (table)
metadata_extract=pd.read_excel("metadata_extract_20260127.xlsx")



#display head of metadata_extract on the screen to show a few lines of the table
print(metadata_extract.head())



#specify which columns have the dates I want to change. Define function called standardise date that will attempt to run the code below which is the conversion, if it doesn't work, then it will run the "except" code which will bring back the orignial value unchanged.
def standardise_date(value):
    try:
        metadata_extract["event_dates_start"]=pd.to_datetime(metadata_extract["event_dates_start"].astype(str).str.strip(),dayfirst=True,format="mixed").dt.strftime("%d-%m-%Y") 
        metadata_extract["event_dates_end"]=pd.to_datetime(metadata_extract["event_dates_end"].astype(str).str.strip(),dayfirst=True,format="mixed").dt.strftime("%d-%m-%Y") 
    except:
        return value
    
    

#apply the function to event_dates_start column 
metadata_extract["event_dates_start"] = metadata_extract["event_dates_start"].apply(standardise_date)
metadata_extract["event_dates_end"] = metadata_extract["event_dates_end"].apply(standardise_date)



#export it back into a new excel version
print("Saving new file with standardised dates")
metadata_extract.to_excel("metadata_extract_20260127_standardised_dates1.xlsx,index="False")

