# **Project Title - Open Research Excel Data Cleansing script**
### The scripts in this repository are to be used to perform basic Excel data cleansing tasks, particularly, filtering by variable and standardising values.

<img width="368" height="192" alt="image" src="https://github.com/user-attachments/assets/bd62e064-c785-43a1-892d-7734bbc44303" />
								
<img width="368" height="192" alt="image" src="https://github.com/user-attachments/assets/e09e12d8-ade8-4dd8-9406-7b377e7ca225" />

-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

##*Description: This repository contains scripts that have been created to aid in the cleansing and standardisation of data from the MMU research repository.*
##*The tasks were as follows:*

1/ Standardise event dates
Parameters:
File:  metadata_extract_20260127.xlsx
Filter column D eprints_type with values conference_item AND exhibition AND performance
Go to columns BN event_dates_start and BO event_dates_end – analyse value variants and propose a methodology for standardisation of dates. Values must be in date format, i.e. DD-MM-YYYY OR YYYY-MM-DD OR other equivalent of your choice.

2/  Standardise publisher names for articles and conference papers
Parameters:
File:  metadata_extract_20260127.xlsx
Filter column D eprints_type with values article AND conference_item
Go to column AS publisher – analyse value variants and propose a methodology for name standardisation/publisher names authorities

3/ Author names
Parameters: 
File:   authors_20260127_WorkingFile.xlsx
Analyse columns B and C – first_name and last_name and propose a methodology for name standardisation. Column D id can be used to validate duplicate entries. This is the university ID of the staff member, so it could be used to check if (e.g.) A. Marsden and Andrew Marsden is the same person. However Column D id is not always populated, so caution is advised where it is empty.
Alternatively make a ‘review index’, marking author entries (and their related records – Column A resource_id) which should be reviewed to validate if the authors are the same person or not. 

-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

System Requirements/Software versions: 
Python
Pandas




### **clean_eprints.py:**
-filters metadata to keep "conference_item", "exhibition", and "performance" records.
main steps involved:
-Read Excel -> filter "e_prints_type" using the Pandas dataframe isin() method -> export to Outputs -> New Excel file created. 



### **standardisation_of_dates.py**
-standardises "event_dates_start" and "event_dates_end" to DD-MM-YYYY. 
main steps involved: 
-Read filtered Excel -> apply standardise_date using pd.to_datetime() -> export to Outputs -> New Excel File created. 



### **activity_2_metadata_extract.py:** 
-filters metadata to keep only "article" and "conference_item" records. 
main steps involved:
-Read Excel -> filter "eprints_type" with isin() -> export to Outputs -> New Excel file created. 







