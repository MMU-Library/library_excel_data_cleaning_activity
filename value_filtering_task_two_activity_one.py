#import pandas as package
import pandas as pd


#---------------import file and assign it to variable as a data frame--------------------
metadata_extract = pd.read_excel(r"C:\Users\46071956\Downloads\library_data_cleansing_activity_1\Inputs\metadata_extract_20260127.xlsx")




#--------------define what rows to keep ---------------------------------------------

allowed_types= ("article","conference_item")



#---------apply filter to create new data frame with only the rows that contain the accepted values------------------------

filtered_data=metadata_extract[metadata_extract["eprints_type"].isin(allowed_types)]




#----------print first 5 lines of the filtered data frame to check it has worked and print total number of rows in filtered data frame--------------------

print(filtered_data.head)
print(len(filtered_data))



#-----------------export the filtered data frame to a new excel file----------------------

print("saving new file")

filtered_data.to_excel(r"C:\Users\46071956\Downloads\library_data_cleansing_activity_1\Outputs\metadata_extract_20260127_filtered_task_two_activity_one.xlsx", index=False, engine="openpyxl")

print("complete")

