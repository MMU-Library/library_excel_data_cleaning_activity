#---------------import pandas package--------------------
import os
import pandas as pd
import re 
from dateutil.parser import parse
from datetime import datetime



#---------------import file and assign it to variable as a data frame--------------------
df_filtered=metadata_extract=pd.read_excel(r"C:\Users\46071956\Downloads\library_data_cleansing_activity_1\Outputs\metadata_extract_20260127_filtered_task_one_activity_one.xlsx")



    

#-------Define the date parser with logging of errors. This is a more complex function that will attempt to parse the date, and if it fails, it will log the error in an error log---------
def parse_date_to_ddmmyyyy(value, row_id_col_name, error_log):
    """Convert a date string to dd-mm-yyyy format. If conversion fails, return None and log the error."""
    if pd.isna(value) or value == "":
        return None
    val = str(value).strip()

#-----------primary parser------------------
    try:
        dt = pd.to_datetime(val, dayfirst=True, errors="raise")
#check if the year is within a reasonable range (e.g. 1900-2025) to catch outliers
        if dt.year < 1900 or dt.year > 2025:
            error_log.append(
                f"year out of range: row={row_id_col_name}, value='{val}'"
            )
            return None
        return dt.strftime("%d-%m-%Y")
    except Exception:
        pass

    #--------------- secondary parser ------------------
    #remove suffixes from day numbers (e.g. 1st, 2nd, 3rd, 4th) and try parsing again.
    try:
        clean = re.sub(r'(\d{1,2})(st|nd|rd|th)', r'\1', val)
        dt = parse(clean, dayfirst=True)
        if dt.year < 1900 or dt.year > 2025:
            error_log.append(f"year out of range: row={row_id_col_name}, value='{val}'")
            return None
        return dt.strftime("%d-%m-%Y")
    except Exception:
        error_log.append(f"unparseable date: row={row_id_col_name}, value='{val}'")
        return None


#-----------apply the function to the event_dates_start and event_dates_end columns, and captur errors in the error log during the parsing process----------

error_log = []

for col in ["event_dates_start", "event_dates_end"]:
   df_filtered[col] = df_filtered.apply(lambda row: parse_date_to_ddmmyyyy(row[col], row_id_col_name="id", error_log=error_log), axis=1)    
 

 #----------print the error log to manually review and problem rows-------------------

log_file = "dates_standardsation_error_log.txt"
with open (log_file,"w") as log:
   log.write("Date parsing errors (see below)./n")
   for entry in error_log:
      log.write(entry + "\n")
      if not error_log:
         log.write("no errors found./n")
print("Date parsing complete. Errors logged to {log_file}")


#------------save new file with standardised dates-------------------

output_file = r"C:\Users\46071956\Downloads\library_data_cleansing_activity_1\Outputs\metadata_extract_20260127_standardised_dates_activity_one_task_two.xlsx"
df_filtered.to_excel(output_file, index=False, engine="openpyxl")
print(f"Standardised dates saved to {output_file}")   