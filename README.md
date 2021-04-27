# Data-migration
This python code takes an vehicle order request table in google sheet as an examplea to demonstrate a data migration from google sheet to sql server. The code also includes a
fuzzy string matching method to clean up employee name typos

# python code steps
1. Set up google sheet API connection and get data in pandas dataframe formmat. A useful [blog](https://medium.com/analytics-vidhya/how-to-read-and-write-data-to-google-spreadsheet-using-python-ebf54d51a72c)
2. Connect to sql server to a HR dataset to get the correct employee names
3. Clean up employee names
4. Move google sheet data to sql server 

# Sample file
A vehicle request sample.csv shows how a google sheet vehicle request table would be like and name typos 

# Result summary
After data migration and cleaning, name tpyos in the sample file will be corrected and this table will be moved to sql server


# Notes
In order for this python script to work, a credential.json file needs to be downloaded after google API is enabled and save in the same folder where this python script is saved
