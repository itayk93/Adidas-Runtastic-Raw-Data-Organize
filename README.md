Hey! This code purpose is to organise the input json files which contains runnings data from Adidas Runtastic App.
It organises the data into an input file to upload to any BI system (in my cace - Tableau)

How to run the software:

1. Export the raw data from Runtastic website - you have to submit a request in the website and the raw data will get in the mail within few days
2. After you get the mail from Runtastic with the data download it and locate it into "Exported Input Data" folder which is in "Input" folder
3. Define the "data_to_cancel.xlsx" file in case if there's any specific data you want to delete from the final data before uploading to a BI system (just copy the data from the output after running the software)
4. Define the "dates_to_cancel.xlsx" file in case if there are dates you want to delete from the final data before uploading to a BI system
5. Define the "main_folder_location.xlsx" file with your folder location
6. Define the "min_average_speed.xlsx" with the minimum average speed you want to show on the final data (the purpose of this is to clean a lot fo data that can effect the final results, like runnings which has mistakenly taken)
7. And finally you can run the "Organize Data - Runtastic App.py" script

Thank you

