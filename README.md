# ProcessGarmin
Process data from Garmin Approach S3 to excel

This program processes data from the Garmin Approach S3 (or other Garmin watches?)
Only the Approach S3 has been tested

Attach Garmin to computer through USB

You have two options to select:
  Year (YYYY) to process (leaving blank will select all years)
  Course - substring of course (leaving blank will select all courses)
  
Press "Extract to Excel" button

You will be presented two directory selection boxes:

  1. Select the input directory:
      MAC OS - /Volumes/Garmin/Garmin/Data/Scorecards
      Windows - Garmin Approach S3 (x:) /Garmin/Data/Scorecards
      
  2. Select the output directory where spreadsheet will go
      The file name will be course(substring).xls or GarminData.xls (if all courses) 
