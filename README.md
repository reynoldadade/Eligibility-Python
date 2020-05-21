# Eligibilty Data Complete Code Documentation

> To be able to run this script you need to have some python libraries installed and they are all listed in the included `requirements.txt` file in this project

To install ensure you have **python** and **pip** installed  on your local machine

Run `pip install -r requirements.txt`

> Please have the most current version of python installed


## Preparing the excel file
>Before running the script ensure that your excel file is  in this format and  this format only

- EMPLOYEE_NUMBER
- FULL_NAME	DATE_OF_BIRTH
- ASSIGNMENT_STATUS	
- GENDER	
- JOB	
- ORGANIZATION
- HIRE_DATE
- DEPARTMENT
- MINISTRY
- SOCIAL_SECURITY_NUMBER
- BANK_NAME
- BANK_BRANCH_NAME
- LOCATION
- DISTRICT
- REGION

 >To make sure that you are formatting is easier for you i am including the sequence of steps i  use to ensure the excel does not have misplaced values anywhere else

 1. Open the excel
 2. go to **Data** 
    - **Remove Duplicates** 
    - **Unselect All** 
    - select **EMPLOYEE_NUMBER** 
    - Click on **OK**
 3. When it is done removing Duplicates 
    - Go to **Home** click on **Sort & Filter** 
    - **Custom Sort** 
    - Sort By **HIRE_DATE** 
    - click **OK**
    - There will be some HIRE_DATE columns empty because the data has shiffted to another column
    - A simple delete in the empty column will enable you to shift the columns to the left
 4. When done with the HIRE_DATE steps proceed to do a **Custom Sort**  and sort by BANK_NAME
    - There will be empty columns too after the sort due to shifted data columns deleted the empty rows should also correct and set the data right
 5. Do a final sort by **EMPLOYEE_NUMBER** and there might be some columns that are empty please delete all columns that have no employee numbers as they cannot be used. 
 6. Copy and move **ASSIGNMENT_STATUS** to the 3rd Column as it will be the last column in the excel
 7. Remove all other columns other than the one listen above.
 8. Fill in all other empty columns with _**NONE**_ using **find and replace**. Make sure you match with **entire cell contents**


## Running the script

> **Always run this script in the same folder as the excel file you want to work on**

- Run **cmd.exe** on your windows machine or **terminal** on linux or mac
- `cd` to the folder that the excel file and the script is in
- input `python complete_code.py` 
- code will request you enter name of excel file
> Please dont enter extension with the file name
- Code should start to run and create **two** files for you on completion
- The two files will have the names 
   - '**name of file**_sql_upload.xlsx'
   - '**name of file**'_CUT_UPPERED_FORMATTED_DATED_DONE.xlsx'

## Uploading to SQL Server   
'**name of file**_sql_upload.xlsx' will be uploaded directly to the sql you need to append **CR** to the EMPLOYEE_NUMBER column 
   >NB Before uploading a few house cleaning
   
   - Ensure that **EMPLOYEE_NUMBER** with the CR appended is saved a text value and not as formula.
   Name of new columns formed from the script in this order and this order only
      - ssn_exempt	
      - pupil_teacher	
      - trainee_teacher

   - Upload to the sql server using management studio or any other form of sql management software you have 
   - Destination Table is **Eligibility Data**
   - Final excel should have all these columns
      - EMPLOYEE_NUMBER	FULL_NAME	
      - DATE_OF_BIRTH	
      - ASSIGNMENT_STATUS	
      - GENDER	
      - JOB	
      - ORGANIZATION	
      - HIRE_DATE	
      - DEPARTMENT	
      - MINISTRY	
      - SOCIAL_SECURITY_NUMBER	
      - BANK_NAME	
      - BANK_BRANCH_NAME	
      - LOCATION	
      - DISTRICT	
      - REGION	
      - ssn_exempt	
      - pupil_teacher	
      - trainee_teacher

##  Uploading to the linux box      

- '**name of file**'_CUT_UPPERED_FORMATTED_DATED_DONE.xlsx' is to be uploaded to the linux box

> Some house cleaning
   - A **VLOOKUP** should be done with the excel file with the name **TEL_NUMBERS.xlsx** which is also included in the repo, you will match with **EMPLOYEE_NUMBER** to Telephone numbers
   - The **vlookup** will be done twice to match two numbers for each client, not all clients have numbers in the TEL_NUMBERS excel file so the results will show as **#N/A**, ensure that results are aslo  saved as text values and not formulas thesese values are added to the **tel 1** and **tel 2** columns 
   - After the vlookup added another column called Blacklist and set it all to **No**

      - Final excel should have all these columns
      - EMPLOYEE_NUMBER	FULL_NAME	
      - DATE_OF_BIRTH	
      - ASSIGNMENT_STATUS	
      - GENDER	
      - JOB	
      - ORGANIZATION	
      - HIRE_DATE	
      - DEPARTMENT	
      - MINISTRY	
      - SOCIAL_SECURITY_NUMBER	
      - BANK_NAME	
      - BANK_BRANCH_NAME	
      - LOCATION	
      - DISTRICT	
      - REGION	
      - ssn_exempt	
      - pupil_teacher	
      - trainee_teacher
      - tel 1	
      - tel 2	
      - Blacklist

      > **Note**: for the linux box uploads names of the columns do not matter as upload is done based on the index of the columns therefore it is very important that columns are arranged specifically as stated above

      > **Note** blacklist is always No 
   - Save this file  as **Final_Product**

   - Connect to the linux box using Anydesk using the desk number **966510990** contact IT admin if connection is failing
   - Enter the box using the **Dalex Ghost** Account
   - Password is **D@lexGhost**
   - Go to **home > djcode > dalex**
   - Copy **Final_Product** to the directory it should already have one versio there **overewrite it**
   - Open **terminal** navigate to the **home > djcode > dalex**
   - Enter `python manage.py uploading_data`
   - Press Enter
      


      



## Further Notes

> Included in the repo will be sample data that has been worked on so that you can knowthe starting and endpoint of the excel data






