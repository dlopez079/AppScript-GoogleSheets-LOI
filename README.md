# AppScript-GoogleSheets-LOI
I created a Letter of Intent generator within Google Sheets that will use App Script to generate the Google Docs Letter of intent, using the template that you create. 

## How to use Step By Step

### Step 1: Create and Save Generator File 
Create a Google Sheets file and save it in your Google Drive folder (My Drive).   The file must be saved within your Google drive My drive folder.

### Step 2: App Script
When you are within the Google Sheets file, you will see a tab on the top that says extension.  Click on that and you will see App Script.  Inside of App Script, you will begin with a file already created called Code.gs.  

### Step 3: Copy and Paste
You will see that I have 4 files in my github called, Code.gs, OfferCalc.gs, EMDCalc.gs, and createNewGoogleDocs.gs.  Each file has it's own set of instructions.  

Code.gs will create a new tab in your google sheets called Mailer.  You can change the name in the code if you want to use something else.  

OfferCalc.gs will get a percentage of the asking price that you are going to offer the receipient.  

  Asking Price (100,000) * 60% = Offer Price ($60,000).  I obtained the percentage from a table I created. 
  List Price Range	Percent
  <$150k	60%
  $150-200k	65%
  $200-300k	70%
  $300-400k	75%
  $400-500k	80%
  >
  ![image](https://github.com/dlopez079/AppScript-GoogleSheets-LOI/assets/50810168/d5929642-503e-4ca0-93a3-b68187a02bd8)
  >

EMDCalc.gs will create the Ernest Money Deposit based on 2% of the offer price.

createNewGoogleDocs.gs will create the Google Docs for each row based on the Google Docs Template you created for the Letter of Intent.

### Step 4: Obtain Your Receipient List.
Obtain your Excel Spreadsheet from your third party vendor.  Most of you in the real estate world will obtain a list from a third party source like propstream, property radar, etc.  Sometimes, the list will download as a CSV.  Change it to an excel document and make sure you have the followwing columns (I only used these columns to create a letter of intent.  Some of you may have more information that you would like to add.  If you do add columns, YOU WILL NEED TO MAKE ADJUSTMENTS TO THE CODE.)

The columns that I used are | First Name	| Last Name	| Address	| City 	| State	| Zip	| APN	| Asking Price.

Once you cleaned up your excel spreadsheet and you only have these columns, you will revise the name of the current worksheet that you are in and name it data (important-will come up later).  You then save it onto the same Google Drive (My Drive) as your generator and call it lists.xlsx (keep it a simple name).  You now created a total of two files.  The Google Sheets Generator and the Excel Spreadsheet List

### Step 5: Create your Google Sheets Letter of Intent Template
Create a letter of intent and add the placeholders that you need for the each of the columns you created in the files before.  The placeholders should like as follows: 

{{First Name}}, {{Last Name}}, {{Address}}, {{City}}, {{State}}, {{Zip}}, {{APN}}, {{Asking Price}}

### Step 6: Mailer Tab in Mailer Generator
You should not have a tab in your Google Generator Sheets that says Mailer (Step 3).  Once you click on Mailer, you should have 4 options.  

Option 1: Import Excel File from Drive
You will get a prompt for the file that you would like to import.  If you saved your file as list.xlsx, then you will put "list.xlsx" into the file.  It will need the extention to work.  
You will need get another prompt for the workshee that you want to import (remember I said it will come up later.  lol).  You will type in "data" to obtain the data within that worksheet.  
The generator will create a temporary sheet in the back ground then migrate the information into the generator.  
The generator will create a new sheet that will be time stamped along with date. 

Option 2: Get Offer
Mailer > Get Offer
Once you click on Get Offer, the generator will take the asking price, which should be in the last coloumn, create a new column, input the title of the offer column right after Asking Price Column, and generate the offer price based on the table above.  
The Offer Price Column should now be the last column and it should be populated with numbers.  

Option 3: Get EMD
Mailer > Get EMD
Once you click on Get EMD, the generator will take the offer price, which should be in the last column, create a new column, input the title of the EMD column right after Offer Price Column, and generate the EMD price based on 2%.  You can change this percentage in the code if you'd like.
The EMD Columnn shjould now be the last column and it should be populated with numbers. 

Option 4: Mail Merge
Mailer > Mail Merge
Once you click on Mail Merge, the generator will go through all the rows, create a new Google Docs using the template you created, add the data to the placeholders and save the new Letter of Intent to the same Google Drive (My Drive). It will repeat this for each record, or row. 
In addition, the generator will create a new column, populate the title called "Document Link", and save the Google Doc Link for reach row.  
You can use the document link to access the file or you can navigate to your Google Drive (My Drive) to access the files.  


COMMENTS: 
I created this in less than 10 minutes.  There is so much that can be improved but I needed something fast.  I welcome anyone to  make additions or recommendations.  
