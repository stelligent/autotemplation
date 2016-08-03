# autotemplation
## Description
Tired of making custom copies of the same document in Google Drive over and over? With autotemplation, you can add variables to your Google Docs and fill them out with ease!

Gathers variables from a selected file in Google Drive and prompts the user to fill each variable. The completed file is uploaded to a folder in user's Google Drive.
## Setup
The setup requires autotemplation configuration, credentials download, and system setup.
#### autotemplation Configuration
1. Download autotemplation and create ini file:   

  ```
  git clone https://github.com/stelligent/autotemplation.git
  cd autotemplation
  cp autotemplation.ini.sample autotemplation.ini
  ```

2. Edit autotemplation.ini.

  * Add id of folder where templates will live. (TemplateFolderID)   
  (Hint: Browse to desired folder in Google Drive and find the id in the url:  
  https://drive.google.com/drive/folders/XXXXXXXXXXXXXXXXXX  
  copy the X's into autotemplation.ini)
  
  * Optionally, change the DestinationFolderName to whatever you'd like.  
  This folder will be created at the base level of your Drive, if not found.

#### Credentials Download
1. Browse to [Google API Console]. (https://console.developers.google.com/apis/library)
2. Click Credentials on the left side.
3. Create Credentials.
4. Create a project. Use autotemplation for name throughout. 
5. Click Create credentials and choose OAuth Client ID.
6. Click Configure Consent Screen.
7. Enter a Product Name and Save.
8. Click other, enter a name, and click Create.
9. Hit OK on key screen.
10. Click the download button next to your key and save as `client_secret.json` in the autotemplation directory. (This file is ignored and should not be commited back to the repository.
11. Click Overview.
12. Click Drive API under Google Apps APIs.
13. Click Enable. 

#### System Configuration
1. Install Python3 ([Download] (https://www.python.org/downloads/))
2. Install required python libraries (from autotemplation directory):   
`pip3 install -r requirements.txt`

## Template Setup
Variables are set up with the following format: `{{ variable_name }} `   
The space between the curly braces and the variable name are required. Do not use spaces within the variable name.

Variables can be placed in the name of the file as well as anywhere within the file.  
The program will ask for a date to use for the document date. It defaults to the current date and creates the following variables automatically:   

Example Date: 20160101

| Variable Name | Date Format |
| --- | --- |
| DATE_FULL | January 1, 2016 |
| DATE_FULL_NUM | 20160101 |
| DATE_FULL_DASH | 01-01-2016 |
| DATE_FULL_SLASH | 01/01/2016 |
| DATE_MONTH | January |
| DATE_MONTH_NUM | 01 |
| DATE_DAY_FULL | Friday |
| DATE_DAY_SHORT | Fri |
| DATE_DAY_NUM | 01 |
| DATE_DAY_SUFFIX | st |
| DATE_YEAR | 2016 |

#### Using Google Sheets as Input
A double underscore '__' is reserved for Google Sheet lookups. 
When detected in the template, autotemplation will prompt the user for the Google Sheet ID to use.

The Google Sheet lookup is done by column name and row name (for now). Replace spaces with underscores in the template.
Underscores in the row or column headers are not supported at this time.

For Example:

| Example | num one | num two | num three |
| --- | --- | --- | --- |
| ten | 10 | 20 | 30|
| hundred | 100 | 200 | 300|
| thousand | 1,000 | 2,000 | 3,000|

{{ num_two__thousand }} in the template would be replaced with 2,000.

## Running autotemplation
From the autotemplation directory, run either of the following:

`./autotemplation.py`

or

`python3 autotemplation.py`

To run without a terminal (in OSX), ctrl-click or right click the autotemplation.py file in finder and open with IDLE. Press F5 to run.
