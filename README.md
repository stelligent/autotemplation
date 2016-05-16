# autotemplation
## Description
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
TODO

#### System Configuration
1. Install Python3 ([Download] (https://www.python.org/downloads/))
2. Install required python libraries (from autotemplation directory):   
`pip3 -r requirements.txt`

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
| DATE_YEAR | 2016 |

## Running autotemplation
From the autotemplation directory, run either of the following:

`./autotemplation.py`

or

`python3 autotemplation.py`
