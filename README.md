## MXToolbox Email Checker 
## Introduction 

This script is designed to check the MX (Mail Exchange) records for a list of domains using the MXToolbox API. It fetches the MX records for each domain and saves the results to an Excel file. Additionally, it provides insights such as whether DNS records are found, the mail server details, and the email service provider for each domain.

## Prerequisites 
Python 3 requests library (install via pip install requests) pandas library (install via pip install pandas) openpyxl library (install via pip install openpyxl) colorama library (install via pip install colorama) 

## Usage Clone the repository: bash Copy code 
git clone https://github.com/your_username/mxtoolbox-email-checker.git 

## Navigate to the project directory: bash Copy code cd mxtoolbox-email-checker 

Install the required dependencies: bash Copy code pip install -r requirements.txt Place your domain names in an Excel file named domains.xlsx in the root directory. Ensure that the Excel file contains a column named Domain with the list of domains.

Replace "Your-API-Key" in the script with your actual MXToolbox API key.

## Run the script:

bash Copy code python mxtoolbox_email_checker.py The script will process each domain, check its MX records, and save the results to an Excel file named results_.xlsx in the results folder. Configuration domains.xlsx: Excel file containing the list of domain names to be checked. results: Folder where the output Excel file will be saved. Your-API-Key: Replace with your actual MXToolbox API key. License This project is licensed under the MIT License - see the LICENSE file for details.