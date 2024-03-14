import os.path
import requests
import pandas as pd
from datetime import datetime
from colorama import Fore, Style
import openpyxl  # Import openpyxl library for Excel file manipulation

# MXToolbox API endpoint
API_ENDPOINT = "https://api.mxtoolbox.com/api/v1/Lookup/"

# Function to check MX records for a domain
def check_mx_records(domain, api_key):
    params = {
        "Command": "mx",
        "argument": domain
    }
    headers = {
        "Authorization": api_key
    }
    try:
        response = requests.get(API_ENDPOINT, params=params, headers=headers)
        if response.status_code == 200:
            return response.json()
        else:
            print(f"{Fore.RED}Failed to check MX records for {domain}. Status code: {response.status_code}{Style.RESET_ALL}")
            return None
    except Exception as e:
        print(f"{Fore.RED}An error occurred while checking MX records for {domain}: {e}{Style.RESET_ALL}")
        return None

# Function to read domain names from an Excel file
def read_domains_from_excel(file_path):
    try:
        df = pd.read_excel(file_path)
        return df['Domain'].tolist()
    except Exception as e:
        print(f"{Fore.RED}Failed to read domain names from Excel file: {e}{Style.RESET_ALL}")
        return []

# Function to save results to an Excel file
def save_results_to_excel(results, output_folder):
    try:
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)
        
        output_file = os.path.join(output_folder, f"results_{datetime.now().strftime('%d-%m-%Y-%H-%M-%S')}.xlsx")
        results.to_excel(output_file, index=False)
        
        # Adjust column widths and center-align data
        wb = openpyxl.load_workbook(output_file)
        ws = wb.active
        for column_cells in ws.columns:
            length = max(len(str(cell.value)) for cell in column_cells)
            ws.column_dimensions[column_cells[0].column_letter].width = length + 2
            for cell in column_cells:
                cell.alignment = openpyxl.styles.Alignment(horizontal='center')
        wb.save(output_file)
        
        print(f"{Fore.GREEN}Results saved to {output_file}{Style.RESET_ALL}")
    except Exception as e:
        print(f"{Fore.RED}Failed to save results to Excel file: {e}{Style.RESET_ALL}")

# Main function
def main():
    excel_file_path = "domains.xlsx"  # Change this to your Excel file path
    output_folder = "results"  # Output folder name
    api_key = "c314f486-0569-4b21-87a6-5a04d8d24a75"  # Replace "Your-API-Key" with your actual API key

    domains = read_domains_from_excel(excel_file_path)
    if domains:
        results_df = pd.DataFrame(columns=["Domain", "DNS Records Found", "Mail Server 1 Hostname", "Mail Server 1 IPv4", "Email Service Provider"])

        for domain in domains:
            print(f"Processing domain: {domain}")
            result = check_mx_records(domain, api_key)
            if result:
                dns_records_found = "Yes" if result.get('Passed') else "No"
                
                mail_servers = result.get('Information', [])
                mail_server_1_hostname = mail_servers[0]['Hostname'] if mail_servers else "N/A"
                mail_server_1_ipv4 = mail_servers[0]['IP Address'] if mail_servers else "N/A"
                
                email_service_provider = result.get('EmailServiceProvider', "Unknown")
                
                results_df = results_df._append({"Domain": domain, 
                                                "DNS Records Found": dns_records_found,
                                                "Mail Server 1 Hostname": mail_server_1_hostname,
                                                "Mail Server 1 IPv4": mail_server_1_ipv4,
                                                "Email Service Provider": email_service_provider}, 
                                                ignore_index=True)

        if not results_df.empty:
            save_results_to_excel(results_df, output_folder)

if __name__ == "__main__":
    main()
