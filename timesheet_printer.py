# timesheet_printer.py
import requests
from bs4 import BeautifulSoup
import re
import os
from urllib.parse import urljoin
import win32print
import win32api
import time

def print_pdf(pdf_path):
    try:
        printer_name = win32print.GetDefaultPrinter()
        win32api.ShellExecute(0, "print", pdf_path, f'/d:"{printer_name}"', ".", 0)
        time.sleep(2)
        return True
    except Exception as e:
        print(f"Printing error for {pdf_path}: {e}")
        return False

def print_timesheets(base_url, page_url):
    if not os.path.exists('downloads'):
        os.makedirs('downloads')
    
    session = requests.Session()
    
    try:
        response = session.get(page_url)
        response.raise_for_status()
        
        soup = BeautifulSoup(response.text, 'html.parser')
        print_links = soup.find_all('a', href=lambda href: href and 'printTimeSheet' in href)
        
        for link in print_links:
            try:
                onclick = link.get('onclick', '')
                url_match = re.search(r"window\.open\('([^']+)'", onclick)
                
                if url_match:
                    pdf_url = urljoin(base_url, url_match.group(1))
                    username = re.search(r'username=([^&]+)', pdf_url)
                    username = username.group(1) if username else 'unknown'
                    
                    pdf_response = session.get(pdf_url)
                    pdf_response.raise_for_status()
                    
                    filename = f'downloads/timesheet_{username}.pdf'
                    with open(filename, 'wb') as f:
                        f.write(pdf_response.content)
                    print(f"Downloaded: {filename}")
                    
                    if print_pdf(filename):
                        print(f"Sent to printer: {filename}")
                    
            except Exception as e:
                print(f"Error processing link: {e}")
                
    except requests.exceptions.RequestException as e:
        print(f"Error accessing page: {e}")
    
    finally:
        session.close()

if __name__ == "__main__":
    base_url = "https://staffservices.ics.forth.gr/TimeSheets/"
    page_url = f"{base_url}TeamHistory.php?search=1&month_from=12&month_to=12&date_to=&year=2024&page=1"
    print_timesheets(base_url, page_url)
