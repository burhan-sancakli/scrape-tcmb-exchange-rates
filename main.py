import os
import requests
import xml.etree.ElementTree as ET
import pandas as pd
from datetime import datetime, timedelta
from drive_utility import upload_photo

# Define start and end dates
start_date = datetime(2022, 12, 30)
end_date = datetime.today()

# Path to the Excel file
output_file_path = r'./tcmb_currency_data.xlsx'

# Check if the Excel file already exists
if os.path.exists(output_file_path):
    # Load the existing Excel file
    existing_data = pd.read_excel(output_file_path)

    # Get the last date in the existing data
    last_date_str = existing_data['Date'].max()
    last_date = datetime.strptime(last_date_str, '%Y-%m-%d')

    # Start scraping from the day after the last recorded date
    start_date = last_date + timedelta(days=1)
    print(f"Resuming from {start_date.strftime('%Y-%m-%d')}")
else:
    # If the file doesn't exist, initialize an empty DataFrame
    existing_data = pd.DataFrame()

# List to hold all the data (including previously saved data)
all_data = existing_data.to_dict('records') if not existing_data.empty else []

# Variable to hold the last available data (for missing days)
last_available_data = all_data[-1:] if all_data else None

# Iterate through each date
current_date = start_date
while current_date <= end_date:
    year = current_date.strftime('%Y')
    month = current_date.strftime('%m')
    day = current_date.strftime('%d')

    # Build the URL for the current date
    url = f'https://www.tcmb.gov.tr/kurlar/{year}{month}/{day}{month}{year}.xml'
    print(url)

    # Send a request and process the XML if the request is successful
    response = requests.get(url)
    if response.status_code == 200:
        xml_content = response.content
        root = ET.fromstring(xml_content)

        # Temporary list to hold data for the current day
        current_day_data = []

        # Iterate through each currency and extract relevant data, excluding XDR
        for currency in root.findall('Currency'):
            code = currency.get('Kod')
            if code == 'XDR':
                continue  # Skip XDR

            name = currency.find('Isim').text
            forex_buying = currency.find('ForexBuying').text
            forex_selling = currency.find('ForexSelling').text
            banknote_buying = currency.find('BanknoteBuying').text
            banknote_selling = currency.find('BanknoteSelling').text
            if forex_buying: forex_buying.replace(".",",")
            if forex_selling: forex_selling.replace(".",",")
            if banknote_buying: banknote_buying.replace(".",",")
            if banknote_selling: banknote_selling.replace(".",",")
            # Append the data for the current day
            current_day_data.append({
                'Date': current_date.strftime('%Y-%m-%d'),
                'Currency Code': code,
                'Currency Name': name,
                'Forex Buying': forex_buying,
                'Forex Selling': forex_selling,
                'Banknote Buying': banknote_buying,
                'Banknote Selling': banknote_selling
            })

        # Update last available data
        last_available_data = current_day_data
        all_data.extend(current_day_data)

    else:
        print("No data for the day")
        # If data is not available for the current day, copy the last available data
        if last_available_data:
            for entry in last_available_data:
                # Make sure to update the date to the current date
                new_entry = entry.copy()
                new_entry['Date'] = current_date.strftime('%Y-%m-%d')
                all_data.append(new_entry)

    # Create a DataFrame from all data (old and new)
    df = pd.DataFrame(all_data)

    # Save the entire data to an Excel file after each day's scrape
    df.to_excel(output_file_path, index=False)
    print(
        f"Data for {current_date.strftime('%Y-%m-%d')} saved to {output_file_path}")

    # Move to the next day
    current_date += timedelta(days=1)


print("Data scraping completed.")
upload_photo("tcmb_currency_data.xlsx")
print("Data has been uploaded to Google Drive.")
