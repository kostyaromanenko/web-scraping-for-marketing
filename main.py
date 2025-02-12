import pandas as pd
import requests
from bs4 import BeautifulSoup
import re
import random
import time
import googlemaps
import os


API_KEY = 'AIzaSyDCaPq3ltSWKug8OUd9cOeGWd_6lUxCBf0'
gmaps = googlemaps.Client(key=API_KEY)

def search_company_website(company_name):

    search_url = f"https://duckduckgo.com/html/?q={company_name}"
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3"
    }

    response = requests.get(search_url, headers=headers)

    if response.status_code == 200:
        soup = BeautifulSoup(response.text, 'html.parser')
        for link in soup.find_all('a', href=True):
            href = link['href']
            if "http" in href and "duckduckgo.com" not in href:
                return href
    return None

def find_email_on_website(url, timeout=10, max_retries=3):
    for attempt in range(max_retries):
        try:
            response = requests.get(url, timeout=timeout)
            if response.status_code == 200:
                soup = BeautifulSoup(response.text, 'html.parser')
                emails = set(re.findall(r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}", soup.text))
                return emails if emails else None
            else:
                print(f"Failed to access {url}, status code: {response.status_code}")
        except requests.Timeout:
            print(f"Request to {url} timed out. Attempt {attempt + 1} of {max_retries}.")
        except requests.ConnectionError:
            print(f"Connection error when trying to access {url}. Attempt {attempt + 1} of {max_retries}.")
        except Exception as e:
            print(f"An unexpected error occurred while fetching {url}: {e}")
            break
    print(f"Skipping {url} after {max_retries} failed attempts.")
    return None

def get_address_from_google_maps(company_name):
    try:
        geocode_result = gmaps.geocode(company_name)
        if geocode_result:
            for result in geocode_result:
                address_types = result['types']
                if 'street_address' in address_types or 'locality' in address_types:
                    return result['formatted_address']
            return geocode_result[0]['formatted_address']
        else:
            return None
    except Exception:
        return None

def calculate_distance(address, destination):
    if pd.isna(address) or not isinstance(address, str) or address.strip() == "":
        return None
    try:
        result = gmaps.distance_matrix(address, destination, mode='driving')
        distance_info = result['rows'][0]['elements'][0]
        if distance_info['status'] == 'OK':
            distance = distance_info['distance']['value'] / 1000
            return int(round(distance))
        else:
            return None
    except Exception:
        return None

file_path = 'NICHE_GRAINUKRAINE_09.10.24.xlsb'
data = pd.read_excel(file_path, engine='pyxlsb')

if 'Website' not in data.columns:
    data['Website'] = ''
if 'E-mail' not in data.columns:
    data['E-mail'] = ''
if 'Address' not in data.columns:
    data['Address'] = ''
cities = ["Kyiv", "Kharkiv", "Khmelnytskyi", "Mykolaiv"]
for city in cities:
    if city not in data.columns:
        data[city] = None

def save_progress(dataframe, path):
    temp_file = path + '.xlsx'
    dataframe.to_excel(temp_file, index=False)
    os.replace(temp_file, path + '.temp')

updated_file_path = 'updated_NICHE_GRAINUKRAINE_09.10.24.xlsx'

try:
    for index, row in data.iterrows():
        company_name = row['Importer']
        if pd.notna(company_name):
            website = search_company_website(company_name)
            if website:
                email = find_email_on_website(website, max_retries=3)
            else:
                email = None

            address = get_address_from_google_maps(company_name)

            data.at[index, 'Website'] = website if website else ""
            data.at[index, 'E-mail'] = ', '.join(email) if email else ""
            data.at[index, 'Address'] = address if address else ""

            for city in cities:
                distance = calculate_distance(address, city)
                data.at[index, city] = distance

        if index % 10 == 0:
            save_progress(data, updated_file_path)
            print(f"Progress saved at row {index}")

        wait_time = random.uniform(2, 5)
        print(f"Waiting for {wait_time:.2f} seconds before the next request...")
        time.sleep(wait_time)

except KeyboardInterrupt:
    print("\nProcess interrupted by the user.")
finally:
    final_file_path = updated_file_path
    temp_file_path = updated_file_path + '.temp'
    if os.path.exists(temp_file_path):
        os.rename(temp_file_path, final_file_path)
        print(f"Final file saved to {final_file_path}")
    else:
        print("No temporary file found to rename.")
