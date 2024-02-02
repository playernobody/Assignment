import os
import requests
import googlemaps
import pandas as pd
import openpyxl

# Constants
GOOGLE_MAPS_API_KEY = 'AIzaSyC3Mfkc0ib2yqkJ3cQCiJ2DDyTyeF8snno'
WHATSAPP_API_KEY = 'UAKbd8e26ad-01ab-4aaf-8859-9989a1185b2b'
YOUR_PHONE_NUMBER = '+918580610536'

def get_google_maps_client(api_key):
    return googlemaps.Client(key=api_key)

def get_business_details(gmaps, place_id):
    try:
        result = gmaps.place(place_id=place_id)
        return result['result']
    except Exception as e:
        print(f"Error fetching business details: {e}")
        return None

def is_whatsapp_verified(api_key, your_number, check_number):
    try:
        url = f"https://api.p.2chat.io/open/whatsapp/check-number/{your_number}/{check_number}?extra-information=false"
        headers = {'X-User-API-Key': api_key}
        response = requests.get(url, headers=headers)
        data = response.json()

        # Check if the number is on WhatsApp
        return data.get("on_whatsapp", False)
    except Exception as e:
        print(f"Error verifying WhatsApp: {e}")
        return False

def save_business_details_to_excel(df, excel_file_path):
    if os.path.isfile(excel_file_path):
        existing_df = pd.read_excel(excel_file_path, engine='openpyxl')
        combined_df = pd.concat([existing_df, df], ignore_index=True)
        combined_df.to_excel(excel_file_path, index=False, engine='openpyxl')
        print(f"Business details appended to {excel_file_path}")
    else:
        new_file_path = 'business_details.xlsx'
        df.to_excel(new_file_path, index=False, engine='openpyxl')
        print(f"Business details saved to {new_file_path}")

def main():
    gmaps = get_google_maps_client(GOOGLE_MAPS_API_KEY)
    place_id = 'ChIJm-UHpDY_BTkR5jrKZyE1mSI'

    business_details = get_business_details(gmaps, place_id)

    if business_details:
        data = {}
        for key, value in business_details.items():
            if isinstance(value, (str, int, float, bool)):
                data[key] = [value]
            elif isinstance(value, dict):
                for sub_key, sub_value in value.items():
                    data[f'{key}_{sub_key}'] = [sub_value]
            elif isinstance(value, list):
                for i, item in enumerate(value):
                    if isinstance(item, dict):
                        for sub_key, sub_value in item.items():
                            data[f'{key}_{sub_key}_{i + 1}'] = [sub_value]

        df = pd.DataFrame(data)
        excel_file_path = 'business_details.xlsx'
        save_business_details_to_excel(df, excel_file_path)
        print(f"Business details saved to {excel_file_path}")

if __name__ == "__main__":
    main()
