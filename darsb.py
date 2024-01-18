from nordpool import elspot
from datetime import datetime, timedelta
import pandas as pd

def convert_to_unaware(dt):
    return dt.replace(tzinfo=None)

def fetch_and_save_hourly_prices_to_excel():
    prices_spot = elspot.Prices()

    try:
        elspot_prices_latvia = prices_spot.hourly(areas=['LV'])

        if not elspot_prices_latvia:
            print("No hourly prices available.")
            return

        data = {'Date': [], 'Price (EUR/kWh)': [], 'Message': []}

        for area, values in elspot_prices_latvia.get('areas', {}).items():
            for entry in values.get('values', []):
                start_time = entry.get('start', 'N/A')
                price_mwh = entry.get('value', 'N/A')

                price_kwh = price_mwh / 1000.0

                if price_kwh > 0.10:
                    message = "Šajā stundā elektrību nevajadzētu izmantot"
                else:
                    message = ""

                start_time_unaware = convert_to_unaware(start_time)
                formatted_date = start_time_unaware.strftime('%Y-%m-%d %H:%M:%S')

                data['Date'].append(formatted_date)
                data['Price (EUR/kWh)'].append(price_kwh)
                data['Message'].append(message)

        df = pd.DataFrame(data)

        user_choice = input("Vai jūs gribētu zināt cenu konkrētā stundā? (yes/no): ").lower()

        if user_choice == 'yes':
            specific_time = input("Ierakstiet konkrēto datumu un laiku (YYYY-MM-DD HH:MM:SS): ")
            specific_data = df[df['Date'] == specific_time]

            if specific_data.empty:
                print(f"No data available for {specific_time}")
            else:
                specific_price = specific_data['Price (EUR/kWh)'].values[0]
                specific_message = specific_data['Message'].values[0]

                print(f"Cena {specific_time}: {specific_price} EUR/kWh")
                if specific_message:
                    print(f"Brīdinājums: {specific_message}")

        elif user_choice == 'no':
            excel_file_path = "Stundas_cena.xlsx"
            df.to_excel(excel_file_path, index=False)
            print(f"Cenas pa stundai ir saglabātas {excel_file_path}")

        else:
            print("Nepareiza ievade lūdzu izmantot: 'yes' or 'no'.")

    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    fetch_and_save_hourly_prices_to_excel()
