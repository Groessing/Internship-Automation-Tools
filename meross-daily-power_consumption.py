import asyncio
import os
from datetime import date, timedelta
from meross_iot.http_api import MerossHttpClient
from meross_iot.manager import MerossManager
import gspread
from google.oauth2.service_account import Credentials

# Returns yesterday's date as a string in YYYY-MM-DD format
def get_yesterday():
    today = date.today()
    yesterday = today - timedelta(days=1)
    return yesterday.strftime("%Y-%m-%d")

# This function collects daily power consumption data from Meross smart devices and records it into a Google Sheets spreadsheet
async def sync_meross_consumption_to_sheets():
    # Initialize the Meross device manager
    client = await MerossHttpClient.async_from_user_password(api_base_url="https://iotx-eu.meross.com", email="email@domain.com", password="password")
    manager = MerossManager(http_client=client, timeout=10)
    await manager.async_init()
    await manager.async_device_discovery()
    devices = manager.find_devices()

    scope = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
    ]

    # Load credentials from service account key
    creds = Credentials.from_service_account_file("key.json", scopes=scope)
    gc = gspread.authorize(creds)

    # Open target spreadsheet
    sh = gc.open_by_key("XXXXXX")
    worksheet = sh.sheet1

    # Read first column (device names) from the sheet
    sheet_devices = worksheet.col_values(1)

    # Map device names to their row index in the sheet
    device_row_map = {name: i+1 for i, name in enumerate(sheet_devices) if name}


    # --- Determine the next available column (for yesterday’s data) ---
    next_col = 2 # Column 1 = device names, so start from col 2
    while True:
        value = worksheet.cell(1, next_col).value
        if value == "" or value is None:
            yesterday = get_yesterday()
            worksheet.update_cell(1, next_col, yesterday) # Write yesterday’s date in row 1
            break
        else:
            next_col += 1 # Move right until we find an empty column
        

        # --- Fetch and record energy consumption ---
        for device in manager.find_devices():
            await device.async_update()

            # Only process devices that support daily consumption stats
            if hasattr(device, "async_get_daily_power_consumption"):
                daily_consumption =  await device.async_get_daily_power_consumption()
                device_name = device.name

                # Check if this device exists in the sheet
                if device_name in device_row_map:
                    row = device_row_map[device_name]

                    # Iterate through consumption history entries
                    for entry in daily_consumption:
                        date_str = entry['date'].strftime("%Y-%m-%d")
                        consumption_str = str(entry['total_consumption_kwh'])

                        # Only record yesterday’s value (ignore older entries)
                        if date_str == yesterday:
                            worksheet.update_cell(row, next_col, consumption_str)
                            break

async def main():
    await sync_meross_consumption_to_sheets()


if name == "__main__":
    asyncio.run(main())
