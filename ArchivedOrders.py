import os
import pandas as pd
import requests
import time
import datetime
from tqdm import tqdm
import atexit
import subprocess

# Force console window to stay open after execution
def keep_console_open():
    subprocess.call("pause", shell=True)

atexit.register(keep_console_open)

# Read the Excel file into a DataFrame
df = pd.read_excel('OrderIDs.xlsx', sheet_name='Sheet1')

# Convert Order ID column to string
df['Order ID'] = df['Order ID'].astype(str)

# Extract the order IDs from the 'Order ID' column
order_ids = df['Order ID'].tolist()

# Mark the archived orders as paid using the Shopify API
api_url = "https://nesaa-shop.myshopify.com/admin/api/2023-04/orders/{{order_id}}/transactions.json"
username = "5bcd72a3f88118c0ffb7cefd4427c153"
password = "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"

# Create a new column for Reason
df['Reason'] = ''

start_time = time.time()  # Start time

# Use tqdm to show the progress bar
with tqdm(total=len(order_ids), desc="Processing Orders") as pbar:
    for index, order_id in enumerate(order_ids):
        # Mark the order as paid
        transaction_url = api_url.replace('{{order_id}}', order_id)
        transaction_data = {
            "transaction": {
                "kind": "capture"
            }
        }
        response = requests.post(transaction_url, json=transaction_data, auth=(username, password))

        if response.status_code == 201:
            df.at[index, 'Reason'] = 'Marked as Paid'
        else:
            df.at[index, 'Reason'] = response.json().get('errors')

        # Update the progress bar
        pbar.update(1)

end_time = time.time()  # End time
elapsed_time = end_time - start_time  # Calculate elapsed time

# Print the total number of orders
total_orders = len(order_ids)
print(f"Total Orders: {total_orders}")

# Print the time taken to finish
print(f"Time Taken: {elapsed_time} seconds")

# Save the updated DataFrame to a file with today's date and version number
output_folder = 'output'
os.makedirs(output_folder, exist_ok=True)
output_file = f"{output_folder}/OrdersIDs_{datetime.date.today()}_V1.xlsx"
version = 2
while os.path.exists(output_file):
    output_file = f"{output_folder}/OrdersIDs_{datetime.date.today()}_V{version}.xlsx"
    version += 1

df.to_excel(output_file, sheet_name='Sheet1', index=False)
