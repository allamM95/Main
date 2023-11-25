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

# Extract the order numbers from the 'Order Number' column
order_numbers = df['Order Number'].tolist()

# Shopify API credentials
username = 'f19cdf9afc43f6045d8c30440129f0c3'
password = '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'
shop = 'bfe-world'

# Mark the orders as paid using the Shopify API
api_url = f"https://f19cdf9afc43f6045d8c30440129f0c3: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@bfe-world.myshopify.com/admin/api/2023-04/orders.json"

# Create a new column for Order ID, Reason, and API Response
df['Order ID'] = ''
df['Reason'] = ''
df['API Response'] = ''  # Add this line to create a new column for API Response

start_time = time.time()  # Start time

# Use tqdm to show the progress bar
with tqdm(total=len(order_numbers), desc="Processing Orders") as pbar:
    for index, order_number in enumerate(order_numbers):
        # Get the order ID based on the order number
        query_params = {'name': order_number}
        response = requests.get(api_url, params=query_params)

        if response.status_code == 200:
            orders = response.json().get('orders')
            if orders:
                order_id = orders[0]['id']
                df.at[index, 'Order ID'] = order_id

                # Mark the order as paid
                transaction_url = f"https://f19cdf9afc43f6045d8c30440129f0c3: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@bfe-world.myshopify.com/admin/api/2023-04/orders/{order_id}/transactions.json"
                transaction_data = {
                    "transaction": {
                        "kind": "capture"
                    }
                }
                response_transaction = requests.post(transaction_url, json=transaction_data)

                if response_transaction.status_code == 201:
                    df.at[index, 'Reason'] = 'Marked as Paid'
                else:
                    df.at[index, 'Reason'] = response_transaction.json().get('errors')

                # Store the API response in the 'API Response' column
                df.at[index, 'API Response'] = response_transaction.text  # or use response_transaction.json() if you want the JSON response
            else:
                df.at[index, 'Reason'] = 'Order not found'
        else:
            df.at[index, 'Reason'] = response.text
            # Store the API response even in case of an error
            df.at[index, 'API Response'] = response.text  # This line captures the API response text in case of an error

        # Update the progress bar
        pbar.update(1)

end_time = time.time()  # End time
elapsed_time = end_time - start_time  # Calculate elapsed time

# Print the total number of orders
total_orders = len(order_numbers)
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
