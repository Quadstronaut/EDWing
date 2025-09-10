import os
import json
import requests
import datetime
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options

# Configuration
WEBHOOK_URL = "YOURWEBHOOKHERE"
PRICE_FILE = "wine_price_history.json"
INARA_URL = "https://inara.cz/elite/station-market/230278/"
CHROMEDRIVER_PATH = '/usr/bin/chromedriver'
PRICE_THRESHOLD = 250000

def send_discord_message(message):
    """Sends a message to the Discord webhook."""
    payload = {
        "content": message
    }
    try:
        response = requests.post(WEBHOOK_URL, json=payload)
        response.raise_for_status()
        print("Successfully sent message to Discord.")
    except requests.exceptions.RequestException as e:
        print(f"Failed to send message to Discord: {e}")

def get_current_price():
    """Fetches the current price of Wine from Inara.cz using Selenium."""
    chrome_options = Options()
    chrome_options.add_argument("--headless")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")

    service = Service(executable_path=CHROMEDRIVER_PATH)

    try:
        driver = webdriver.Chrome(service=service, options=chrome_options)
        driver.get(INARA_URL)
        
        # Find the row for 'Wine'
        wine_row = driver.find_element(By.XPATH, "//a[text()='Wine']/ancestor::tr")

        # Get the sell price from the data-order attribute
        sell_price_element = wine_row.find_element(By.XPATH, ".//td[2]")
        wine_price_str = sell_price_element.get_attribute("data-order")
        
        return int(wine_price_str)

    except Exception as e:
        print(f"An error occurred while fetching the price: {e}")
        return None
    finally:
        if 'driver' in locals():
            driver.quit()

def get_price_history():
    """Reads the price history from a local file, including notification state.
       Handles conversion from the old single-record format to the new list format."""
    if os.path.exists(PRICE_FILE):
        try:
            with open(PRICE_FILE, 'r') as f:
                data = json.load(f)
                
                # Check for the old, single-record format
                if "wine_price" in data and not isinstance(data.get("history"), list):
                    print("Converting old file format to new history format.")
                    old_price = data.get("wine_price")
                    old_timestamp = data.get("last_updated")
                    notified_state = data.get("notified_over_250k", False)
                    return {
                        "history": [{"price": old_price, "timestamp": old_timestamp}],
                        "notified_over_250k": notified_state
                    }
                
                # Ensure the new key exists, for backward compatibility
                if "notified_over_250k" not in data:
                    data["notified_over_250k"] = False
                if "history" not in data:
                    data["history"] = []
                    
                return data
        except (IOError, json.JSONDecodeError) as e:
            print(f"Error reading from file: {e}")
            # Fallback to an empty structure if file is corrupted
    return {"history": [], "notified_over_250k": False}

def update_price_history(current_price, data):
    """Updates the price history, prunes old data, and saves to file."""
    
    # Prune data older than 365 days
    pruning_threshold = datetime.datetime.now() - datetime.timedelta(days=365)
    
    # Filter out entries older than the threshold
    data['history'] = [
        entry for entry in data['history']
        if datetime.datetime.fromisoformat(entry['timestamp']) > pruning_threshold
    ]
    
    # Append the new price entry
    new_entry = {
        "price": current_price,
        "timestamp": datetime.datetime.now().isoformat()
    }
    data['history'].append(new_entry)
    
    # Save the updated data
    try:
        with open(PRICE_FILE, 'w') as f:
            json.dump(data, f, indent=4)
        print("Price history and notification state successfully saved to file.")
    except IOError as e:
        print(f"Error writing to file: {e}")

def main():
    """Main function to check and notify about price changes based on threshold."""
    current_price = get_current_price()

    if current_price is None:
        print("Could not retrieve the current price. Exiting.")
        return

    price_data = get_price_history()
    notified = price_data.get('notified_over_250k', False)

    # Calculate the profit
    profit = current_price * 60000
    
    # Check if the price has exceeded the threshold for the first time
    if current_price > PRICE_THRESHOLD and not notified:
        message = f"# **Wine Price Alert!**\n\nThe price of Wine has been detected at a new high of **{current_price} Cr.**, exceeding the {PRICE_THRESHOLD} Cr. threshold. The profit is estimated to be approximately **{profit:,} Cr.**"
        send_discord_message(message)
        print(message)
        
        # Update the notification state to True
        price_data['notified_over_250k'] = True
    
    # Check if the price has dropped back below the threshold to reset the state
    elif current_price <= PRICE_THRESHOLD:
        if notified:
            print(f"Price has dropped back below {PRICE_THRESHOLD} Cr. Resetting notification state.")
        price_data['notified_over_250k'] = False
    
    else:
        print(f"Wine price is {current_price} Cr., but has already been notified. No new alert will be sent.")

    # Always update the price history and prune old entries
    update_price_history(current_price, price_data)

if __name__ == "__main__":
    main()
