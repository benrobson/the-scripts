import os
import csv
import requests

def get_api_key():
    """Prompts the user for their UniFi API key."""
    api_key = os.environ.get("UNIFI_API_KEY")
    if not api_key:
        api_key = input("Enter your UniFi API key: ")
    return api_key

def main():
    """Main function to run the scraper."""
    api_key = get_api_key()
    headers = {"X-API-KEY": api_key}
    base_url = "https://api.ui.com/v1"

    try:
        # Get all sites
        response = requests.get(f"{base_url}/sites", headers=headers)
        response.raise_for_status()
        sites_response = response.json()
        sites = sites_response.get("data", [])
        print("Successfully retrieved sites.")

        # Get all devices
        response = requests.get(f"{base_url}/devices", headers=headers)
        response.raise_for_status()
        devices_response = response.json()
        hosts = devices_response.get("data", [])
        print("Successfully retrieved devices.")

        all_devices = []
        for host in hosts:
            site_name = host.get("hostName", "N/A")
            for device in host.get("devices", []):
                if device.get("status") != "online":
                    continue

                all_devices.append({
                    "Site": site_name,
                    "Model": device.get("model", "N/A"),
                    "Device Name": device.get("name", "N/A"),
                    "MAC Address": device.get("mac", "N/A"),
                    "Adoption Time": device.get("adoptionTime", "N/A"),
                    "Last Seen": device.get("startupTime", "N/A"),
                    "Status": device.get("status", "N/A"),
                })

        # Write the data to a CSV file
        if all_devices:
            with open("unifi_devices_export.csv", "w", newline="") as csvfile:
                fieldnames = ["Site", "Model", "Device Name", "MAC Address", "Adoption Time", "Last Seen", "Status"]
                writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
                writer.writeheader()
                writer.writerows(all_devices)
            print("Successfully exported devices to unifi_devices_export.csv")
        else:
            print("No devices found to export.")

    except requests.exceptions.RequestException as e:
        print(f"Error communicating with UniFi API: {e}")

if __name__ == "__main__":
    main()
