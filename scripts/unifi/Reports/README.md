# UniFi Device Exporter

This script connects to the UniFi Site Manager API to extract a list of all UniFi devices across all sites associated with your account. It then outputs this data to a CSV file.

## Prerequisites

*   Python 3.6+
*   The `requests` Python library. You can install it using `pip`:
    ```
    pip install requests
    ```
*   A UniFi account with API access enabled.
*   Your UniFi API key. You can find instructions on how to generate an API key in the UniFi Site Manager API documentation.

## How to Use

1.  **Clone or download the script.**
    *   `unifi_api_scraper.py`
2.  **Run the script from your terminal:**
    ```
    python unifi_api_scraper.py
    ```
3.  **Enter your UniFi API Key:**
    *   The script will prompt you to enter your UniFi API key. Paste your key and press Enter.
    *   Alternatively, you can set the `UNIFI_API_KEY` environment variable to your API key to avoid being prompted.
4.  **Output:**
    *   The script will connect to the UniFi API, fetch the device data, and create a file named `unifi_devices_export.csv` in the same directory.
    *   The CSV file will contain the following columns:
        *   **Site:** The name of the site the device belongs to.
        *   **Model:** The model of the device.
        *   **Device Name:** The alias or name of the device.
        *   **MAC Address:** The MAC address of the device.
        *   **Adoption Time:** The time the device was adopted.
        *   **Last Seen:** The last time the device was seen online.
        *   **Status:** The current status of the device.

## Notes

*   The script only exports devices that are currently online. Offline and unadopted devices are ignored.
*   If you encounter any issues, make sure your API key is correct and has the necessary permissions.
*   The script uses the official UniFi Site Manager API, so it should be reliable and efficient.
