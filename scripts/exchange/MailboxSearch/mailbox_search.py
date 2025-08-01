import csv
import json
import os
import msal
import requests

# Constants for Graph API
GRAPH_API_ENDPOINT = 'https://graph.microsoft.com/v1.0'
SCOPES = ['Mail.Read']

# Function to acquire a token using a client secret
def acquire_token(app_id, tenant_id, client_secret):
    """Acquire a token from Azure AD using a client secret."""
    authority = f"https://login.microsoftonline.com/{tenant_id}"
    app = msal.ConfidentialClientApplication(
        client_id=app_id,
        authority=authority,
        client_credential=client_secret
    )

    # The scope for client credentials flow is different
    result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    return result

# Function to search emails
def search_emails(access_token, user_id, search_query):
    """Search emails in a user's mailbox."""
    headers = {'Authorization': 'Bearer ' + access_token}

    # Select only the fields we need to make the response smaller
    select_fields = "subject,from,toRecipients,receivedDateTime"
    search_url = f"{GRAPH_API_ENDPOINT}/users/{user_id}/messages?$search=\"{search_query}\"&$select={select_fields}"

    all_messages = []
    page_count = 1
    while search_url:
        print(f"Fetching page {page_count}...")
        response = requests.get(search_url, headers=headers)
        if response.status_code != 200:
            raise Exception(f"Graph API returned error: {response.status_code} {response.text}")

        data = response.json()
        all_messages.extend(data.get('value', []))
        search_url = data.get('@odata.nextLink')
        page_count += 1
        print(f"Found {len(all_messages)} emails so far...")

    return all_messages

# Function to write emails to CSV
def write_to_csv(emails, filename):
    """Write email data to a CSV file."""
    if not emails:
        print("No emails to write.")
        return

    fieldnames = ['receivedDateTime', 'subject', 'from', 'toRecipients']

    with open(filename, 'w', newline='', encoding='utf-8') as csvfile:
        writer = csv.writer(csvfile)

        # Write header
        writer.writerow(['Received Date Time', 'Subject', 'From', 'To'])

        # Write email data
        for email in emails:
            from_address = email.get('from', {}).get('emailAddress', {}).get('address', 'N/A')

            to_recipients = email.get('toRecipients', [])
            to_addresses = '; '.join([
                recipient.get('emailAddress', {}).get('address', 'N/A')
                for recipient in to_recipients
            ])

            writer.writerow([
                email.get('receivedDateTime', 'N/A'),
                email.get('subject', 'N/A'),
                from_address,
                to_addresses
            ])

def main():
    """Main function to run the script."""
    # Get user input for app_id, tenant_id, user_id, and search_email
    app_id = input("Enter your Application (client) ID: ")
    tenant_id = input("Enter your Directory (tenant) ID: ")
    client_secret = input("Enter your Client Secret: ")
    user_id = input("Enter the email address of the mailbox to search: ")
    search_email = input("Enter the email address to search for (sender or recipient): ")

    # Acquire token
    token_result = acquire_token(app_id, tenant_id, client_secret)

    if "access_token" in token_result:
        access_token = token_result['access_token']
        # Construct search query
        search_query = f"from:{search_email} OR to:{search_email}"
        # Search emails
        print("Searching for emails...")
        try:
            emails = search_emails(access_token, user_id, search_query)
            print(f"Found a total of {len(emails)} emails.")

            # Write to CSV
            if emails:
                output_filename = f"email_log_{user_id.replace('@', '_')}.csv"
                write_to_csv(emails, output_filename)
                print(f"Results written to {output_filename}")
            else:
                print("No emails found matching the criteria.")
        except Exception as e:
            print(f"An error occurred: {e}")
    else:
        print("Failed to acquire token.")
        print(token_result.get("error"))
        print(token_result.get("error_description"))
        print(token_result.get("correlation_id"))

if __name__ == '__main__':
    main()
