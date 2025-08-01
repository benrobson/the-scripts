# Mailbox Search Tool

This tool recursively searches a Microsoft 365 mailbox for all emails sent to or received from a specific person and logs the results to a CSV file.

## Prerequisites

To use this tool, you need to have an Azure AD application with the appropriate permissions. Follow these steps to create and configure the application:

### 1. Register an Application in Azure AD

1.  Go to the [Azure portal](https://portal.azure.com).
2.  Navigate to **Azure Active Directory**.
3.  Go to **App registrations** and click **+ New registration**.
4.  Give your application a name (e.g., "Mailbox Search Tool").
5.  For **Supported account types**, select **Accounts in this organizational directory only (Default Directory only - Single tenant)**.
6.  You can leave the **Redirect URI** blank for now.
7.  Click **Register**.

### 2. Get Application (client) ID and Directory (tenant) ID

1.  After the app is created, you'll be taken to its overview page.
2.  Copy the **Application (client) ID** and the **Directory (tenant) ID**. You will need these to run the script.

### 3. Configure Authentication

1.  In your app registration, go to the **Authentication** tab.
2.  Scroll down to the **Advanced settings** section.
3.  Set **Allow public client flows** to **Yes**.
4.  Click **Save**.

### 4. Add API Permissions

1.  Go to the **API permissions** tab.
2.  Click **+ Add a permission**.
3.  Select **Microsoft Graph**.
4.  Select **Delegated permissions**.
5.  Search for and select `Mail.Read`.
6.  Click **Add permissions**.
7.  After adding the permission, you may need to grant admin consent. Click the **Grant admin consent for [Your Tenant]** button.

## How to Run the Script

1.  Install the required Python libraries:
    ```bash
    pip install msal requests
    ```
2.  Run the script from the command line:
    ```bash
    python mailbox_search.py
    ```
3.  The script will prompt you to enter the following information:
    *   **Application (client) ID**: The ID you copied from your Azure AD app registration.
    *   **Directory (tenant) ID**: The ID you copied from your Azure AD app registration.
    *   **The email address of the mailbox to search**: The email address of the user whose mailbox you want to search.
    *   **The email address to search for (sender or recipient)**: The email address of the person whose emails you are looking for.

4.  The first time you run the script, it will provide you with a code and a URL to sign in and authorize the application.

5.  Once the script has finished, it will create a CSV file in the same directory with the search results.
