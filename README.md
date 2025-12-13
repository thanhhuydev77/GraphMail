# Email - Graph User Connector for Business Central

## Overview

This extension for Microsoft Dynamics 365 Business Central provides a custom email connector that enables sending emails through the Microsoft Graph API using a shared email identity (service-to-service authentication). It is designed for scenarios where automated emails need to be sent from a system account rather than a specific user's account.

The connector authenticates using OAuth 2.0 client credentials, ensuring a secure connection to the Graph API without storing user passwords.

## Features

- **Shared Email Account**: Send emails from a single, configured email address across the organization.
- **Secure Authentication**: Utilizes OAuth 2.0 client credentials flow for secure, service-to-service communication with Microsoft Graph.
- **Large Attachment Support**: Automatically handles large attachments by creating an upload session and sending them in chunks, overcoming the standard size limits of a single API request.
- **Easy Configuration**: A simple setup page allows for easy configuration of the Email Address, Client ID, Client Secret, and Tenant ID.
- **Standard Interface Implementation**: Implements the standard `Email Connector v4` interface, ensuring seamless integration with the Business Central email framework.

## Code Flow (Sending an Email)
<img width="1420" height="810" alt="image" src="https://github.com/user-attachments/assets/1aa931de-5c1c-41bd-b786-69d70ed8d80e" />

## Setup and Configuration

1.  **Register an Application in Azure Active Directory**:
    - Go to the Azure portal and register a new application.
    - Create a new **Client Secret** and copy the value securely.
    - Under **API Permissions**, add the `Mail.Send` permission for **Microsoft Graph** (Application permission).
    - Grant admin consent for the permission.
2.  **Configure in Business Central**:
    - In Business Central, search for **Email Accounts**.
    - Run the **New** action.
    - Select the **Shared Email** account type.
    - Fill in the **Email Address**, **Account Name**, **Tenant ID**, **Client ID**, and the **Client Secret** you generated in Azure AD.
    - Click **Next** to save the account.
    - Set this account as the default if desired.
