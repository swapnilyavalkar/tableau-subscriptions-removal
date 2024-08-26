---

# 📊 Tableau Server Subscription Cleanup Script

This Python script automates the cleanup of Tableau Server subscriptions for unlicensed users. It identifies users who no longer have a license, deletes their subscriptions, and notifies them via email. Additionally, it logs the entire process and sends summary emails to the admin team.

## 🚀 Features

- **Identify Unlicensed Users**: Fetches users from Tableau Server who no longer have a license.
- **Delete Subscriptions**: Automatically deletes subscriptions for unlicensed users.
- **Email Notifications**: Sends an email to users whose subscriptions have been removed, with details of the removed subscriptions.
- **Admin Notifications**: Sends a success or failure summary email to the Tableau Admin team.
- **Log Management**: Cleans up old log files to manage disk space.

## 📋 Prerequisites

Before running the script, ensure you have the following installed:

- **Python 3.x**
- **pandas**: For data manipulation and Excel file generation.
- **tableauserverclient**: To interact with the Tableau Server REST API.
- **smtplib**: For sending emails.
- **logging**: For detailed logging of the script's operations.

You can install the required Python packages using pip:

```bash
pip install pandas tableauserverclient
```

## 🛠️ Configuration

### SMTP Settings

Update the SMTP host details to enable email notifications:

```python
smtp_host = 'mail.abc.com'
smtp_port = 25
FROM = 'admin@abc.com'
CC = ''
admin_dl = 'admin@abc.com'
```

### Tableau Server Login Details

Specify your Tableau Server URL and credentials:

```python
server_url = 'https://abc.com/'
sites = ''
username = ''
password = ''
```

## 📂 Directory Structure

Ensure the following directory structure is present:

```
/logs            # Directory to store log files
/data            # Directory to store Excel files with subscription data
```

## 📝 Script Overview

### 1. **delete_subscriptions()**

This function deletes subscriptions for all unlicensed users on the Tableau Server:

- **Input**: List of unlicensed users.
- **Output**: Excel file with details of deleted subscriptions.

### 2. **get_unlicensed_users()**

Fetches all users from Tableau Server with an "Unlicensed" role:

- **Output**: List of unlicensed users and Excel file with their details.

### 3. **save_data_excel()**

Saves the unlicensed users and deleted subscription details to Excel files.

### 4. **send_user_email()**

Sends an email to each user whose subscriptions were deleted, detailing the removed subscriptions.

### 5. **send_success_email_admin()**

Sends a success email to the admin team, summarizing the cleanup operation.

### 6. **send_failed_email_admin()**

Sends a failure email to the admin team if the cleanup operation fails.

### 7. **delete_logs()**

Deletes log files that are older than 10 days to manage disk space.

## 🔧 How to Use

1. **Set up SMTP and Tableau Server configurations** in the script as described in the Configuration section.

2. **Run the script** to start the cleanup operation:

```bash
python main.py
```

3. **Monitor the log files** in the `/logs` directory for detailed information on the script's execution.

4. **Check the `/data` directory** for Excel files containing details of unlicensed users and deleted subscriptions.

5. **Review email notifications** sent to users and admins for success or failure of the cleanup operation.

## 📅 Scheduled Execution

To automate the execution, you can schedule the script using a task scheduler like **Windows Task Scheduler** or **cron** on Linux. This will enable regular cleanup of subscriptions without manual intervention.

## 🛡️ Security Considerations

- **Credentials**: Ensure that the Tableau Server username and password are securely stored, and consider using environment variables or secure vaults.
- **SMTP**: Make sure your SMTP server is configured securely, especially if it's exposed to the internet.

## 📄 License

This script is provided under the MIT License. Feel free to use and modify it to suit your needs.

---
