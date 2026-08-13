# Setup Guide

## Windows Credential Manager Configuration

Follow these steps to store your database configuration securely in Windows Generic Credentials.

### Quick Steps

1. Open the **Start Menu** and search for **Credential Manager**.
2. Click on **Windows Credentials**.
3. Under *Generic Credentials*, click **Add a generic credential**.
4. Enter the details listed below and click **OK**.

---

### Credential Details

| Field | Value |
| :--- | :--- |
| **Internet or network address** | `SpeedpointRecon` |
| **User name** | `DATABASE_URL` |
| **Password** | `database_secret_url` |

> 💡 **Note:** Replace `database_secret_url` with your actual secret database connection string if necessary.