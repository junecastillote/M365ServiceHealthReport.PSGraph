# Registering Azure AD App for Automation

Note: This procedure requires that you have one of the following Entra roles, from least to most privileged - *Cloud Application Administrator*, *Application Administrator*, *Global Administrator*.

- [Register a new app](#register-a-new-app)
- [Assign Application API Permissions](#assign-application-api-permissions)
- [Add a Client Secret Credential](#add-a-client-secret-credential)
- [Add a Certificate Credential](#add-a-certificate-credential)

## Register a new app

1. Sign in to the [Microsoft Entra admin center](https://entra.microsoft.com/).
2. Navigate to **Identity -> Applications -> App registrations** and click **New registration**.

    ![register-001](register-001.png)

3. Enter a name for your app. In this example, then app name is `Service Health Report App`.
4. Select **Accounts in this organizational directory only** under the **Supported account types**.
5. Under the **Redirect URI**, select **Web** as the platform and enter `http://localhost` for the URL.
6. Click **Register**.

   ![register-002](register-002.png)

7. Once registered, copy the **Application (client) ID** and **Directory (tenant) ID** values for later use.

    ![register-003](register-003.png)

## Assign Application API Permissions

1. While still on your app's page, click the **API permissions**.

   ![register-004](register-004.png)

2. By default, there is one API permission called `User.Read`. Delete this permisison first by clicking on the context menu (**...**) and **Remove permission**.

    ![register-005](register-005.png)

3. Now the permissions list is empty. Click **Add a permission**.

    ![register-006](register-006.png)

4. Click **Microsoft Graph** from the list of APIs.

    ![register-007](register-007.png)

5. Select **Application permission** and make sure to enable these permissions:

   - `ServiceMessage.Read.All` - Permission to read the announcements from the message center.
   - `ServiceHealth.Read.All` - Permission to read the service health status.
   - `Mail.Send` - Permission to send the report via email using a valid Exchange Online mailbox.

    Once the permissions are selected, click the **Add permissions** button.

    ![register-008](register-008.png)

6. The permissions are now added to the application, as you can see below, but they are not yet granted. Click the **Grant admin consent for [organization]** button.

    ![register-009](register-009.png)

7. Click **Yes** on the confirmation prompt.

    ![register-010](register-010.png)

    Each permission's status has now changed to **Granted for [organization]**.

    ![register-011](register-011.png)

## Add a Client Secret Credential

> **Note**: While the client secret key is a convenient type of app credential, a [certificate credential](#add-a-certificate-credential) is more secure as it does not expose a plain-text key.

A client secret key serves as the password of the application and must be kept as such. ***It has a maximum life of two years and has to be rotated before it expires or the app will stop working due to authentication failure***.

1. Navigate to the app's **Certificates & secrets** page and click **New client secret**.

    ![register-012](register-012.png)

2. Enter the key's description and when it expires, and click **Add**.

    ![register-013](register-013.png)

3. Once created, copy the new key value and keep it safe as you would a password.

    ![register-014](register-014.png)

## Add a Certificate Credential

1. Generate the certificate first if you don't have it yet. Refer to [How to Generate a Self-Signed Certificate for the App][new-cert]

    [new-cert]: ../new-cert/new-certificate.md

2. Once you have the certificate ready, navigate to the **Certificates & secrets** page of the app. and click **Upload certificate**.

    ![register-015](register-015.png)

3. Browse and select the certificate file and click **Add**.

    ![register-016](register-016.png)

4. Confirm the certificate is now visible in the list.

    ![register-017](register-017.png)
