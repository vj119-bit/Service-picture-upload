# Customer Media Upload Portal

This project gives you a simple Streamlit upload portal for phone users.

## What it does
- User enters customer name
- User optionally enters order/job number
- User uploads one or more photos/videos
- App creates a customer folder automatically in SharePoint
- App renames uploaded files into a clean format
- App can optionally email a notification after upload

## Files included
- `app.py` → main Streamlit app
- `requirements.txt` → Python dependencies
- `.streamlit/config.toml` → raises upload limit
- `.streamlit/secrets.toml.example` → sample secrets file
- `.gitignore` → keeps secrets out of Git

## Step 1 - Microsoft app registration
Create an app registration in Microsoft Entra and save:
- Tenant ID
- Client ID
- Client Secret

## Step 2 - Graph permissions
Add Microsoft Graph application permissions:
- `Files.ReadWrite.All`
- `Sites.Read.All`

Optional for email notification:
- `Mail.Send`

Grant admin consent after adding permissions.

## Step 3 - SharePoint destination
Create a folder or library where uploads should go.
Example:
- Library: `Documents`
- Base folder: `Customer Uploads`

## Step 4 - Configure secrets
Copy `.streamlit/secrets.toml.example` to `.streamlit/secrets.toml` and fill in your real values.

## Step 5 - Install packages
```bash
pip install -r requirements.txt
```

## Step 6 - Run locally
```bash
streamlit run app.py
```

## Step 7 - Deploy to Streamlit Community Cloud
Upload this project to GitHub, then deploy it in Streamlit Community Cloud.

Paste your secrets into the app's **Secrets** section in Streamlit Cloud.

## Notes
- Streamlit Community Cloud should be used as the form/front-end only.
- SharePoint is the actual storage.
- Large videos may take longer to upload.
- The optional email notification uses Graph and requires `Mail.Send`.

## Suggested folder naming
`Customer Uploads / Customer Name - Order Number`

## Suggested file naming
`CustomerName_OrderNumber_YYYYMMDD_HHMMSS_photo_01.jpg`
