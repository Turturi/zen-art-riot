# ZEN EVENT — Streamlit College Fest Registration

A QR-driven festival registration website with eight clickable event cards, separate registration forms, one Excel workbook with **eight event-specific worksheets**, and a password-protected admin dashboard.

## Features

- Singing, Reel Roulette, Film Making, Dance, Antakshari, Mr Sau Ms Sau, Star Night, Fashion Show.
- Each event opens its own form (name, 10-digit Indian mobile number, college, department, year, consent).
- **The same mobile number may register for different events AND register multiple times for the same event.** Every successful submission creates a fresh registration ID and Excel row.
- Each event has its own Excel tab; admin can filter/search, view totals, and download the workbook.
- Admin can generate a QR image and printable festival poster linked to the public website.
- Local Excel storage for testing, or Azure Blob Storage for persistent online storage with ETag-based concurrent write protection.

## Windows PowerShell — run locally

```powershell
cd "C:\path\to\zen_event_final"
py -3.12 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
Copy-Item .streamlit\secrets.toml.example .streamlit\secrets.toml
# Edit .streamlit\secrets.toml and set a strong ADMIN_PASSWORD
python -m streamlit run app.py
```

Visit http://localhost:8501. If `python` is not on PATH, run `.\.venv\Scripts\python.exe -m streamlit run app.py`.

## Excel file

In local mode the workbook is saved to `college_fest_registrations.xlsx` in the **working directory**. If it is open in Excel while registering, Windows may prevent saving: close the workbook first. **Do not overwrite your old workbook when updating the code.** The workbook's eight tabs are named after the events. Export the workbook from the admin dashboard.

## Publish and create the scannable QR code

1. Push the project to GitHub. **Do not commit `.streamlit/secrets.toml` or the Excel workbook.**
2. Create an Azure Storage Account and private Blob container called `zen-event`.
3. Deploy the GitHub repository on https://share.streamlit.io/ with entry point `app.py`.
4. In the Streamlit Cloud app settings, add the contents of `secrets.toml.example`, replacing the placeholders. Set `AZURE_STORAGE_CONNECTION_STRING` to your Azure Storage connection string and `REGISTRATION_URL` to the exact public `https://...streamlit.app` URL. Use a strong admin password.
5. Visit the Admin Dashboard, log in, and download the QR image or printable poster. Scanning the QR opens the public website on students' phones.
6. Test two registrations using the **same mobile number** for Singing, then one for Dance. Confirm all three rows in the exported workbook.

**Note:** Streamlit Community Cloud local files are not persistent. Azure Blob Storage (or other persistent storage) is required for production. Azure usage may incur charges. The Excel file contains personal data; restrict access and retain it only as needed.

## Files

- `app.py`: interface, event selection, forms, admin dashboard, QR poster
- `storage.py`: Excel workbook, eight tabs, unique IDs, Azure/local persistence
- `requirements.txt`: dependencies
- `.streamlit/config.toml`: theme
- `.streamlit/secrets.toml.example`: example settings
- `test_storage.py`: local automated test for repeated numbers and separate tabs
