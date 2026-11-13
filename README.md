# Publicis Sapient Check-wise Tracker

Streamlit app that builds the "Publicis Sapient Check wise Tracker" Excel report (tracker sheet + native Summary pivot) from the live checkpoint DB.

## Setup
1. `pip install -r requirements.txt`
2. Copy `.env.example` to `.env` and fill in `DB_HOST`, `DB_NAME`, `DB_USER`, `DB_PASSWORD`.
3. `streamlit run app.py --server.port 8599`

Default window is the last 2 years by Received date. All database access is read-only SELECTs.
