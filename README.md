# Flight Data Formatter ✈️

A Streamlit app that formats and processes aircraft daily ops Excel files.

## 📦 Features
- Sorts flights by category (Transit, On Call, Canceled)
- Converts STA, ATA, STD, and ATD into `MM/DD/YYYY HH:MM:SS` format
- Auto-fills services based on checkmarks (√)
- Cleans up employee codes and other flight metadata

## 🚀 How to Use Locally
```bash
pip install -r requirements.txt
streamlit run flight_formatter.py
```

## 🌐 Deploy on Streamlit Cloud
Upload these files to a GitHub repo and deploy at [streamlit.io/cloud](https://streamlit.io/cloud).
