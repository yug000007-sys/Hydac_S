# HYDAC Lead Agent - Step 5

This version adds AI customer understanding.

Features:
- MSG upload
- Reads full email body/trail
- Uses AI to identify actual customer/contact, not HYDAC forwarder
- Ignores signature images
- Keeps valid customer attachments
- Shows Agent Review Panel before Excel export

Streamlit setup:
1. Upload these files to GitHub:
   - app.py
   - requirements.txt
   - runtime.txt
   - README.md

2. Deploy on Streamlit Cloud:
   - Main file path: app.py

3. Add OpenAI key in Streamlit:
   - App Settings
   - Secrets
   - Add:

OPENAI_API_KEY = "your_api_key_here"

If no OpenAI key is added, the app will still run in manual/rule mode.
