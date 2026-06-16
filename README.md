# HYDAC Lead Agent V1

This is an agent-style Streamlit app for HYDAC lead processing.

## What it does

- Upload `.msg` email files
- Reads full email trail
- Uses AI to identify the real external customer, not HYDAC/internal forwarders
- Extracts customer/contact/company/request details
- Filters real customer attachments from signature images
- Shows a review panel before Excel export
- Exports Excel in the HYDAC lead header format
- Exports valid attachments as ZIP

## Deploy on Streamlit

Upload only these files to GitHub:

- app.py
- requirements.txt
- runtime.txt
- README.md

Streamlit settings:

- Main file path: app.py
- Branch: main

## Streamlit Secrets

Add this in Streamlit Cloud secrets:

OPENAI_API_KEY = "your_api_key_here"

Without the key, the app still opens, but AI extraction will not run.

## Current working rules

1. This is an agent, not a fixed extractor.
2. It should understand the real external customer from the mail trail.
3. Ignore HYDAC/internal senders and forwarders.
4. If the email contains an EXTERNAL EMAIL block, prefer that external block.
5. Customer request should be the actual request text, not the whole forwarded trail.
6. Split name into FirstName and LastName.
7. PhoneSupplied must include customer phones only, not HYDAC signature phones.
8. Phone format example: 43-31669151 : 43-6649614266
9. Ignore signature images such as image.png, image001.png, logos, icons.
10. Keep meaningful customer attachments like BIERI 3999534.png.
11. Leave unavailable fields blank. Do not guess.
