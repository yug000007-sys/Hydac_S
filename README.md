# HYDAC Lead Processor - Step 3

This version adds:
- MSG parsing
- Customer detail extraction from email text
- Excel output
- Attachment filtering

Attachment rule:
- Ignore generic/signature images like image.png, image001.png, logo.png
- Keep specific customer/request attachments like BIERI 3999534.png

Deploy settings:
- Main file path: app.py
- Branch: main
