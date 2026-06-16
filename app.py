import streamlit as st
from io import BytesIO
from openpyxl import Workbook

st.set_page_config(
    page_title="HYDAC Lead Processor",
    layout="wide"
)

HEADER = [
    "Referral",
    "Brand",
    "Product",
    "ReceivedDateTime",
    "FirstName",
    "LastName",
    "ContactTitle",
    "Email",
    "Company",
    "Address",
    "County",
    "City",
    "State",
    "ZipCode",
    "Country",
    "LeadSource1",
    "LeadSource2",
    "LeadSource3",
    "LeadComments",
    "PhoneSupplied",
    "PhSuppliedExtension",
    "PhoneResearched",
    "CSRName",
    "PDF",
    "DUNS",
    "WebAddress",
    "Linkedin_Title",
    "Linkedin_Link",
    "SIC",
    "NAICS",
    "noOfEmployees",
    "ParentName",
    "LineOfBusiness",
    "PQ",
    "Latitude",
    "Longitude",
    "DemoLead",
    "ScreenReason",
    "about_me",
    "college_1",
    "college_1_degree",
    "college_1_start",
    "college_1_end",
    "college_2",
    "college_2_degree",
    "college_2_start",
    "college_2_end",
    "month_of_joining",
    "about_experience",
    "searched_on_google",
    "linkedin_city",
    "linkedin_state",
    "linkedin_country",
]

FIXED_COMMENT = 'Please click on "Click Here" below to view customer request.'


def make_excel(file_name: str) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Leads"

    ws.append(HEADER)

    pdf_name = file_name.rsplit(".", 1)[0] + ".pdf"

    row = {h: "" for h in HEADER}
    row["LeadSource1"] = "Email"
    row["LeadComments"] = FIXED_COMMENT
    row["PDF"] = pdf_name

    ws.append([row[h] for h in HEADER])

    for col in ws.columns:
        max_length = 0
        col_letter = col[0].column_letter
        for cell in col:
            value = "" if cell.value is None else str(cell.value)
            max_length = max(max_length, len(value))
        ws.column_dimensions[col_letter].width = min(max_length + 2, 35)

    output = BytesIO()
    wb.save(output)
    return output.getvalue()


st.title("HYDAC Lead Processor")
st.success("Step 2 running: Excel generation enabled.")

uploaded_file = st.file_uploader("Upload MSG file", type=["msg"])

if uploaded_file:
    st.info(f"File uploaded: {uploaded_file.name}")

    excel_bytes = make_excel(uploaded_file.name)

    st.download_button(
        label="Download Excel",
        data=excel_bytes,
        file_name="hydac_lead_output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.write("This step creates an Excel file with your exact header and fixed LeadComments.")
    st.write("Next step: add safe MSG text extraction.")
