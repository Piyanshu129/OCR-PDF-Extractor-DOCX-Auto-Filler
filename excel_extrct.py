# # import pandas as pd
# # import streamlit as st
# #
# # st.title("Excel Field Extractor")
# #
# # # Upload the Excel file
# # uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")
# #
# # if uploaded_file is not None:
# #     # Read the Excel file
# #     df = pd.read_excel(uploaded_file, sheet_name=0)
# #
# #     # Display the DataFrame
# #     st.write("Data from the Excel sheet:")
# #     st.write(df)
# #
# #     # Extract fields based on serial numbers
# #     serial_numbers = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10]  # Modify as needed
# #
# #     # Create a DataFrame with extracted fields
# #     extracted_df = df[df['Sr. No.'].isin(serial_numbers)]
# #     st.write("Extracted Fields:")
# #     st.write(extracted_df)
# #
# # # Run the app with: streamlit run your_script.py
# #
# # import streamlit as st
# # import pandas as pd
# # import re
# #
# # st.set_page_config(layout="wide")
# # st.title("GST Tax Invoice Field Extractor")
# #
# # uploaded_file = st.file_uploader("Upload Invoice Excel", type=["xlsx"])
# #
# # if uploaded_file:
# #     # Read all sheets
# #     xls = pd.ExcelFile(uploaded_file)
# #     df = xls.parse(xls.sheet_names[0], header=None)
# #
# #     st.subheader("Raw Excel Preview")
# #     st.dataframe(df)
# #
# #     # Step 1: Identify the header row
# #     # Detect header row from the Excel sheet more flexibly
# #     header_keywords = ["Sr", "Product", "Description", "HSN", "Qty", "Rate", "Amount", "Taxable", "IGST", "Total"]
# #     header_row_index = None
# #
# #     for i, row in df.iterrows():
# #         match_count = sum(
# #             any(k.lower() in str(cell).lower() for cell in row if pd.notna(cell))
# #             for k in header_keywords
# #         )
# #         if match_count >= 4:
# #             header_row_index = i
# #             break
# #
# #     if header_row_index is not None:
# #         st.success(f"Header row found at row {header_row_index}")
# #         # Extract data starting from header row
# #         table_df = pd.read_excel(uploaded_file, header=header_row_index)
# #         table_df = table_df.dropna(how='all')
# #
# #         st.subheader("Extracted Table")
# #         st.dataframe(table_df)
# #     else:
# #         st.error("❌ Could not detect a valid header row.")
# # import streamlit as st
# # import pandas as pd
# # import io
# # from openpyxl import Workbook
# # from openpyxl.utils.dataframe import dataframe_to_rows
# #
# # st.set_page_config(layout="wide")
# # st.title("🧾 GST Invoice to Formatted Data Extractor")
# #
# # uploaded_file = st.file_uploader("Upload GST Invoice Excel File", type=["xlsx"])
# #
# # if uploaded_file:
# #     df_raw = pd.read_excel(uploaded_file, sheet_name=0, header=None)
# #
# #     st.subheader("Raw Excel Preview")
# #     st.dataframe(df_raw.head(25))
# #
# #     header_keywords = ["Sr", "Description", "HSN", "Qty", "Rate", "Amount", "Taxable", "IGST", "Total"]
# #     header_row_index = None
# #     for i, row in df_raw.iterrows():
# #         match_count = sum(
# #             any(k.lower() in str(cell).lower() for cell in row if pd.notna(cell))
# #             for k in header_keywords
# #         )
# #         if match_count >= 4:
# #             header_row_index = i
# #             break
# #
# #     if header_row_index is not None:
# #         st.success(f"Header row found at row {header_row_index}")
# #
# #         df = pd.read_excel(uploaded_file, sheet_name=0, header=header_row_index)
# #         df = df.dropna(how='all')
# #         df.columns = [str(c).strip() for c in df.columns]
# #
# #         st.subheader("Detected Invoice Table")
# #         st.dataframe(df)
# #
# #         column_mapping = {
# #             "Sr. No.": "ITEM_SR_NO",
# #             "Description": "GOODS_DESC1",
# #             "Product": "GOODS_DESC1",
# #             "HSN": "RITC",
# #             "Code": "RITC",
# #             "Qty.": "QTY_NOS",
# #             "Unnamed: 4": "QTY_NOS",
# #             "in USD": "RATE_VALUE",
# #             "in USD.1": "TOTAL_VAL_FC",
# #             "Taxable value.1": "TAXABLE_VALUE",
# #             "IGST": "IGST_RATE",
# #             "In Set": "UNIT_OF_RATE",
# #             "In Set":"UNIT_QTY",
# #
# #
# #         }
# #
# #         mapped_df = pd.DataFrame()
# #         for col, new_col in column_mapping.items():
# #             if col in df.columns:
# #                 mapped_df[new_col] = df[col]
# #
# #         # Remove known footer rows
# #         footer_keywords = [
# #             "Total Invoice amount", "Bank Name", "Bank A/C", "Bank IFSC",
# #             "Authorised Signatory", "Certified", "Total Amount", "IGST",
# #             "Terma", "conditions", "GST ON Reverse", "Add :", "words", "In USD"
# #         ]
# #
# #         def is_footer_row(row):
# #             return any(
# #                 any(kw.lower() in str(cell).lower() for kw in footer_keywords)
# #                 for cell in row
# #             )
# #
# #         mapped_df = mapped_df[~mapped_df.apply(is_footer_row, axis=1)].reset_index(drop=True)
# #
# #         # Drop rows with 0/NaN/blank ITEM_SR_NO, RITC, or TAXABLE_VALUE
# #         for field in ["ITEM_SR_NO", "RITC", "TAXABLE_VALUE"]:
# #             if field in mapped_df.columns:
# #                 mapped_df = mapped_df[
# #                     mapped_df[field].notna() &
# #                     ~(mapped_df[field].astype(str).str.strip().isin(["", "0", "nan"]))
# #                 ]
# #
# #         mapped_df = mapped_df.reset_index(drop=True)
# #
# #         # Recalculate IGST_AMOUNT
# #         if "IGST_RATE" in mapped_df.columns and "TAXABLE_VALUE" in mapped_df.columns:
# #             try:
# #                 mapped_df["IGST_AMOUNT"] = pd.to_numeric(mapped_df["IGST_RATE"], errors='coerce') * pd.to_numeric(
# #                     mapped_df["TAXABLE_VALUE"], errors='coerce')
# #             except Exception as e:
# #                 st.warning("Couldn't compute IGST_AMOUNT: " + str(e))
# #
# #         st.subheader("📄 Cleaned Mapped Table with IGST Amount")
# #         st.dataframe(mapped_df)
# #
# #         # --- Final Output Format ---
# #         required_columns = [
# #             "INVOICE_SR_NO", "ITEM_SR_NO", "SCHEME_CODE", "RITC", "GOODS_DESC1", "GOODS_DESC2", "GOODS_DESC3",
# #             "QTY_NOS", "UNIT_QTY", "RATE_VALUE", "NO_OF_UNIT", "UNIT_OF_RATE", "PMV_AMT", "TOTAL_PMV",
# #             "ACCESSORIES_FLG", "CESS_FLG", "THIRD_PARTY_FLG", "AR4_FLG", "REWARD_FLG", "TOTAL_VAL_FC",
# #             "STR_FLG", "END_USE", "IGST_PAYMENT_STATUS", "TAXABLE_VALUE", "IGST_AMOUNT", "SWI_STO",
# #             "SWI_DOO", "SWI_EPT", "SWI_UQC", "SWI_QTY", "SWI_GCESS_AMT", "SWI_GCESS_CUR", "RODTEP_FLG",
# #             "SOURCE_STATE", "DBK_SRNO", "DBK_QUANTITY"
# #         ]
# #
# #         output_df = pd.DataFrame(columns=required_columns)
# #
# #         for col in output_df.columns:
# #             if col in mapped_df.columns:
# #                 output_df[col] = mapped_df[col]
# #
# #         # Fill required defaults
# #         output_df["INVOICE_SR_NO"] = 1
# #         output_df["GOODS_DESC2"] = "SIZE:"
# #         # output_df["UNIT_QTY"] = "PCS"
# #         output_df["NO_OF_UNIT"] = 1
# #         output_df["UNIT_OF_RATE"] = output_df["UNIT_QTY"]
# #         output_df["ACCESSORIES_FLG"] = "N"
# #         output_df["CESS_FLG"] = "N"
# #         output_df["THIRD_PARTY_FLG"] = "N"
# #         output_df["STR_FLG"] = "N"
# #         output_df["AR4_FLG"] = "N"
# #         output_df["REWARD_FLG"] = "Y"
# #         output_df["IGST_PAYMENT_STATUS"] = "P"
# #         output_df["SWI_STO"] = "06"
# #         output_df["SWI_DOO"] = "71"
# #         output_df["SWI_EPT"] = "ECTAAU"
# #         output_df["SWI_GCESS_CUR"] = "INR"
# #         output_df["SOURCE_STATE"] = "06"
# #
# #         # Re-assign IGST_AMOUNT if missed
# #         if "TAXABLE_VALUE" in mapped_df.columns and "IGST_RATE" in mapped_df.columns:
# #             try:
# #                 output_df["IGST_AMOUNT"] = pd.to_numeric(mapped_df["TAXABLE_VALUE"], errors='coerce') * pd.to_numeric(
# #                     mapped_df["IGST_RATE"], errors='coerce')
# #             except:
# #                 output_df["IGST_AMOUNT"] = ""
# #
# #         def compute_scheme(ritc):
# #             try:
# #                 return "60" if str(ritc).startswith("60") or str(ritc).startswith("61") or str(ritc).startswith("62") or str(ritc).startswith("63") else "19"
# #             except:
# #                 return "19"
# #
# #         output_df["SCHEME_CODE"] = output_df["RITC"].apply(compute_scheme)
# #
# #
# #         def compute_rodtep_flag(scheme_code):
# #             try:
# #                 return "N" if str(scheme_code).strip() == "60" else "Y"
# #             except:
# #                 return "Y"
# #
# #         output_df["RODTEP_FLG"] = output_df["SCHEME_CODE"].apply(compute_rodtep_flag)
# #
# #         # === Load RITC to SWI_UQC Mapping ===
# #         try:
# #             uqc_map_df = pd.read_excel("/Users/piyanshu/PycharmProjects/pdftoexcel/SQC_1.xlsx")
# #             uqc_map_df['RITC'] = uqc_map_df['RITC'].astype(str).str.strip()
# #             ritc_to_uqc_dict = dict(zip(uqc_map_df['RITC'], uqc_map_df['SWI_UQC']))
# #
# #             # Map SWI_UQC based on RITC
# #             output_df['SWI_UQC'] = output_df['RITC'].astype(str).str.strip().map(ritc_to_uqc_dict).fillna("NOS")
# #         except Exception as e:
# #             st.error(f"⚠️ Failed to load UQC mapping: {e}")
# #             output_df['SWI_UQC'] = "NOS"  # default fallback
# #
# #         # Excel Output
# #         output = io.BytesIO()
# #         wb = Workbook()
# #         ws = wb.active
# #         ws.title = "FormattedInvoice"
# #         for r in dataframe_to_rows(output_df, index=False, header=True):
# #             ws.append(r)
# #
# #         wb.save(output)
# #         output.seek(0)
# #
# #         st.download_button(
# #             label="📥 Download Formatted Excel",
# #             data=output,
# #             file_name="formatted_invoice.xlsx",
# #             mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
# #         )
# #
# #     else:
# #         st.error("❌ Could not detect a valid header row with invoice data.")
# import streamlit as st
# import pandas as pd
# import io
# from openpyxl import Workbook
# from openpyxl.utils.dataframe import dataframe_to_rows
# import os
# from datetime import datetime

# st.set_page_config(layout="wide")
# st.title("🧾 GST Invoice to Formatted Data Extractor")

# uploaded_file = st.file_uploader("Upload GST Invoice Excel File", type=["xlsx"])

# if uploaded_file:
#     df_raw = pd.read_excel(uploaded_file, sheet_name=0, header=None)


#     # st.subheader("Raw Excel Preview")
#     # st.dataframe(df_raw.head(25))

#     header_keywords = ["Sr", "Description", "HSN", "Qty", "Rate", "Amount", "Taxable", "IGST", "Total"]
#     header_row_index = None
#     for i, row in df_raw.iterrows():
#         match_count = sum(
#             any(k.lower() in str(cell).lower() for cell in row if pd.notna(cell))
#             for k in header_keywords
#         )
#         if match_count >= 4:
#             header_row_index = i
#             break

#     if header_row_index is not None:
#         st.success(f"Header row found at row {header_row_index}")

#         df = pd.read_excel(uploaded_file, sheet_name=0, header=header_row_index)
#         df = df.dropna(how='all')
#         df.columns = [str(c).strip() for c in df.columns]

#         st.subheader("Detected Invoice Table")
#         st.dataframe(df)

#         column_mapping = {
#             "Sr. No.": "ITEM_SR_NO",
#             "Description": "GOODS_DESC1",
#             "Product": "GOODS_DESC1",
#             "HSN": "RITC",
#             "Code": "RITC",
#             "Qty.": "QTY_NOS",
#             "Unnamed: 4": "QTY_NOS",
#             "in USD": "RATE_VALUE",
#             "in USD.1": "TOTAL_VAL_FC",
#             "Taxable value.1": "TAXABLE_VALUE",
#             "IGST": "IGST_RATE",
#             "In Set": "UNIT_OF_RATE",
#             "In Set":"UNIT_QTY",
#         }

#         mapped_df = pd.DataFrame()
#         for col, new_col in column_mapping.items():
#             if col in df.columns:
#                 mapped_df[new_col] = df[col]

#         # Remove known footer rows
#         footer_keywords = [
#             "Total Invoice amount", "Bank Name", "Bank A/C", "Bank IFSC",
#             "Authorised Signatory", "Certified", "Total Amount", "IGST",
#             "Terma", "conditions", "GST ON Reverse", "Add :", "words", "In USD"
#         ]

#         def is_footer_row(row):
#             return any(
#                 any(kw.lower() in str(cell).lower() for kw in footer_keywords)
#                 for cell in row
#             )

#         mapped_df = mapped_df[~mapped_df.apply(is_footer_row, axis=1)].reset_index(drop=True)

#         # Drop rows with 0/NaN/blank ITEM_SR_NO, RITC, or TAXABLE_VALUE
#         for field in ["ITEM_SR_NO", "RITC", "TAXABLE_VALUE"]:
#             if field in mapped_df.columns:
#                 mapped_df = mapped_df[
#                     mapped_df[field].notna() &
#                     ~(mapped_df[field].astype(str).str.strip().isin(["", "0", "nan"]))
#                 ]

#         mapped_df = mapped_df.reset_index(drop=True)

#         # Recalculate IGST_AMOUNT
#         if "IGST_RATE" in mapped_df.columns and "TAXABLE_VALUE" in mapped_df.columns:
#             try:
#                 mapped_df["IGST_AMOUNT"] = pd.to_numeric(mapped_df["IGST_RATE"], errors='coerce') * pd.to_numeric(
#                     mapped_df["TAXABLE_VALUE"], errors='coerce')
#             except Exception as e:
#                 st.warning("Couldn't compute IGST_AMOUNT: " + str(e))

#         st.subheader("📄 Cleaned Mapped Table with IGST Amount")
#         st.dataframe(mapped_df)

#         # --- Final Output Format ---
#         required_columns = [
#             "INVOICE_SR_NO", "ITEM_SR_NO", "SCHEME_CODE", "RITC", "GOODS_DESC1", "GOODS_DESC2", "GOODS_DESC3",
#             "QTY_NOS", "UNIT_QTY", "RATE_VALUE", "NO_OF_UNIT", "UNIT_OF_RATE", "PMV_AMT", "TOTAL_PMV",
#             "ACCESSORIES_FLG", "CESS_FLG", "THIRD_PARTY_FLG", "AR4_FLG", "REWARD_FLG", "TOTAL_VAL_FC",
#             "STR_FLG", "END_USE", "IGST_PAYMENT_STATUS", "TAXABLE_VALUE", "IGST_AMOUNT", "SWI_STO",
#             "SWI_DOO", "SWI_EPT", "SWI_UQC", "SWI_QTY", "SWI_GCESS_AMT", "SWI_GCESS_CUR", "RODTEP_FLG",
#             "SOURCE_STATE", "DBK_SRNO", "DBK_QUANTITY"
#         ]

#         output_df = pd.DataFrame(columns=required_columns)

#         for col in output_df.columns:
#             if col in mapped_df.columns:
#                 output_df[col] = mapped_df[col]

#         # Fill required defaults
#         output_df["INVOICE_SR_NO"] = 1
#         output_df["GOODS_DESC2"] = "SIZE:"
#         output_df["NO_OF_UNIT"] = 1
#         output_df["UNIT_OF_RATE"] = output_df["UNIT_QTY"]
#         output_df["ACCESSORIES_FLG"] = "N"
#         output_df["CESS_FLG"] = "N"
#         output_df["THIRD_PARTY_FLG"] = "N"
#         output_df["STR_FLG"] = "N"
#         output_df["AR4_FLG"] = "N"
#         output_df["REWARD_FLG"] = "Y"
#         output_df["IGST_PAYMENT_STATUS"] = "P"
#         output_df["SWI_STO"] = "06"
#         output_df["SWI_DOO"] = "71"
#         output_df["SWI_EPT"] = "ECTAAU"
#         output_df["SWI_GCESS_CUR"] = "INR"
#         output_df["SOURCE_STATE"] = "06"

#         # Re-assign IGST_AMOUNT if missed
#         if "TAXABLE_VALUE" in mapped_df.columns and "IGST_RATE" in mapped_df.columns:
#             try:
#                 output_df["IGST_AMOUNT"] = pd.to_numeric(mapped_df["TAXABLE_VALUE"], errors='coerce') * pd.to_numeric(
#                     mapped_df["IGST_RATE"], errors='coerce')
#             except:
#                 output_df["IGST_AMOUNT"] = ""

#         def compute_scheme(ritc):
#             try:
#                 return "60" if str(ritc).startswith(("60", "61", "62", "63")) else "19"
#             except:
#                 return "19"

#         output_df["SCHEME_CODE"] = output_df["RITC"].apply(compute_scheme)

#         def compute_rodtep_flag(scheme_code):
#             try:
#                 return "N" if str(scheme_code).strip() == "60" else "Y"
#             except:
#                 return "Y"

#         output_df["RODTEP_FLG"] = output_df["SCHEME_CODE"].apply(compute_rodtep_flag)

#         # === Load RITC to SWI_UQC Mapping ===
#         try:
#             uqc_map_df = pd.read_excel("/Users/piyanshu/PycharmProjects/pdftoexcel/SQC_1.xlsx")
#             uqc_map_df['RITC'] = uqc_map_df['RITC'].astype(str).str.strip()
#             ritc_to_uqc_dict = dict(zip(uqc_map_df['RITC'], uqc_map_df['SWI_UQC']))

#             # Map SWI_UQC based on RITC
#             output_df['SWI_UQC'] = output_df['RITC'].astype(str).str.strip().map(ritc_to_uqc_dict).fillna("NOS")
#         except Exception as e:
#             st.error(f"⚠️ Failed to load UQC mapping: {e}")
#             output_df['SWI_UQC'] = "NOS"  # default fallback

#         # Excel Output
#         output = io.BytesIO()
#         wb = Workbook()
#         ws = wb.active
#         ws.title = "FormattedInvoice"
#         for r in dataframe_to_rows(output_df, index=False, header=True):
#             ws.append(r)

#         wb.save(output)
#         output.seek(0)

#         # === 🖴 Save Automatically to Disk ===
#         SAVE_DIRECTORY = "/Users/piyanshu/Downloads"  # 👈 change this path if needed
#         timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
#         save_path = os.path.join(SAVE_DIRECTORY, f"formatted_invoice_{timestamp}.xlsx")

#         try:
#             with open(save_path, "wb") as f:
#                 f.write(output.getbuffer())
#             st.success(f"✅ File also saved automatically to: {save_path}")
#         except Exception as e:
#             st.error(f"❌ Failed to save to disk: {e}")

#         # === 📥 Streamlit Download Button ===
#         st.download_button(
#             label="📥 Download Formatted Excel",
#             data=output,
#             file_name="formatted_invoice.xlsx",
#             mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
#         )

#     else:
#         st.error("❌ Could not detect a valid header row with invoice data.")


import streamlit as st
import pandas as pd
import io
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import os
from datetime import datetime

# === Constants ===
COUNTRY_CODE_PATH = "/Users/piyanshu/PycharmProjects/pdftoexcel/Country Code.xlsx"  # 👈 change if needed
UQC_MAPPING_PATH = "/Users/piyanshu/PycharmProjects/pdftoexcel/SQC_1.xlsx"          # 👈 change if needed
SAVE_DIRECTORY = "/Users/piyanshu/Downloads"  # 👈 your preferred save location

# === App Config ===
# === App Config ===
st.set_page_config(layout="wide")
st.title("🧾 GST Invoice to Formatted Data Extractor")

# Dropdown selection for company
company_options = ["Bhajan", "Gupta International","PACHRANGA FOODS"]
selected_company = st.selectbox("Select Company", company_options)

if selected_company == "Gupta International":
    uploaded_file = st.file_uploader("Upload GST Invoice Excel File", type=["xlsx"])
    dollar_price = st.number_input("💵 Enter Dollar Price (Exchange Rate)", min_value=0.0, value=83.0, step=0.1)
    st.info(f"🧮 TAXABLE_VALUE will be calculated as TOTAL_VAL_FC × {dollar_price}")

    if uploaded_file:
        # Existing code for Gupta International
        df_raw = pd.read_excel(uploaded_file, sheet_name=0, header=None)

        # Extract company name and invoice number
        company_name = ""
        invoice_number = ""

        for row in df_raw[0].dropna().head(20):
            row_str = str(row).strip()
            if not company_name and any(
                    x in row_str.upper() for x in ["INTERNATIONAL", "LIMITED", "LTD", "PVT", "LLP"]):
                company_name = row_str.split(",")[0].strip()
            if "invoice no" in row_str.lower():
                try:
                    invoice_number = row_str.split("/")[-1].strip()
                except:
                    invoice_number = "XXXX"

        clean_co_name = company_name.upper().replace(" ", "_").replace(",", "")
        file_name = f"{clean_co_name}_{invoice_number}.xlsx"

    # === Load Country Mapping ===
    try:
        country_df = pd.read_excel(COUNTRY_CODE_PATH)
        country_df['Country'] = country_df['Country'].astype(str).str.strip().str.upper()
        country_to_code = dict(zip(country_df['Country'], country_df['Code']))
        country_to_code['ANYTHING ELSE'] = 'NCPTI'
    except Exception as e:
        st.error(f"❌ Failed to load country code mapping: {e}")
        country_to_code = {}

    if uploaded_file:
        df_raw = pd.read_excel(uploaded_file, sheet_name=0, header=None)
        # st.subheader("Raw Excel Preview")
        # st.dataframe(df_raw.head(25))

        # === Extract Country ===
        detected_country = None
        for i in range(min(40, len(df_raw))):
            row = df_raw.iloc[i]
            for j in range(len(row) - 1):
                if isinstance(row[j], str) and 'country' in row[j].lower():
                    detected_country = str(row[j + 1]).strip().upper()
                    break
            if detected_country:
                break
        st.success(f"🌍 Detected Country: {detected_country if detected_country else 'NOT FOUND'}")

        # === Get SWI_EPT Code ===
        swi_ept_code = country_to_code.get(detected_country, country_to_code.get('ANYTHING ELSE', 'NCPTI'))
        st.info(f"📄 Setting SWI_EPT to: {swi_ept_code}")

        # === Header Detection ===
        header_keywords = ["Sr", "Description", "HSN", "Qty", "Rate", "Amount", "Taxable", "IGST", "Total"]
        header_row_index = None
        for i, row in df_raw.iterrows():
            match_count = sum(
                any(k.lower() in str(cell).lower() for cell in row if pd.notna(cell)) for k in header_keywords)
            if match_count >= 4:
                header_row_index = i
                break

        if header_row_index is not None:
            df = pd.read_excel(uploaded_file, sheet_name=0, header=header_row_index)
            df = df.dropna(how='all')
            df.columns = [str(c).strip() for c in df.columns]
            st.subheader("Detected Invoice Table")
            st.dataframe(df)

            # === Column Mapping ===
            column_mapping = {
                "Sr. No.": "ITEM_SR_NO", "Description": "GOODS_DESC1", "Product": "GOODS_DESC1",
                "HSN": "RITC", "Code": "RITC", "Qty.": "QTY_NOS", "Unnamed: 4": "QTY_NOS",
                "in USD": "RATE_VALUE", "in USD.1": "TOTAL_VAL_FC",
                "Taxable value.1": "TAXABLE_VALUE", "IGST": "IGST_RATE",
                "In Set": "UNIT_QTY"
            }

            mapped_df = pd.DataFrame()
            for col, new_col in column_mapping.items():
                if col in df.columns:
                    mapped_df[new_col] = df[col]


            # === Standardize UNIT_QTY values ===
            def standardize_unit_qty(value):
                if isinstance(value, str) and "pcs" in value.lower():
                    return "PCS"
                return value


            if 'UNIT_QTY' in mapped_df.columns:
                mapped_df['UNIT_QTY'] = mapped_df['UNIT_QTY'].apply(standardize_unit_qty)

            # === Footer Cleanup ===
            footer_keywords = ["Total Invoice amount", "Bank Name", "Bank A/C", "Certified", "Total Amount",
                               "GST ON Reverse"]


            def is_footer_row(row):
                return any(any(kw.lower() in str(cell).lower() for kw in footer_keywords) for cell in row)


            mapped_df = mapped_df[~mapped_df.apply(is_footer_row, axis=1)].reset_index(drop=True)

            # === Filter ===
            for field in ["ITEM_SR_NO", "RITC"]:
                if field in mapped_df.columns:
                    mapped_df = mapped_df[
                        mapped_df[field].notna() & ~(mapped_df[field].astype(str).str.strip().isin(["", "0", "nan"]))]

            if "TOTAL_VAL_FC" in mapped_df.columns:
                mapped_df["TAXABLE_VALUE"] = pd.to_numeric(mapped_df["TOTAL_VAL_FC"], errors='coerce') * dollar_price

            if "IGST_RATE" in mapped_df.columns:
                mapped_df["IGST_AMOUNT"] = pd.to_numeric(mapped_df["IGST_RATE"], errors='coerce') * pd.to_numeric(
                    mapped_df["TAXABLE_VALUE"], errors='coerce')

            # === Output ===
            st.subheader("📄 Cleaned Mapped Table with IGST Amount")
            st.dataframe(mapped_df)

            # === Final Format ===
            required_columns = [
                "INVOICE_SR_NO", "ITEM_SR_NO", "SCHEME_CODE", "RITC", "GOODS_DESC1", "GOODS_DESC2", "GOODS_DESC3",
                "QTY_NOS", "UNIT_QTY", "RATE_VALUE", "NO_OF_UNIT", "UNIT_OF_RATE", "PMV_AMT", "TOTAL_PMV",
                "ACCESSORIES_FLG", "CESS_FLG", "THIRD_PARTY_FLG", "AR4_FLG", "REWARD_FLG", "TOTAL_VAL_FC",
                "STR_FLG", "END_USE", "IGST_PAYMENT_STATUS", "TAXABLE_VALUE", "IGST_AMOUNT", "SWI_STO",
                "SWI_DOO", "SWI_EPT", "SWI_UQC", "SWI_QTY", "SWI_GCESS_AMT", "SWI_GCESS_CUR", "RODTEP_FLG",
                "SOURCE_STATE", "DBK_SRNO", "DBK_QUANTITY"
            ]
            output_df = pd.DataFrame(columns=required_columns)
            for col in output_df.columns:
                if col in mapped_df.columns:
                    output_df[col] = mapped_df[col]

            output_df["INVOICE_SR_NO"] = 1
            output_df["GOODS_DESC2"] = "SIZE:"
            output_df["NO_OF_UNIT"] = 1
            output_df["UNIT_OF_RATE"] = output_df["UNIT_QTY"]
            output_df["ACCESSORIES_FLG"] = "N"
            output_df["CESS_FLG"] = "N"
            output_df["THIRD_PARTY_FLG"] = "N"
            output_df["STR_FLG"] = "N"
            output_df["AR4_FLG"] = "N"
            output_df["REWARD_FLG"] = "Y"
            output_df["IGST_PAYMENT_STATUS"] = "P"
            output_df["SWI_STO"] = "06"
            output_df["SWI_DOO"] = "71"
            output_df["SWI_EPT"] = swi_ept_code
            output_df["SWI_GCESS_CUR"] = "INR"
            output_df["SOURCE_STATE"] = "06"

            # SCHEME & RODTEP
            output_df["SCHEME_CODE"] = output_df["RITC"].astype(str).apply(
                lambda x: "60" if x.startswith(("60", "61", "62", "63")) else "19")
            output_df["RODTEP_FLG"] = output_df["SCHEME_CODE"].apply(lambda x: "N" if x == "60" else "Y")

            # === Load UQC Mapping ===
            try:
                uqc_map_df = pd.read_excel(UQC_MAPPING_PATH)
                uqc_map_df['RITC'] = uqc_map_df['RITC'].astype(str).str.strip()
                ritc_to_uqc = dict(zip(uqc_map_df['RITC'], uqc_map_df['SWI_UQC']))
                output_df['SWI_UQC'] = output_df['RITC'].astype(str).str.strip().map(ritc_to_uqc).fillna("NOS")
            except Exception as e:
                st.error(f"⚠️ Failed to load UQC mapping: {e}")
                output_df['SWI_UQC'] = "NOS"

            # === Excel Output ===
            output = io.BytesIO()
            wb = Workbook()
            ws = wb.active
            ws.title = "FormattedInvoice"
            for r in dataframe_to_rows(output_df, index=False, header=True):
                ws.append(r)
            wb.save(output)
            output.seek(0)

            # Save and Download
            if st.button("📁 Add to Folder"):
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                save_path = os.path.join(SAVE_DIRECTORY, f"formatted_invoice_{timestamp}.xlsx")
                try:
                    with open(save_path, "wb") as f:
                        f.write(output.getbuffer())
                    st.success(f"✅ File saved to: {save_path}")
                except Exception as e:
                    st.error(f"❌ Failed to save file: {e}")

            st.download_button("📥 Download Formatted Excel", output, "formatted_invoice.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            st.error("❌ Could not detect header row.")



def process_bhajan_to_export_format(df):
    """
    Process Bhajan invoices to match the to_export.csv format.
    Creates separate rows for sub-items with different rates.

    Args:
        df (pd.DataFrame): The DataFrame after initial header detection and cleaning.

    Returns:
        pd.DataFrame: A simplified DataFrame matching to_export format.
    """
    import pandas as pd

    # Make a copy to work on
    df_copy = df.copy()

    # --- 1. Initial Cleaning: Remove footers and non-item rows ---
    df_copy['Sr. No.'] = pd.to_numeric(df_copy['Sr. No.'], errors='coerce')

    # --- 2. Grouping Logic: Identify all rows belonging to one original item ---
    df_copy['Group_ID'] = df_copy['Sr. No.'].ffill()
    df_copy.dropna(subset=['Group_ID'], inplace=True)

    # --- 3. Process Groups and Create Records for Each Sub-item with Rate ---
    final_records = []
    output_sr_no = 1  # Counter for the output Sr. No.

    for _, group in df_copy.groupby('Group_ID'):
        # The first row of any group is the "parent" item
        parent_row = group.iloc[0]
        parent_description = str(parent_row['Product']).strip() if pd.notna(parent_row['Product']) else ""

        # Find all sub-items that have both Rate and 'in USD' values
        sub_items_with_rates = []

        for i in range(1, len(group)):
            sub_row = group.iloc[i]

            # Skip if no product description
            if pd.isna(sub_row['Product']) or not str(sub_row['Product']).strip():
                continue

            # Check if this sub-item has rate and USD values
            if pd.notna(sub_row.get('Rate')) and pd.notna(sub_row.get('in USD')):
                sub_items_with_rates.append({
                    'description': str(sub_row['Product']).strip(),
                    'rate': sub_row.get('Rate'),
                    'usd': sub_row.get('in USD')
                })

        # If no sub-items with rates found, check if parent has rate
        if not sub_items_with_rates:
            if pd.notna(parent_row.get('Rate')) and pd.notna(parent_row.get('in USD')):
                sub_items_with_rates.append({
                    'description': parent_description,
                    'rate': parent_row.get('Rate'),
                    'usd': parent_row.get('in USD')
                })

        # Create output records - one for each sub-item with rate
        for sub_item in sub_items_with_rates:
            # Combine parent description with sub-item description
            if parent_description and sub_item['description']:
                if parent_description != sub_item['description']:
                    consolidated_description = f"{parent_description} / {sub_item['description']}"
                else:
                    consolidated_description = parent_description
            else:
                consolidated_description = parent_description or sub_item['description']

            record = {
                'Sr. No.': output_sr_no,
                'Product': consolidated_description,
                'Rate': sub_item['rate'],
                'in USD': sub_item['usd'],
                'Taxable value': sub_item['usd']  # Same as 'in USD'
            }

            final_records.append(record)
            output_sr_no += 1

    if not final_records:
        return pd.DataFrame()

    # Create the final DataFrame with only the required columns
    final_df = pd.DataFrame(final_records)

    # Ensure we have the exact columns as in to_export.csv
    target_columns = ['Sr. No.', 'Product', 'Rate', 'in USD', 'Taxable value']

    # Reorder columns to match target format
    final_df = final_df[target_columns]

    return final_df.reset_index(drop=True)


def process_bhajan_to_export_format_alternative(df):
    """
    Alternative approach: Process based on your exact expected output pattern.
    This handles the specific case where we need to duplicate descriptions.
    """
    import pandas as pd

    # Make a copy to work on
    df_copy = df.copy()

    # Clean and group
    df_copy['Sr. No.'] = pd.to_numeric(df_copy['Sr. No.'], errors='coerce')
    df_copy['Group_ID'] = df_copy['Sr. No.'].ffill()
    df_copy.dropna(subset=['Group_ID'], inplace=True)

    final_records = []
    output_sr_no = 1

    for _, group in df_copy.groupby('Group_ID'):
        parent_row = group.iloc[0]
        parent_description = str(parent_row['Product']).strip() if pd.notna(parent_row['Product']) else ""

        # Collect all sub-items with rates
        rate_rows = []
        sub_descriptions = []

        for i in range(1, len(group)):
            sub_row = group.iloc[i]

            # Collect sub-descriptions
            if pd.notna(sub_row['Product']) and str(sub_row['Product']).strip():
                sub_descriptions.append(str(sub_row['Product']).strip())

            # Collect rows with rates
            if pd.notna(sub_row.get('Rate')) and pd.notna(sub_row.get('in USD')):
                rate_rows.append(sub_row)

        # Create the consolidated description (parent + all sub-descriptions)
        all_descriptions = [parent_description] + sub_descriptions
        consolidated_description = " / ".join(filter(None, all_descriptions))

        # Create one record for each rate found
        for rate_row in rate_rows:
            record = {
                'Sr. No.': output_sr_no,
                'Product': consolidated_description,
                'Rate': rate_row.get('Rate'),
                'in USD': rate_row.get('in USD'),
                'Taxable value': rate_row.get('in USD')
            }

            final_records.append(record)
            output_sr_no += 1

    if not final_records:
        return pd.DataFrame()

    final_df = pd.DataFrame(final_records)
    target_columns = ['Sr. No.', 'Product', 'Rate', 'in USD', 'Taxable value']
    final_df = final_df[target_columns]

    return final_df.reset_index(drop=True)

def process_bhajan_to_export_format_alternative(df):
    """
    Enhanced version that processes Bhajan invoices and maps additional fields like HSN, Qty, etc.
    Produces consolidated descriptions and all important numeric fields.
    """
    import pandas as pd

    df_copy = df.copy()

    # Clean and group
    df_copy['Sr. No.'] = pd.to_numeric(df_copy['Sr. No.'], errors='coerce')
    df_copy['Group_ID'] = df_copy['Sr. No.'].ffill()
    df_copy.dropna(subset=['Group_ID'], inplace=True)

    final_records = []
    output_sr_no = 1

    for _, group in df_copy.groupby('Group_ID'):
        parent_row = group.iloc[0]
        parent_description = str(parent_row.get('Product', '')).strip()

        # Gather sub-descriptions and rows with rates
        rate_rows = []
        sub_descriptions = []

        for i in range(1, len(group)):
            sub_row = group.iloc[i]

            desc = str(sub_row.get('Product', '')).strip()
            if desc:
                sub_descriptions.append(desc)

            if pd.notna(sub_row.get('Rate')) and pd.notna(sub_row.get('in USD')):
                rate_rows.append(sub_row)

        all_descriptions = [parent_description] + sub_descriptions
        consolidated_description = " / ".join(filter(None, all_descriptions))

        for rate_row in rate_rows:
            record = {
                'Sr. No.': output_sr_no,
                'Product': consolidated_description,
                'Unit_qty': rate_row.get('Unnamed: 5') or rate_row.get('In Pcs'),
                'in USD': rate_row.get('in USD'),
                'Taxable value(in USD)': rate_row.get('in USD'),  # Redundant but kept for clarity

                # Additional fields below
                'HSN Code': rate_row.get('HSN') or rate_row.get('HSN Code'),
                'Qty (in pcs)': rate_row.get('Unnamed: 4') or rate_row.get('Qty.'),
                'Qty (Sq./Mtrs)': rate_row.get('Sq.') or rate_row.get('Mtrs'),#in usd/  Pcs
                'Exchange Rate (in usd/  Pcs)': rate_row.get('Rate') or rate_row.get('in usd/  Pcs'),
                'Taxable Value (in Rs)': rate_row.get('Taxable value.1'),
                'IGST Rate (%)': rate_row.get('IGST'),
                'IGST Amount (in Rs)':  rate_row.get('Amount in Rs.'),
                'Total (in Rs)':  rate_row.get('Total in Rs.')
            }

            final_records.append(record)
            output_sr_no += 1

    if not final_records:
        return pd.DataFrame()

    final_df = pd.DataFrame(final_records)

    # Final column order
    target_columns = [
        'Sr. No.', 'Product', 'Unit_qty', 'in USD', 'Taxable value(in USD)',
        'HSN Code', 'Qty (in pcs)', 'Qty (Sq./Mtrs)', 'Exchange Rate (in usd/  Pcs)',
        'Taxable Value (in Rs)', 'IGST Rate (%)', 'IGST Amount (in Rs)', 'Total (in Rs)'
    ]

    final_df = final_df[target_columns]

    return final_df.reset_index(drop=True)


import pandas as pd
import io


def create_final_mapped_excel(export_df2, dollar_price, uqc_mapping_path=None, swi_ept_code="NCPTI"):
    # Optional UQC Mapping
    ritc_to_uqc = {}
    if uqc_mapping_path:
        try:
            uqc_map_df = pd.read_excel(uqc_mapping_path)
            uqc_map_df['RITC'] = uqc_map_df['RITC'].astype(str).str.strip()
            ritc_to_uqc = dict(zip(uqc_map_df['RITC'], uqc_map_df['SWI_UQC']))
        except Exception as e:
            st.error(f"⚠️ Failed to load UQC mapping: {e}")
            ritc_to_uqc = {}

    output_df = pd.DataFrame()
    output_df['INVOICE_SR_NO'] = [1] * len(export_df2)
    output_df['ITEM_SR_NO'] = export_df2['Sr. No.']
    output_df['SCHEME_CODE'] = export_df2['HSN Code'].astype(str).apply(
        lambda x: "60" if x.startswith(("60", "61", "62", "63")) else "19")
    output_df['RITC'] = export_df2['HSN Code']
    output_df['GOODS_DESC1'] = export_df2['Product']
    output_df['GOODS_DESC2'] = "SIZE:"
    output_df['GOODS_DESC3'] = ""
    output_df['QTY_NOS'] = export_df2['Qty (in pcs)']
    output_df['UNIT_QTY'] = export_df2['Unit_qty']
    output_df['RATE_VALUE'] = export_df2['Exchange Rate (in usd/  Pcs)']
    output_df['NO_OF_UNIT'] = 1
    output_df['UNIT_OF_RATE'] = output_df['UNIT_QTY']
    output_df['PMV_AMT'] = ""
    output_df['TOTAL_PMV'] = ""
    output_df['ACCESSORIES_FLG'] = "N"
    output_df['CESS_FLG'] = "N"
    output_df['THIRD_PARTY_FLG'] = "N"
    output_df['AR4_FLG'] = "N"
    output_df['REWARD_FLG'] = "Y"
    output_df['TOTAL_VAL_FC'] = export_df2['Taxable value(in USD)']
    output_df['STR_FLG'] = "N"
    output_df['END_USE'] = "GNX100"
    output_df['IGST_PAYMENT_STATUS'] = "P"
    output_df['TAXABLE_VALUE'] = pd.to_numeric(output_df["TOTAL_VAL_FC"], errors='coerce') * dollar_price
    output_df['IGST_AMOUNT'] = (pd.to_numeric(export_df2["IGST Rate (%)"], errors='coerce')) * output_df[
        "TAXABLE_VALUE"]
    output_df['SWI_STO'] = "06"
    output_df['SWI_DOO'] = "71"
    output_df['SWI_EPT'] = swi_ept_code
    output_df['SWI_UQC'] = output_df['RITC'].astype(str).str.strip().map(ritc_to_uqc).fillna("NOS")
    output_df['SWI_QTY'] = ""
    output_df['SWI_GCESS_AMT'] = 0
    output_df['SWI_GCESS_CUR'] = "INR"
    output_df['RODTEP_FLG'] = output_df['SCHEME_CODE'].apply(lambda x: "N" if x == "60" else "Y")
    output_df['SOURCE_STATE'] = "06"
    output_df['DBK_SRNO'] = ""
    output_df['DBK_QUANTITY'] = ""

    # Final Column Order as per your specification
    final_columns = [
        'INVOICE_SR_NO', 'ITEM_SR_NO', 'SCHEME_CODE', 'RITC', 'GOODS_DESC1', 'GOODS_DESC2', 'GOODS_DESC3',
        'QTY_NOS', 'UNIT_QTY', 'RATE_VALUE', 'NO_OF_UNIT', 'UNIT_OF_RATE', 'PMV_AMT', 'TOTAL_PMV',
        'ACCESSORIES_FLG', 'CESS_FLG', 'THIRD_PARTY_FLG', 'AR4_FLG', 'REWARD_FLG', 'TOTAL_VAL_FC',
        'STR_FLG', 'END_USE', 'IGST_PAYMENT_STATUS', 'TAXABLE_VALUE', 'IGST_AMOUNT',
        'SWI_STO', 'SWI_DOO', 'SWI_EPT', 'SWI_UQC', 'SWI_QTY', 'SWI_GCESS_AMT', 'SWI_GCESS_CUR',
        'RODTEP_FLG', 'SOURCE_STATE', 'DBK_SRNO', 'DBK_QUANTITY'
    ]

    return output_df[final_columns]


# 💾 Save to Excel buffer (for download)
def save_to_excel_buffer(df, sheet_name="FormattedInvoice"):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    output.seek(0)
    return output



# Updated Streamlit code section for Bhajan
if selected_company == "Bhajan":
    uploaded_file = st.file_uploader("Upload Bhajan Invoice", type=["xlsx", "xls"])

    if uploaded_file:
        st.success("File Uploaded Successfully!")

        df_raw = pd.read_excel(uploaded_file, sheet_name=0, header=None)

        # === Load Country Mapping ===
        try:
            country_df = pd.read_excel(COUNTRY_CODE_PATH)
            country_df['Country'] = country_df['Country'].astype(str).str.strip().str.upper()
            country_to_code = dict(zip(country_df['Country'], country_df['Code']))
            country_to_code['ANYTHING ELSE'] = 'NCPTI'
        except Exception as e:
            st.error(f"❌ Failed to load country code mapping: {e}")
            country_to_code = {}

        # === Detect Country in Raw Data ===
        detected_country = None
        for i in range(min(40, len(df_raw))):
            row = df_raw.iloc[i]
            for j in range(len(row) - 1):
                if isinstance(row[j], str) and 'country' in row[j].lower():
                    detected_country = str(row[j + 1]).strip().upper()
                    break
            if detected_country:
                break

        # st.success(f"🌍 Detected Country: {detected_country if detected_country else 'NOT FOUND'}")

        # === Get SWI_EPT Code ===
        swi_ept_code = country_to_code.get(detected_country, country_to_code.get('ANYTHING ELSE', 'NCPTI'))
        # st.info(f"📄 Setting SWI_EPT to: {swi_ept_code}")

        # --- Header Detection ---
        header_keywords = ["Sr. No.", "Product", "Description", "HSN", "Qty", "Rate", "Amount"]
        header_row_index = None

        for i, row in df_raw.iterrows():
            match_count = sum(
                any(k.lower() in str(cell).lower() for cell in row if pd.notna(cell)) for k in header_keywords)
            if match_count >= 3:
                header_row_index = i
                break

        # --- Process if Header is Found ---
        if header_row_index is not None:
            # st.info(f"Header row detected at index: {header_row_index}")
            df = pd.read_excel(uploaded_file, sheet_name=0, header=header_row_index)
            df.columns = [str(c).strip() for c in df.columns]

            # st.subheader("Detected Invoice Table (Before Cleaning)")
            # st.dataframe(df)

            # --- PRE-PROCESSING PIPELINE ---
            # 1. Remove completely blank rows
            df.dropna(how='all', inplace=True)
            df.reset_index(drop=True, inplace=True)

            trailing_keywords = ['less discount', 'discount', 'adjustment', 'deduction']

            rows_to_drop = []
            for idx, row in df.iterrows():
                row_str = ' '.join(row.dropna().astype(str)).lower()
                if any(k in row_str for k in trailing_keywords):
                    rows_to_drop.append(idx)

            if rows_to_drop:
                st.info(f"🧹 Removing {len(rows_to_drop)} trailing rows containing discount/adjustment summaries.")
                df.drop(rows_to_drop, inplace=True)
                df.reset_index(drop=True, inplace=True)

            # 2. Truncate the footer
            footer_keywords = [
                'round off', 'total invoice amount', 'in words', 'bank details',
                'bank name', 'bank a/c', 'certified that', 'authorised signatory',
                'common seal', 'total amount before tax', 'gst on reverse', 'add packing','Less Discount'
            ]
            footer_start_index = None
            for index, row in df.iterrows():
                row_as_string = ' '.join(row.dropna().astype(str)).lower()
                if any(keyword in row_as_string for keyword in footer_keywords):
                    footer_start_index = index
                    break

            if footer_start_index is not None:
                # st.info(f"Footer section detected starting at row {footer_start_index}. Truncating data...")
                df = df.iloc[:footer_start_index]
                df.reset_index(drop=True, inplace=True)

            # --- 3. Remove "Descriptive Header" Rows ---
            rows_to_drop = []
            key_value_columns = ['HSN', 'Qty.', 'Rate', 'in USD', 'Taxable value']

            for index, row in df.iterrows():
                is_sr_no_missing = pd.isna(row.get('Sr. No.'))
                existing_key_cols = [col for col in key_value_columns if col in df.columns]
                are_values_missing = row[existing_key_cols].isnull().all()

                if is_sr_no_missing and are_values_missing:
                    rows_to_drop.append(index)

            if rows_to_drop:
                st.info(f"Found and removed {len(rows_to_drop)} descriptive header rows that were not actual items.")
                df.drop(rows_to_drop, inplace=True)
                df.reset_index(drop=True, inplace=True)

            st.subheader("Detected Invoice Table (After All Cleaning)")
            st.dataframe(df)

            # --- CALL THE NEW EXPORT FORMAT FUNCTION ---


            st.subheader("✨ Export Format(Consolidated Descriptions)")
            with st.spinner("Converting to export format..."):
                export_df2 = process_bhajan_to_export_format_alternative(df)
            st.dataframe(export_df2)

            st.success("✅ Export format conversion complete!")

            # Let user choose which method to download


            chosen_df = export_df2

            # Add download button for the export format
            st.subheader("📦 Final Mapped Excel Format")
            dollar_price = st.number_input("💵 Enter Dollar Price (Exchange Rate)", min_value=0.0, value=83.0, step=0.1)


            if st.button("🔄 Generate Final Excel Table"):
                final_mapped_df = create_final_mapped_excel(export_df2, dollar_price,
                                                            uqc_mapping_path="/Users/piyanshu/PycharmProjects/pdftoexcel/SQC_1.xlsx")
                st.dataframe(final_mapped_df)

                final_excel = save_to_excel_buffer(final_mapped_df)
                st.download_button(
                    label="📥 Download Final Excel Table",
                    data=final_excel,
                    file_name="final_bhajan_invoice.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )



        else:
            st.error("Could not detect a standard header row.")

import pandas as pd


import pandas as pd
import streamlit as st
import io

def structure_pachranga_invoice(raw_df):
    """
    Extracts structured item data from Pachranga invoice DataFrame.
    Adapts to layout like:
    1 TO 98 | 98 | 6.86 | Peepa Biscuit Tin | 19053100 | 2 KG | 4 | 392 | 9.87 | 3869.04
    """

    # Define final column structure
    target_columns = [
        'CARTON NO.', 'NO OF CTN', 'CBM', 'Item Name', 'HSN CODE','CARTON TO',
        'SIZE', 'QTY', 'TOTAL PCS', 'RATE (USD)', 'TOTAL USD'
    ]

    # Find the start row: where 'CARTON NO.' header appears
    start_index = None
    for i, row in raw_df.iterrows():
        if any('CARTON NO' in str(cell).upper() for cell in row):
            start_index = i + 1
            break

    if start_index is None:
        st.error("❌ 'CARTON NO.' header not found in invoice.")
        return pd.DataFrame(columns=target_columns)

    data_rows = raw_df.iloc[start_index:].dropna(how='all')
    structured_data = []

    for _, row in data_rows.iterrows():
        values = row.dropna().tolist()

        # Skip totals or footer lines
        if any("SUPPLY MEANT" in str(v).upper() or "DISCOUNT" in str(v).upper() for v in values):
            continue

        # Expecting: CARTON NO. | NO OF CTN | CBM | Item Name | HSN CODE | SIZE | PCS/CTN | TOTAL PCS | RATE | TOTAL
        if len(values) >= 11:
            try:
                # Flatten carton range (e.g., '1 TO 98') to starting carton number
                carton_raw = str(values[0])
                if 'TO' in carton_raw.upper():
                    carton_no = carton_raw.split()[0]
                else:
                    carton_no = carton_raw

                structured_data.append({
                    'CARTON NO.': carton_no,
                    'CARTON TO': values[2],   # Skipping "TO"
                    'NO OF CTN': values[3],
                    'CBM': values[4],
                    'Item Name':values[5],
                    'HSN CODE': values[6],
                    'SIZE': values[7],
                    'QTY': values[8],
                    'TOTAL PCS': values[9],
                    'RATE (USD)': values[10],
                    'TOTAL USD': values[11]
                })
            except Exception as e:
                st.warning(f"⚠️ Skipping row due to error: {e}")

    df_structured = pd.DataFrame(structured_data, columns=target_columns)
    return df_structured


def to_excel_buffer(df, sheet_name="StructuredData"):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    output.seek(0)
    return output


# === Streamlit UI for Pachranga Foods ===
if selected_company == "PACHRANGA FOODS":
    uploaded_file = st.file_uploader("📤 Upload PACHRANGA FOODS Invoice", type=["xlsx", "xls"])

    if uploaded_file:
        st.success("📄 File Uploaded Successfully!")

        df_raw = pd.read_excel(uploaded_file, sheet_name=0, header=None)
        st.subheader("📄 Raw Invoice Preview")
        st.dataframe(df_raw)

        structured_df = structure_pachranga_invoice(df_raw)

        if not structured_df.empty:
            st.subheader("🧾 Structured Invoice Data")
            st.dataframe(structured_df)

            # Dollar Price input (optional)
            dollar_price = st.number_input("💵 Enter Dollar Price (Exchange Rate)", min_value=0.0, value=83.0, step=0.1)

            # Download button
            st.subheader("📥 Download Structured Excel")
            excel_buffer = to_excel_buffer(structured_df)
            st.download_button(
                label="⬇️ Download Structured Invoice",
                data=excel_buffer,
                file_name="pachranga_structured_invoice.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("⚠️ No valid structured data extracted.")

