# import pandas as pd
# import streamlit as st
#
# st.title("Excel Field Extractor")
#
# # Upload the Excel file
# uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")
#
# if uploaded_file is not None:
#     # Read the Excel file
#     df = pd.read_excel(uploaded_file, sheet_name=0)
#
#     # Display the DataFrame
#     st.write("Data from the Excel sheet:")
#     st.write(df)
#
#     # Extract fields based on serial numbers
#     serial_numbers = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10]  # Modify as needed
#
#     # Create a DataFrame with extracted fields
#     extracted_df = df[df['Sr. No.'].isin(serial_numbers)]
#     st.write("Extracted Fields:")
#     st.write(extracted_df)
#
# # Run the app with: streamlit run your_script.py
#
# import streamlit as st
# import pandas as pd
# import re
#
# st.set_page_config(layout="wide")
# st.title("GST Tax Invoice Field Extractor")
#
# uploaded_file = st.file_uploader("Upload Invoice Excel", type=["xlsx"])
#
# if uploaded_file:
#     # Read all sheets
#     xls = pd.ExcelFile(uploaded_file)
#     df = xls.parse(xls.sheet_names[0], header=None)
#
#     st.subheader("Raw Excel Preview")
#     st.dataframe(df)
#
#     # Step 1: Identify the header row
#     # Detect header row from the Excel sheet more flexibly
#     header_keywords = ["Sr", "Product", "Description", "HSN", "Qty", "Rate", "Amount", "Taxable", "IGST", "Total"]
#     header_row_index = None
#
#     for i, row in df.iterrows():
#         match_count = sum(
#             any(k.lower() in str(cell).lower() for cell in row if pd.notna(cell))
#             for k in header_keywords
#         )
#         if match_count >= 4:
#             header_row_index = i
#             break
#
#     if header_row_index is not None:
#         st.success(f"Header row found at row {header_row_index}")
#         # Extract data starting from header row
#         table_df = pd.read_excel(uploaded_file, header=header_row_index)
#         table_df = table_df.dropna(how='all')
#
#         st.subheader("Extracted Table")
#         st.dataframe(table_df)
#     else:
#         st.error("❌ Could not detect a valid header row.")
# import streamlit as st
# import pandas as pd
# import io
# from openpyxl import Workbook
# from openpyxl.utils.dataframe import dataframe_to_rows
#
# st.set_page_config(layout="wide")
# st.title("🧾 GST Invoice to Formatted Data Extractor")
#
# uploaded_file = st.file_uploader("Upload GST Invoice Excel File", type=["xlsx"])
#
# if uploaded_file:
#     df_raw = pd.read_excel(uploaded_file, sheet_name=0, header=None)
#
#     st.subheader("Raw Excel Preview")
#     st.dataframe(df_raw.head(25))
#
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
#
#     if header_row_index is not None:
#         st.success(f"Header row found at row {header_row_index}")
#
#         df = pd.read_excel(uploaded_file, sheet_name=0, header=header_row_index)
#         df = df.dropna(how='all')
#         df.columns = [str(c).strip() for c in df.columns]
#
#         st.subheader("Detected Invoice Table")
#         st.dataframe(df)
#
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
#
#
#         }
#
#         mapped_df = pd.DataFrame()
#         for col, new_col in column_mapping.items():
#             if col in df.columns:
#                 mapped_df[new_col] = df[col]
#
#         # Remove known footer rows
#         footer_keywords = [
#             "Total Invoice amount", "Bank Name", "Bank A/C", "Bank IFSC",
#             "Authorised Signatory", "Certified", "Total Amount", "IGST",
#             "Terma", "conditions", "GST ON Reverse", "Add :", "words", "In USD"
#         ]
#
#         def is_footer_row(row):
#             return any(
#                 any(kw.lower() in str(cell).lower() for kw in footer_keywords)
#                 for cell in row
#             )
#
#         mapped_df = mapped_df[~mapped_df.apply(is_footer_row, axis=1)].reset_index(drop=True)
#
#         # Drop rows with 0/NaN/blank ITEM_SR_NO, RITC, or TAXABLE_VALUE
#         for field in ["ITEM_SR_NO", "RITC", "TAXABLE_VALUE"]:
#             if field in mapped_df.columns:
#                 mapped_df = mapped_df[
#                     mapped_df[field].notna() &
#                     ~(mapped_df[field].astype(str).str.strip().isin(["", "0", "nan"]))
#                 ]
#
#         mapped_df = mapped_df.reset_index(drop=True)
#
#         # Recalculate IGST_AMOUNT
#         if "IGST_RATE" in mapped_df.columns and "TAXABLE_VALUE" in mapped_df.columns:
#             try:
#                 mapped_df["IGST_AMOUNT"] = pd.to_numeric(mapped_df["IGST_RATE"], errors='coerce') * pd.to_numeric(
#                     mapped_df["TAXABLE_VALUE"], errors='coerce')
#             except Exception as e:
#                 st.warning("Couldn't compute IGST_AMOUNT: " + str(e))
#
#         st.subheader("📄 Cleaned Mapped Table with IGST Amount")
#         st.dataframe(mapped_df)
#
#         # --- Final Output Format ---
#         required_columns = [
#             "INVOICE_SR_NO", "ITEM_SR_NO", "SCHEME_CODE", "RITC", "GOODS_DESC1", "GOODS_DESC2", "GOODS_DESC3",
#             "QTY_NOS", "UNIT_QTY", "RATE_VALUE", "NO_OF_UNIT", "UNIT_OF_RATE", "PMV_AMT", "TOTAL_PMV",
#             "ACCESSORIES_FLG", "CESS_FLG", "THIRD_PARTY_FLG", "AR4_FLG", "REWARD_FLG", "TOTAL_VAL_FC",
#             "STR_FLG", "END_USE", "IGST_PAYMENT_STATUS", "TAXABLE_VALUE", "IGST_AMOUNT", "SWI_STO",
#             "SWI_DOO", "SWI_EPT", "SWI_UQC", "SWI_QTY", "SWI_GCESS_AMT", "SWI_GCESS_CUR", "RODTEP_FLG",
#             "SOURCE_STATE", "DBK_SRNO", "DBK_QUANTITY"
#         ]
#
#         output_df = pd.DataFrame(columns=required_columns)
#
#         for col in output_df.columns:
#             if col in mapped_df.columns:
#                 output_df[col] = mapped_df[col]
#
#         # Fill required defaults
#         output_df["INVOICE_SR_NO"] = 1
#         output_df["GOODS_DESC2"] = "SIZE:"
#         # output_df["UNIT_QTY"] = "PCS"
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
#
#         # Re-assign IGST_AMOUNT if missed
#         if "TAXABLE_VALUE" in mapped_df.columns and "IGST_RATE" in mapped_df.columns:
#             try:
#                 output_df["IGST_AMOUNT"] = pd.to_numeric(mapped_df["TAXABLE_VALUE"], errors='coerce') * pd.to_numeric(
#                     mapped_df["IGST_RATE"], errors='coerce')
#             except:
#                 output_df["IGST_AMOUNT"] = ""
#
#         def compute_scheme(ritc):
#             try:
#                 return "60" if str(ritc).startswith("60") or str(ritc).startswith("61") or str(ritc).startswith("62") or str(ritc).startswith("63") else "19"
#             except:
#                 return "19"
#
#         output_df["SCHEME_CODE"] = output_df["RITC"].apply(compute_scheme)
#
#
#         def compute_rodtep_flag(scheme_code):
#             try:
#                 return "N" if str(scheme_code).strip() == "60" else "Y"
#             except:
#                 return "Y"
#
#         output_df["RODTEP_FLG"] = output_df["SCHEME_CODE"].apply(compute_rodtep_flag)
#
#         # === Load RITC to SWI_UQC Mapping ===
#         try:
#             uqc_map_df = pd.read_excel("/Users/piyanshu/PycharmProjects/pdftoexcel/SQC_1.xlsx")
#             uqc_map_df['RITC'] = uqc_map_df['RITC'].astype(str).str.strip()
#             ritc_to_uqc_dict = dict(zip(uqc_map_df['RITC'], uqc_map_df['SWI_UQC']))
#
#             # Map SWI_UQC based on RITC
#             output_df['SWI_UQC'] = output_df['RITC'].astype(str).str.strip().map(ritc_to_uqc_dict).fillna("NOS")
#         except Exception as e:
#             st.error(f"⚠️ Failed to load UQC mapping: {e}")
#             output_df['SWI_UQC'] = "NOS"  # default fallback
#
#         # Excel Output
#         output = io.BytesIO()
#         wb = Workbook()
#         ws = wb.active
#         ws.title = "FormattedInvoice"
#         for r in dataframe_to_rows(output_df, index=False, header=True):
#             ws.append(r)
#
#         wb.save(output)
#         output.seek(0)
#
#         st.download_button(
#             label="📥 Download Formatted Excel",
#             data=output,
#             file_name="formatted_invoice.xlsx",
#             mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
#         )
#
#     else:
#         st.error("❌ Could not detect a valid header row with invoice data.")
import streamlit as st
import pandas as pd
import io
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import os
from datetime import datetime

st.set_page_config(layout="wide")
st.title("🧾 GST Invoice to Formatted Data Extractor")

uploaded_file = st.file_uploader("Upload GST Invoice Excel File", type=["xlsx"])

if uploaded_file:
    df_raw = pd.read_excel(uploaded_file, sheet_name=0, header=None)


    # st.subheader("Raw Excel Preview")
    # st.dataframe(df_raw.head(25))

    header_keywords = ["Sr", "Description", "HSN", "Qty", "Rate", "Amount", "Taxable", "IGST", "Total"]
    header_row_index = None
    for i, row in df_raw.iterrows():
        match_count = sum(
            any(k.lower() in str(cell).lower() for cell in row if pd.notna(cell))
            for k in header_keywords
        )
        if match_count >= 4:
            header_row_index = i
            break

    if header_row_index is not None:
        st.success(f"Header row found at row {header_row_index}")

        df = pd.read_excel(uploaded_file, sheet_name=0, header=header_row_index)
        df = df.dropna(how='all')
        df.columns = [str(c).strip() for c in df.columns]

        st.subheader("Detected Invoice Table")
        st.dataframe(df)

        column_mapping = {
            "Sr. No.": "ITEM_SR_NO",
            "Description": "GOODS_DESC1",
            "Product": "GOODS_DESC1",
            "HSN": "RITC",
            "Code": "RITC",
            "Qty.": "QTY_NOS",
            "Unnamed: 4": "QTY_NOS",
            "in USD": "RATE_VALUE",
            "in USD.1": "TOTAL_VAL_FC",
            "Taxable value.1": "TAXABLE_VALUE",
            "IGST": "IGST_RATE",
            "In Set": "UNIT_OF_RATE",
            "In Set":"UNIT_QTY",
        }

        mapped_df = pd.DataFrame()
        for col, new_col in column_mapping.items():
            if col in df.columns:
                mapped_df[new_col] = df[col]

        # Remove known footer rows
        footer_keywords = [
            "Total Invoice amount", "Bank Name", "Bank A/C", "Bank IFSC",
            "Authorised Signatory", "Certified", "Total Amount", "IGST",
            "Terma", "conditions", "GST ON Reverse", "Add :", "words", "In USD"
        ]

        def is_footer_row(row):
            return any(
                any(kw.lower() in str(cell).lower() for kw in footer_keywords)
                for cell in row
            )

        mapped_df = mapped_df[~mapped_df.apply(is_footer_row, axis=1)].reset_index(drop=True)

        # Drop rows with 0/NaN/blank ITEM_SR_NO, RITC, or TAXABLE_VALUE
        for field in ["ITEM_SR_NO", "RITC", "TAXABLE_VALUE"]:
            if field in mapped_df.columns:
                mapped_df = mapped_df[
                    mapped_df[field].notna() &
                    ~(mapped_df[field].astype(str).str.strip().isin(["", "0", "nan"]))
                ]

        mapped_df = mapped_df.reset_index(drop=True)

        # Recalculate IGST_AMOUNT
        if "IGST_RATE" in mapped_df.columns and "TAXABLE_VALUE" in mapped_df.columns:
            try:
                mapped_df["IGST_AMOUNT"] = pd.to_numeric(mapped_df["IGST_RATE"], errors='coerce') * pd.to_numeric(
                    mapped_df["TAXABLE_VALUE"], errors='coerce')
            except Exception as e:
                st.warning("Couldn't compute IGST_AMOUNT: " + str(e))

        st.subheader("📄 Cleaned Mapped Table with IGST Amount")
        st.dataframe(mapped_df)

        # --- Final Output Format ---
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

        # Fill required defaults
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
        output_df["SWI_EPT"] = "ECTAAU"
        output_df["SWI_GCESS_CUR"] = "INR"
        output_df["SOURCE_STATE"] = "06"

        # Re-assign IGST_AMOUNT if missed
        if "TAXABLE_VALUE" in mapped_df.columns and "IGST_RATE" in mapped_df.columns:
            try:
                output_df["IGST_AMOUNT"] = pd.to_numeric(mapped_df["TAXABLE_VALUE"], errors='coerce') * pd.to_numeric(
                    mapped_df["IGST_RATE"], errors='coerce')
            except:
                output_df["IGST_AMOUNT"] = ""

        def compute_scheme(ritc):
            try:
                return "60" if str(ritc).startswith(("60", "61", "62", "63")) else "19"
            except:
                return "19"

        output_df["SCHEME_CODE"] = output_df["RITC"].apply(compute_scheme)

        def compute_rodtep_flag(scheme_code):
            try:
                return "N" if str(scheme_code).strip() == "60" else "Y"
            except:
                return "Y"

        output_df["RODTEP_FLG"] = output_df["SCHEME_CODE"].apply(compute_rodtep_flag)

        # === Load RITC to SWI_UQC Mapping ===
        try:
            uqc_map_df = pd.read_excel("/Users/piyanshu/PycharmProjects/pdftoexcel/SQC_1.xlsx")
            uqc_map_df['RITC'] = uqc_map_df['RITC'].astype(str).str.strip()
            ritc_to_uqc_dict = dict(zip(uqc_map_df['RITC'], uqc_map_df['SWI_UQC']))

            # Map SWI_UQC based on RITC
            output_df['SWI_UQC'] = output_df['RITC'].astype(str).str.strip().map(ritc_to_uqc_dict).fillna("NOS")
        except Exception as e:
            st.error(f"⚠️ Failed to load UQC mapping: {e}")
            output_df['SWI_UQC'] = "NOS"  # default fallback

        # Excel Output
        output = io.BytesIO()
        wb = Workbook()
        ws = wb.active
        ws.title = "FormattedInvoice"
        for r in dataframe_to_rows(output_df, index=False, header=True):
            ws.append(r)

        wb.save(output)
        output.seek(0)

        # === 🖴 Save Automatically to Disk ===
        SAVE_DIRECTORY = "/Users/piyanshu/Downloads"  # 👈 change this path if needed
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        save_path = os.path.join(SAVE_DIRECTORY, f"formatted_invoice_{timestamp}.xlsx")

        try:
            with open(save_path, "wb") as f:
                f.write(output.getbuffer())
            st.success(f"✅ File also saved automatically to: {save_path}")
        except Exception as e:
            st.error(f"❌ Failed to save to disk: {e}")

        # === 📥 Streamlit Download Button ===
        st.download_button(
            label="📥 Download Formatted Excel",
            data=output,
            file_name="formatted_invoice.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    else:
        st.error("❌ Could not detect a valid header row with invoice data.")
