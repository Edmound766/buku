import gc
import io
import re

import pandas as pd
import pdfplumber
import streamlit as st
import streamlit.components.v1 as components
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


def display_ad_banner():
    # Insert your AdSense or custom ad HTML/JS script here
    ad_code = """
    <div style="display:flex; justify-content:center; align-items:center; padding: 10px; background-color: #f9f9f9; border-radius: 8px;">
        <!-- Replace with your actual Ad Network Code snippet -->
        <script async src="https://pagead2.googlesyndication.com/pagead/js/adsbygoogle.js"></script>
        <ins class="adsbygoogle"
             style="display:block"
             data-ad-client="ca-pub-XXXXXXXXXXXXXXXX"
             data-ad-slot="1234567890"
             data-ad-format="auto"
             data-full-width-responsive="true"></ins>
        <script>
             (adsbygoogle = window.adsbygoogle || []).push({});
        </script>
    </div>
    """
    components.html(ad_code, height=100)


st.set_page_config(page_title="Buku - M-Pesa to Excel", page_icon="💳", layout="wide")

st.title("💳 Buku")
display_ad_banner()

st.subheader("Convert your M-Pesa statements into clean, professional Excel sheets instantly.")
st.markdown("---")
st.write("🔒 **Privacy First:** Buku processes your PDF statements 100% locally on your machine.")

# Initialize session state trackers
if "df" not in st.session_state:
    st.session_state.df = None
if "excel_buffer" not in st.session_state:
    st.session_state.excel_buffer = None

with st.container():
    col1, col2 = st.columns([2, 1])
    with col1:
        uploaded_file = st.file_uploader("Drop your Safaricom M-Pesa PDF statement here", type=["pdf"])
    with col2:
        password = st.text_input("PDF Password (ID / Passport)", type="password")

def export_with_totals(cleaned_df: pd.DataFrame):
    """
    Takes your parsed Pandas dataframe, appends formula-driven total rows,
    applies professional typography layout styles, and outputs raw bytes
    safe for Streamlit's st.download_button.
    """
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        cleaned_df.to_excel(writer, index=False, sheet_name='M-Pesa Statement')
        workbook = writer.book
        worksheet = writer.sheets['M-Pesa Statement']
        
        start_row = 2 
        end_row = len(cleaned_df) + 1
        totals_row = end_row + 2 
        
        # Identify columns dynamically
        # Map amount columns based on dataframe layout
        paid_in_col_idx = cleaned_df.columns.get_loc("Paid In (Cr)") + 1 if "Paid In (Cr)" in cleaned_df.columns else 5
        withdrawn_col_idx = cleaned_df.columns.get_loc("Withdrawn (Dr)") + 1 if "Withdrawn (Dr)" in cleaned_df.columns else 6
        
        paid_in_letter = get_column_letter(paid_in_col_idx)
        withdrawn_letter = get_column_letter(withdrawn_col_idx)

        # Inject Native Excel Formula Blocks
        worksheet[f'C{totals_row}'] = "TOTAL SUMMARY:"
        worksheet[f'{paid_in_letter}{totals_row}'] = f"=SUM({paid_in_letter}{start_row}:{paid_in_letter}{end_row})"
        worksheet[f'{withdrawn_letter}{totals_row}'] = f"=SUM({withdrawn_letter}{start_row}:{withdrawn_letter}{end_row})"

        # Styling Configuration: Executive Layout Polish
        header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        totals_font = Font(name='Calibri', size=11, bold=True, color='000000')
        subtle_gray_fill = PatternFill(start_color='F2F2F2', end_color='F2F2F2', fill_type='solid')
        border_thin = Border(left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'), top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9'))
        align_center = Alignment(horizontal="center", vertical="center")
        align_right = Alignment(horizontal="right", vertical="center")
        align_left = Alignment(horizontal="left", vertical="center")

        # Format Header
        for col_num in range(1, len(cleaned_df.columns) + 1):
            cell = worksheet.cell(row=1, column=col_num)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = align_center

        # Format Data Rows & Auto-fit Columns
        for col_idx, col_name in enumerate(cleaned_df.columns, 1):
            max_len = len(str(col_name))
            col_letter = get_column_letter(col_idx)

            for row_idx in range(2, len(cleaned_df) + 2):
                cell = worksheet.cell(row=row_idx, column=col_idx)
                cell.border = border_thin
                cell_val = cell.value

                if cell_val is not None:
                    max_len = max(max_len, len(str(cell_val)))

                if col_name in ["Paid In (Cr)", "Withdrawn (Dr)", "Balance"]:
                    cell.alignment = align_right
                    if isinstance(cell_val, (int, float)):
                        cell.number_format = '#,##0.00'
                elif col_name in ["Receipt No.", "Completion Time", "Status"]:
                    cell.alignment = align_center
                else:
                    cell.alignment = align_left

            worksheet.column_dimensions[col_letter].width = max(max_len + 4, 12)

        # Apply formatting elements to the totals row sequence
        for col in range(1, worksheet.max_column + 1):
            cell = worksheet.cell(row=totals_row, column=col)
            cell.font = totals_font
            cell.fill = subtle_gray_fill
            cell.border = border_thin
            if col in [paid_in_col_idx, withdrawn_col_idx]:
                cell.alignment = align_right
                cell.number_format = '#,##0.00'

    processed_bytes = output.getvalue()
    output.close()
    gc.collect()
    return processed_bytes

# Process file only when uploaded and button clicked
if uploaded_file and password and st.button("🚀 Generate Clean Ledger", type="primary", use_container_width=True):
    with st.spinner("Buku is securely decrypting and mapping your transactions..."):
        try:
            with pdfplumber.open(uploaded_file, password=password) as pdf:
                all_rows = []
                tx_pattern = re.compile(r'^([A-Z0-9]{10})\s+(\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2})\s+(.+)$')

                all_lines = []
                for page in pdf.pages:
                    text = page.extract_text()
                    if text:
                        for line in text.split("\n"):
                            all_lines.append(line.strip())

                i = 0
                while i < len(all_lines):
                    line = all_lines[i]
                    match = tx_pattern.match(line)

                    if match:
                        receipt_no = match.group(1)
                        completion_time = match.group(2)
                        remainder = match.group(3)

                        extra_details = []
                        j = i + 1
                        while j < len(all_lines):
                            next_line = all_lines[j]
                            if tx_pattern.match(next_line) or "Page " in next_line or "M-PESA Statement" in next_line or "Disclaimer:" in next_line:
                                break
                            if next_line:
                                extra_details.append(next_line)
                            j += 1

                        parts = remainder.rsplit(maxsplit=4)

                        if len(parts) >= 3:
                            status = parts[-3]
                            base_details = " ".join(parts[:-3])
                            details = base_details + " " + " ".join(extra_details) if extra_details else base_details

                            if "Completed" in details and status != "Completed":
                                details = details.replace("Completed", "").strip()
                                status = "Completed"

                            raw_balance = parts[-1].replace('"', '').replace("'", "").replace(",", "").strip()
                            raw_amount = parts[-2].replace('"', '').replace("'", "").replace(",", "").strip()

                            try:
                                balance_num = float(raw_balance)
                            except ValueError:
                                balance_num = raw_balance

                            try:
                                if "-" in raw_amount:
                                    amount_num = float(raw_amount.replace("-", "").strip())
                                    withdrawn = amount_num
                                    paid_in = None
                                else:
                                    amount_num = float(raw_amount)
                                    if "charge" in details.lower() or "withdrawal" in details.lower():
                                        withdrawn = amount_num
                                        paid_in = None
                                    else:
                                        withdrawn = None
                                        paid_in = amount_num
                            except ValueError:
                                if "-" in raw_amount:
                                    withdrawn = raw_amount
                                    paid_in = None
                                else:
                                    withdrawn = None
                                    paid_in = raw_amount

                            all_rows.append({
                                "Receipt No.": receipt_no,
                                "Completion Time": completion_time,
                                "Details": details,
                                "Status": status,
                                "Paid In (Cr)": paid_in,
                                "Withdrawn (Dr)": withdrawn,
                                "Balance": balance_num
                            })
                        i = j
                    else:
                        i += 1

            if all_rows:
                st.session_state.df = pd.DataFrame(all_rows)
                st.session_state.excel_buffer = export_with_totals(st.session_state.df)
                st.rerun()
            else:
                st.error("Buku couldn't find any transaction entries.")
        except Exception as exc:
            st.error(f"Decryption or parsing failed. Error: {exc!s}")

# Presentation & Interactive Toolbar Layer
if st.session_state.df is not None:
    df = st.session_state.df
    st.success(f"✨ Successfully mapped {len(df)} transactions from your statement!")

    # Analytics Dashboard Metrics
    total_in = df["Paid In (Cr)"].dropna().sum()
    total_out = df["Withdrawn (Dr)"].dropna().sum()
    net_flow = total_in - total_out

    st.markdown("### 📊 Statement Summary")
    mcol1, mcol2, mcol3, mcol4 = st.columns(4)
    mcol1.metric("Total Inflows (Cr)", f"Ksh {total_in:,.2f}")
    mcol2.metric("Total Outflows (Dr)", f"Ksh {total_out:,.2f}")
    mcol3.metric("Net Cash Flow", f"Ksh {net_flow:,.2f}")
    mcol4.metric("Total Entries", f"{len(df)}")
    st.markdown("---")

    # Interactive Search & Filter Toolbar
    st.markdown("### 🔍 Filter Ledger")
    fcol1, fcol2 = st.columns([3, 1])
    with fcol1:
        search_query = st.text_input("Search transactions by keyword (Names, Paybills, Tills, etc.)")
    with fcol2:
        filter_type = st.selectbox("Filter Type", ["All", "Paid In (Cr)", "Withdrawn (Dr)"])

    # Apply Filters instantly on cached dataset
    filtered_df = df.copy()
    if search_query:
        query_str = str(search_query).strip()
        filtered_df = filtered_df[
            filtered_df["Details"].str.contains(query_str, case=False, na=False) |
            filtered_df["Receipt No."].str.contains(query_str, case=False, na=False) |
            filtered_df["Paid In (Cr)"].dropna().astype(str).str.contains(query_str, case=False, na=False) |
            filtered_df["Withdrawn (Dr)"].dropna().astype(str).str.contains(query_str, case=False, na=False)
        ]
    if filter_type == "Paid In (Cr)":
        filtered_df = filtered_df[filtered_df["Paid In (Cr)"].notnull()]
    elif filter_type == "Withdrawn (Dr)":
        filtered_df = filtered_df[filtered_df["Withdrawn (Dr)"].notnull()]

    st.dataframe(filtered_df, use_container_width=True)

    # Secure Download Block
    st.markdown("### 📥 Download Export")
    if st.session_state.excel_buffer:
        st.download_button(
            label="📥 Download Professional Excel File (.xlsx)",
            data=st.session_state.excel_buffer,
            file_name="buku_cleaned_mpesa_statement.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    # Clean reset button
    if st.button("Clear Cache & Load New Statement", use_container_width=True):
        st.session_state.df = None
        st.session_state.excel_buffer = None
        st.rerun()
