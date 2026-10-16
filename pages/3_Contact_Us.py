import streamlit as st

st.set_page_config(page_title="Contact Us - Buku", page_icon="📩")

st.title("Contact Us & Help Center")

st.markdown("""
Have questions, bug reports, or feature requests? Reach out to us below.

**Email Support:** `support@buku.co.ke`
""")

st.divider()

st.subheader("Frequently Asked Questions")

with st.expander("Is my financial data or M-Pesa password safe?"):
    st.write(
        "Yes, completely. Files and passwords are processed entirely in memory "
        "and immediately wiped upon completion. Nothing is stored on disks or databases."
    )

with st.expander("Is Buku affiliated with Safaricom or M-Pesa?"):
    st.write(
        "No. Buku is an independent document formatting tool created to help users "
        "convert exported PDF statements into spreadsheets for budgeting."
    )

with st.expander("Why are some transaction descriptions formatted differently?"):
    st.write(
        "Safaricom occasionally updates the layout of M-Pesa PDFs. If a row fails to "
        "parse correctly, send us feedback at support@buku.co.ke so we can update the parser logic."
    )
