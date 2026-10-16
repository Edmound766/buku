import streamlit as st

st.set_page_config(page_title="Privacy Policy - Buku", page_icon="🔒")

st.title("Privacy Policy")
st.caption("Effective Date: October 16, 2026")

st.markdown("""
### 1. Handling of Sensitive PDF Data & Passwords
* **Temporary Processing:** When you upload an M-Pesa PDF statement, it is held strictly in temporary RAM during extraction and Excel file generation.
* **Zero Password Retention:** Your PDF passwords are used exclusively to unlock your document in real-time. They are **never saved, logged, or recorded** on our systems.
* **Automatic File Disposal:** Converted Excel files and temporary PDF memory streams are discarded immediately after processing. We do not maintain any database of your financial transactions.

### 2. Advertising & Analytics
* **Google AdSense:** We display third-party advertisements served by networks like Google AdSense. Ad networks may use cookies and device identifiers to show relevant non-personalized or personalized ads based on browsing activity.
* **Usage Statistics:** We may log basic, non-identifying telemetry (such as total daily conversions or error rate tracking) to ensure server stability.

### 3. Data Transmission Security
All data transferred between your browser and our application is encrypted end-to-end using standard HTTPS/SSL encryption protocols.
""")
