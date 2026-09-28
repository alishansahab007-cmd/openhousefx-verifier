import streamlit as st
import pandas as pd
import dns.resolver
from concurrent.futures import ThreadPoolExecutor

st.set_page_config(page_title="Openhousefx List Cleaner", page_icon="📧")

st.title("📧 Openhousefx Lead List Verifier")
st.write("Upload your raw lead sheet (CSV or Excel) to automatically verify active email domain records.")

uploaded_file = st.file_uploader("Drag and drop your lead file here", type=["csv", "xlsx", "xls"])

def has_mx_record(email):
    try:
        if not isinstance(email, str) or "@" not in str(email):
            return False
        domain = str(email).split("@")[1].strip()
        dns.resolver.resolve(domain, 'MX')
        return True
    except Exception:
        return False

if uploaded_file is not None:
    try:
        if uploaded_file.name.endswith(".csv"):
            df = pd.read_csv(uploaded_file, encoding="latin1")
        else:
            df = pd.read_excel(uploaded_file)

        st.info(f"Successfully loaded **{len(df):,}** rows from `{uploaded_file.name}`.")

        email_col = next((col for col in df.columns if "email" in str(col).lower()), None)

        if email_col:
            st.success(f"Detected Email Column: `{email_col}`")

            if st.button("🚀 Verify MX Records Now", type="primary"):
                st.write("Checking domain records... Please wait.")

                with ThreadPoolExecutor(max_workers=50) as executor:
                    df['valid_domain'] = list(executor.map(has_mx_record, df[email_col]))

                cleaned_df = df[df['valid_domain'] == True].drop(columns=['valid_domain'])
                removed_count = len(df) - len(cleaned_df)

                st.success(f"Done! Retained **{len(cleaned_df):,}** valid leads. Filtered out **{removed_count:,}** dead domains.")

                st.dataframe(cleaned_df.head(10))

                csv_data = cleaned_df.to_csv(index=False).encode("utf-8")
                st.download_button(
                    label="📥 Download Cleaned CSV",
                    data=csv_data,
                    file_name="mx_verified_leads.csv",
                    mime="text/csv"
                )
        else:
            st.error("Could not find a column named 'email' in your file.")

    except Exception as e:
        st.error(f"Error reading file: {e}")
