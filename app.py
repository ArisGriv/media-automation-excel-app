import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Excel Analyzer", layout="centered")
st.title("📊 AI Excel Automation - Media Automation")

uploaded_file = st.file_uploader("Ανεβάστε το Excel αρχείο με 3 φύλλα εργασίας", type=["xlsx"])

if uploaded_file:
    try:
        # Load the Excel file
        xls = pd.ExcelFile(uploaded_file)

        # Read the 3 sheets
        df1 = xls.parse(0)  # Sheet 1: Έσοδα
        df2 = xls.parse(1)  # Sheet 2: Έξοδα

        # --- Process Revenue per Month and Apartment ---
        df1['Ημερομηνία'] = pd.to_datetime(df1['Ημερομηνία'])
        df1['Μήνας'] = df1['Ημερομηνία'].dt.to_period('M').astype(str)
        revenue_summary = df1.groupby(['Μήνας', 'Διαμέρισμα'])['Ποσό'].sum().reset_index()

        # --- Process Fixed Costs per Revenue Center ---
        fixed_df = df2[df2['Τύπος Έξοδου'] == 'Σταθερό']
        cost_summary = fixed_df.groupby('Κέντρο Εσόδου')['Ποσό'].sum().reset_index()

        # --- Create downloadable Excel ---
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            revenue_summary.to_excel(writer, index=False, sheet_name='Έσοδα_Μήνας_Διαμέρισμα')
            cost_summary.to_excel(writer, index=False, sheet_name='Σταθερά_Έξοδα_Κέντρο')

        output.seek(0)

        st.success("✅ Οι αναφορές δημιουργήθηκαν με επιτυχία!")
        st.download_button(
            label="📥 Κατεβάστε τις αναφορές (Excel)",
            data=output,
            file_name="Αναφορές_MediaAutomation.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"❌ Σφάλμα: {e}")

