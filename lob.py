import pandas as pd
import streamlit as st
import io

# Sidebar
st.sidebar.header("📌 Instructions")
st.sidebar.markdown(
    """
    1. Upload your **ETA file** (CSV or Excel). 
    2. Upload your **LOB format file** (CSV or Excel).   
    3. The app will align the **Stockcodes** and fill the LOB accordingly.  
    4. Review the generated table and download if needed.  

    ---
    💡 Tip: Use Excel files saved as `.xlsx` for best compatibility.
    """
)

# Default supplier colors
supplier_colors = {
    "Sup1": "#FFD966",
    "Sup2": "#A4C2F4"
}

st.title("Dynamic Stock ETA Table Generator")

# Step 1 — Upload supplier ETA file
uploaded_file = st.file_uploader("Upload your supplier ETA file (CSV or Excel)", type=["csv", "xlsx"])

if uploaded_file is not None:
    # Read file
    if uploaded_file.name.endswith(".csv"):
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file)

    df.columns = df.columns.str.strip()
    st.subheader("Uploaded Supplier Data")
    st.dataframe(df)

    # Map columns
    st.markdown("### Map Columns")
    stock_col = st.selectbox("Stockcode column", df.columns, index=0)
    supplier_col = st.selectbox("Supplier column", df.columns, index=1)
    qty_col = st.selectbox("Quantity column", df.columns, index=2)
    eta_col = st.selectbox("ETA column", df.columns, index=3)

    # Ensure ETA column is datetime
    df[eta_col] = pd.to_datetime(df[eta_col], errors='coerce')

    if df[eta_col].isna().any():
        st.warning("Some ETA values could not be parsed as dates and will be blank.")

    # Step 2 — Dynamic stockcodes from ETA file
    stockcodes = df[stock_col].unique().tolist()

    # Step 3 — Upload or define initial LOB table
    init_file = st.file_uploader("Upload your initial LOB table (CSV or Excel)", type=["csv", "xlsx"])
    if init_file is not None:
        if init_file.name.endswith(".csv"):
            initial_df = pd.read_csv(init_file, index_col=0)
        else:
            initial_df = pd.read_excel(init_file, index_col=0)
        # Use uploaded columns as A/C# columns
        ac_columns = initial_df.columns.astype(str).tolist()
    else:
        # Ask user for A/C# columns
        ac_input = st.text_input("Enter A/C# columns separated by commas", "AC001,AC002,AC003")
        ac_columns = [x.strip() for x in ac_input.split(",")]
        initial_df = pd.DataFrame(index=stockcodes, columns=ac_columns)

    # Normalize empty cells
    initial_df = initial_df.replace({None: pd.NA, "": pd.NA})

    # Ensure all stockcodes exist in index
    for stock in stockcodes:
        if stock not in initial_df.index:
            initial_df.loc[stock] = [pd.NA]*len(ac_columns)

    # Ensure all A/C# columns exist
    for col in ac_columns:
        if col not in initial_df.columns:
            initial_df[col] = pd.NA

    initial_df = initial_df[ac_columns]  # keep correct column order
    st.subheader("Initial LOB Table")
    st.dataframe(initial_df)

    # Step 4 — Fill ETAs dynamically
    styled_table = initial_df.copy()

    for stock in stockcodes:
        stock_rows = df[df[stock_col] == stock]
        ac_idx = 0
        for _, row in stock_rows.iterrows():
            for _ in range(int(row[qty_col])):
                # Find next empty A/C# column
                while ac_idx < len(ac_columns) and pd.notna(styled_table.loc[stock, ac_columns[ac_idx]]):
                    ac_idx += 1
                if ac_idx >= len(ac_columns):
                    break
                eta_str = row[eta_col].strftime("%m-%d-%y") if pd.notna(row[eta_col]) else ""
                supplier_name = row[supplier_col]
                color = supplier_colors.get(supplier_name, "#FFFFFF")
                styled_table.loc[stock, ac_columns[ac_idx]] = (
                    f'<div style="background-color:{color}; padding:4px;">'
                    f'{eta_str} - {supplier_name}</div>'
                )
                ac_idx += 1

    # Step 5 — Display colored table
    st.subheader("Filled ETA Table (Color Coded by Supplier)")
    st.write(styled_table.to_html(escape=False), unsafe_allow_html=True)

    # Step 6 — Excel download
    output = io.BytesIO()
    try:
        writer_engine = 'xlsxwriter'
        import xlsxwriter
    except ImportError:
        writer_engine = 'openpyxl'
        import openpyxl

    with pd.ExcelWriter(output, engine=writer_engine) as writer:
        df_for_excel = styled_table.replace(r'<.*?>', '', regex=True)
        df_for_excel.to_excel(writer, index=True, sheet_name='ETA')
        workbook = writer.book
        worksheet = writer.sheets['ETA']

        if writer_engine == 'xlsxwriter':
            for r_idx, stock in enumerate(styled_table.index, start=1):
                for c_idx, col in enumerate(styled_table.columns, start=1):
                    cell_val = styled_table.loc[stock, col]
                    if pd.isna(cell_val):
                        continue
                    for supplier, color in supplier_colors.items():
                        if supplier in str(cell_val):
                            clean_text = (
                                str(cell_val)
                                .replace(f'<div style="background-color:{color}; padding:4px;">', '')
                                .replace('</div>', '')
                            )
                            fmt = workbook.add_format({'bg_color': color})
                            worksheet.write(r_idx, c_idx, clean_text, fmt)
                            break

    output.seek(0)
    st.download_button(
        "Download as Excel",
        data=output,
        file_name="filled_eta_table.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
