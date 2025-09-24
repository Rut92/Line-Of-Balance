import pandas as pd
import streamlit as st
import io

# ---------- Page & Sidebar ----------
st.set_page_config(page_title="LOB Filler", layout="wide")

st.sidebar.header("📌 Instructions")
st.sidebar.markdown(
    """
    1) Upload your **ETA file** (CSV or Excel)  
    2) Upload your **LOB format file** (CSV or Excel)  
    3) The app will fill the next empty A/C# cells for *existing* stockcodes only  
    4) Download the filled table as Excel

    **Note:** No new rows will be added. Any stockcodes in ETA that aren’t in the LOB are ignored.
    """
)

# Toggle to see what’s ignored
show_ignored = st.sidebar.checkbox("Show ignored stockcodes from ETA", value=True)

# Default supplier colors (add more as needed)
supplier_colors = {
    "Sup1": "#FFD966",  # light yellow
    "Sup2": "#A4C2F4",  # light blue
}

st.title("Dynamic Stock ETA Table Generator (Strict LOB Match)")

# ---------- Helpers ----------
def _std_str_series(s):
    """Strip & cast to string safely."""
    return s.astype(str).str.strip()

def _is_empty_cell(val) -> bool:
    """Treat NaN / None / '' / 'Stock' as empty."""
    if val is None:
        return True
    if isinstance(val, float) and pd.isna(val):
        return True
    if isinstance(val, str):
        return val.strip() in ("", "NaN", "nan", "None", "Stock")
    return False

# ---------- Step 1: Upload ETA ----------
uploaded_eta = st.file_uploader("Upload supplier ETA file (CSV or Excel)", type=["csv", "xlsx"])

if uploaded_eta is not None:
    # Read ETA
    if uploaded_eta.name.endswith(".csv"):
        df_eta = pd.read_csv(uploaded_eta)
    else:
        df_eta = pd.read_excel(uploaded_eta)

    # Clean columns
    df_eta.columns = df_eta.columns.str.strip()
    st.subheader("Uploaded ETA Data")
    st.dataframe(df_eta, use_container_width=True)

    # Map columns
    st.markdown("### Map ETA Columns")
    stock_col = st.selectbox("Stockcode column", df_eta.columns, index=0)
    supplier_col = st.selectbox("Supplier column", df_eta.columns, index=1)
    qty_col = st.selectbox("Quantity column", df_eta.columns, index=2)
    eta_col = st.selectbox("ETA (date) column", df_eta.columns, index=3)

    # Standardize & convert
    df_eta[stock_col] = _std_str_series(df_eta[stock_col])
    df_eta[supplier_col] = _std_str_series(df_eta[supplier_col])
    # quantities to numeric (invalid -> 0)
    df_eta[qty_col] = pd.to_numeric(df_eta[qty_col], errors="coerce").fillna(0).astype(int)
    # parse dates
    df_eta[eta_col] = pd.to_datetime(df_eta[eta_col], errors="coerce")

    if df_eta[eta_col].isna().any():
        st.warning("Some ETA values could not be parsed as dates and will be left blank.")

    # ---------- Step 2: Upload LOB ----------
    init_file = st.file_uploader("Upload your LOB table (CSV or Excel)", type=["csv", "xlsx"])
    if init_file is not None:
        if init_file.name.endswith(".csv"):
            df_lob = pd.read_csv(init_file, index_col=0)
        else:
            df_lob = pd.read_excel(init_file, index_col=0)

        # Standardize index (stockcodes) & columns (A/C#)
        df_lob.index = _std_str_series(df_lob.index.to_series())
        df_lob.columns = [str(c).strip() for c in df_lob.columns]
        ac_columns = df_lob.columns.tolist()

        # Normalize empties
        df_lob = df_lob.replace({None: pd.NA, "": pd.NA})

        st.subheader("Initial LOB Table (unchanged)")
        st.dataframe(df_lob, use_container_width=True)

        # ---------- STRICT MATCH: Only use stockcodes present in LOB ----------
        present_stockcodes = set(df_lob.index.tolist())
        df_eta_valid = df_eta[df_eta[stock_col].isin(present_stockcodes)].copy()
        ignored = sorted(set(df_eta[stock_col]) - present_stockcodes)

        if show_ignored and ignored:
            st.info(f"Ignored {len(ignored)} stockcode(s) not in LOB: {', '.join(ignored[:20])}" + (" ..." if len(ignored) > 20 else ""))

        # ---------- Fill ETAs into a working copy ----------
        styled_table = df_lob.copy()

        # For each stockcode present in the LOB:
        for stock in sorted(df_eta_valid[stock_col].unique()):
            rows = df_eta_valid[df_eta_valid[stock_col] == stock]

            # Track the next empty A/C# position per stock
            ac_idx = 0

            # Advance to first empty cell for this stock
            while ac_idx < len(ac_columns) and not _is_empty_cell(styled_table.loc[stock, ac_columns[ac_idx]]):
                ac_idx += 1

            # Iterate ETA rows for this stock
            for _, r in rows.iterrows():
                qty = int(r[qty_col])
                if qty <= 0:
                    continue

                # Format ETA & color
                eta_str = r[eta_col].strftime("%m-%d-%y") if pd.notna(r[eta_col]) else ""
                supplier_name = r[supplier_col]
                bgcolor = supplier_colors.get(supplier_name, "#FFFFFF")

                # Place qty units across next empty A/C# cells
                for _ in range(qty):
                    # Find next empty
                    while ac_idx < len(ac_columns) and not _is_empty_cell(styled_table.loc[stock, ac_columns[ac_idx]]):
                        ac_idx += 1
                    if ac_idx >= len(ac_columns):
                        # No more space in LOB row; stop placing
                        break

                    styled_table.loc[stock, ac_columns[ac_idx]] = (
                        f'<div style="background-color:{bgcolor}; padding:4px;">'
                        f'{eta_str} - {supplier_name}</div>'
                    )
                    ac_idx += 1

        # ---------- Show filled table ----------
        st.subheader("Filled ETA Table (Color Coded by Supplier)")
        st.write(styled_table.to_html(escape=False), unsafe_allow_html=True)

        # ---------- Excel download (preserve colors with xlsxwriter if available) ----------
        output = io.BytesIO()
        try:
            writer_engine = "xlsxwriter"
            import xlsxwriter  # noqa: F401
        except ImportError:
            writer_engine = "openpyxl"
            import openpyxl  # noqa: F401

        with pd.ExcelWriter(output, engine=writer_engine) as writer:
            # Remove HTML for Excel text
            df_for_xlsx = styled_table.replace(r"<.*?>", "", regex=True)
            df_for_xlsx.to_excel(writer, index=True, sheet_name="ETA")
            workbook = writer.book
            worksheet = writer.sheets["ETA"]

            # Apply background colors if using xlsxwriter
            if writer_engine == "xlsxwriter":
                for r_idx, stock in enumerate(styled_table.index, start=1):
                    for c_idx, col in enumerate(styled_table.columns, start=1):
                        cell_val = styled_table.loc[stock, col]
                        if pd.isna(cell_val):
                            continue
                        text_val = str(cell_val)
                        # Determine supplier and color
                        fmt_to_use = None
                        clean_text = text_val
                        for sup, color in supplier_colors.items():
                            if sup in text_val:
                                fmt_to_use = workbook.add_format({"bg_color": color})
                                clean_text = (
                                    text_val.replace(f'<div style="background-color:{color}; padding:4px;">', "")
                                            .replace("</div>", "")
                                )
                                break
                        # Write with/without format
                        if fmt_to_use:
                            worksheet.write(r_idx, c_idx, clean_text, fmt_to_use)
                        else:
                            # No mapped color => write plain text (HTML already stripped)
                            worksheet.write(r_idx, c_idx, clean_text)

        output.seek(0)
        st.download_button(
            "Download filled ETA as Excel",
            data=output,
            file_name="filled_eta_table.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
