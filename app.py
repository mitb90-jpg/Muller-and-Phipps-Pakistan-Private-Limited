import streamlit as st
import pandas as pd
import io
from openpyxl.styles import PatternFill, Font

# ---------------- PAGE CONFIG ----------------
st.set_page_config(
    page_title="Bank Reconciliation Tool - Muller & Phipps",
    page_icon="🏦",
    layout="wide"
)

# ---------------- BANK CONFIGURATION ----------------
# To add a new bank later: add an entry here with its sheet name, header
# row (0-indexed), and the column names for deposit slip / amount / depot.
BANKS = {
    "MCB": {
        "account_code": 212201, "account_label": "MCB Bank Ltd",
        "sheet": "MCB", "header": 1,
        "slip_col": "Deposit Slip No.", "amount_col": "Amount",
        "depot_col": "Dealer Code", "date_col": "Credit Date",
    },
    "HBL": {
        "account_code": 212202, "account_label": "Habib Bank Limited",
        "sheet": "HBL", "header": 6,
        "slip_col": "Deposit Slip No.", "amount_col": "Amount",
        "depot_col": "Dealer Code", "date_col": "Credit Date",
    },
    "UBL": {
        "account_code": 212205, "account_label": "United Bank Limited",
        "sheet": "UBL", "header": 0,
        "slip_col": "DEPOSIT SLIP NO", "amount_col": "AMOUNT",
        "depot_col": "DEPOT'S CODE", "date_col": "REALIZATION DATE",
    },
    "FAYSAL": {
        "account_code": 212209, "account_label": "Faysal Bank Limited",
        "sheet": "MIS", "header": 4,
        "slip_col": "Deposit No.", "amount_col": "Amount",
        "depot_col": "Dealer Code", "date_col": "Credit Date",
    },
}

DEPOT_CODES = [612, 613, 614, 651, 657, 645, 646, 733, 766, 778]
TOLERANCE = 1.0  # rupees - absorbs rounding differences


# ---------------- CSS ----------------
st.markdown("""
<style>
div.stDownloadButton > button, div.stButton > button {
    width: 100%;
    background-color: #1f4e79;
    color: white;
    font-size: 16px;
    font-weight: bold;
    padding: 12px;
    border-radius: 10px;
    border: none;
}
div.stDownloadButton > button:hover, div.stButton > button:hover {
    background-color: #163a5c;
}
</style>
""", unsafe_allow_html=True)


# ---------------- HELPERS ----------------

def load_gl(file):
    df = pd.read_excel(file, sheet_name="GL")
    df.columns = [str(c).strip() for c in df.columns]
    return df


def load_bank(file, bank_key):
    cfg = BANKS[bank_key]
    df = pd.read_excel(file, sheet_name=cfg["sheet"], header=cfg["header"])
    df.columns = [str(c).strip() for c in df.columns]
    return df


def reconcile_account(gl_df, bank_df, bank_key, depot_code):
    cfg = BANKS[bank_key]

    gl_sub = gl_df[gl_df["Natural Account Segment"] == cfg["account_code"]].copy()
    dr = gl_sub[gl_sub["Entered Amount DR"].notna()].copy()
    cr = gl_sub[gl_sub["Entered Amount CR"].notna()].copy()

    bank_sub = bank_df[bank_df[cfg["depot_col"]] == depot_code].copy()
    bank_sub["_used"] = False

    dr["Status"] = "UNMATCHED"
    dr["Bank Deposit Slip"] = pd.Series([""] * len(dr), index=dr.index, dtype="object")
    dr["Bank Amount"] = None
    dr["Bank Date"] = pd.Series([""] * len(dr), index=dr.index, dtype="object")

    for idx, row in dr.iterrows():
        amt = row["Entered Amount DR"]
        if pd.isna(amt):
            continue
        candidates = bank_sub[
            (~bank_sub["_used"]) & ((bank_sub[cfg["amount_col"]] - amt).abs() <= TOLERANCE)
        ]
        if len(candidates) > 0:
            match_idx = candidates.index[0]
            bank_sub.loc[match_idx, "_used"] = True
            dr.loc[idx, "Status"] = "MATCHED"
            dr.loc[idx, "Bank Deposit Slip"] = str(bank_sub.loc[match_idx, cfg["slip_col"]])
            dr.loc[idx, "Bank Amount"] = bank_sub.loc[match_idx, cfg["amount_col"]]
            dr.loc[idx, "Bank Date"] = str(bank_sub.loc[match_idx, cfg["date_col"]])

    unmatched_bank = bank_sub[~bank_sub["_used"]].drop(columns=["_used"])

    return {
        "bank": bank_key,
        "account_code": cfg["account_code"],
        "account_label": cfg["account_label"],
        "ledger_debits": dr,
        "ledger_credits": cr,
        "unmatched_bank_items": unmatched_bank,
    }


def build_excel(results, depot_code):
    output = io.BytesIO()

    red_fill = PatternFill(start_color="FDEAEA", end_color="FDEAEA", fill_type="solid")
    green_fill = PatternFill(start_color="E6F4EA", end_color="E6F4EA", fill_type="solid")
    amber_fill = PatternFill(start_color="FEF3E0", end_color="FEF3E0", fill_type="solid")
    header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)

    with pd.ExcelWriter(output, engine="openpyxl") as writer:

        # ---------------- SUMMARY SHEET ----------------
        summary_rows = []
        for r in results:
            total = len(r["ledger_debits"])
            matched = (r["ledger_debits"]["Status"] == "MATCHED").sum()
            summary_rows.append({
                "Bank": r["account_label"],
                "Account Code": r["account_code"],
                "Ledger Debit Items": total,
                "Matched": matched,
                "Needs Review (Ledger)": total - matched,
                "Unmatched Bank Items": len(r["unmatched_bank_items"]),
            })
        summary_df = pd.DataFrame(summary_rows)
        summary_df.to_excel(writer, sheet_name="Summary", index=False, startrow=1)
        ws = writer.sheets["Summary"]
        ws.cell(row=1, column=1).value = f"Bank Reconciliation Summary - Depot {depot_code}"
        for col in range(1, len(summary_df.columns) + 1):
            ws.cell(row=2, column=col).fill = header_fill
            ws.cell(row=2, column=col).font = header_font

        # ---------------- ONE SHEET PER BANK ----------------
        for r in results:
            sheet_name = r["bank"][:31]

            dr_cols = [
                "Journal Line Number", "Journal Line Description",
                "Entered Amount DR", "Status", "Bank Deposit Slip",
                "Bank Amount", "Bank Date"
            ]
            dr_display = r["ledger_debits"][[c for c in dr_cols if c in r["ledger_debits"].columns]]

            dr_display.to_excel(writer, sheet_name=sheet_name, index=False, startrow=2)
            ws = writer.sheets[sheet_name]
            ws.cell(row=1, column=1).value = (
                f"{r['account_label']} ({r['account_code']}) - Ledger Debits vs Bank"
            )

            for col in range(1, len(dr_display.columns) + 1):
                ws.cell(row=3, column=col).fill = header_fill
                ws.cell(row=3, column=col).font = header_font

            status_col_idx = list(dr_display.columns).index("Status") + 1 if "Status" in dr_display.columns else None
            if status_col_idx:
                for row_num in range(4, 4 + len(dr_display)):
                    status_val = ws.cell(row=row_num, column=status_col_idx).value
                    fill = red_fill if status_val == "UNMATCHED" else green_fill
                    for col_num in range(1, len(dr_display.columns) + 1):
                        ws.cell(row=row_num, column=col_num).fill = fill

            # ---- unmatched bank items, appended below ----
            next_row = 4 + len(dr_display) + 2
            ws.cell(row=next_row, column=1).value = "Bank items for this depot with no matching ledger entry (review):"
            ws.cell(row=next_row, column=1).font = Font(bold=True)

            bank_items = r["unmatched_bank_items"]
            if len(bank_items) > 0:
                start = next_row + 1
                for col_num, col_name in enumerate(bank_items.columns, start=1):
                    cell = ws.cell(row=start, column=col_num)
                    cell.value = col_name
                    cell.fill = header_fill
                    cell.font = header_font
                for r_idx, (_, row_data) in enumerate(bank_items.iterrows(), start=start + 1):
                    for col_num, val in enumerate(row_data, start=1):
                        cell = ws.cell(row=r_idx, column=col_num)
                        cell.value = val
                        cell.fill = amber_fill

        # ---------------- CREDIT ENTRIES (reference only) ----------------
        all_credits = pd.concat(
            [r["ledger_credits"].assign(Bank=r["bank"]) for r in results if len(r["ledger_credits"]) > 0],
            ignore_index=True
        ) if any(len(r["ledger_credits"]) > 0 for r in results) else pd.DataFrame()

        if not all_credits.empty:
            cr_cols = ["Bank", "Journal Line Number", "Journal Line Description", "Entered Amount CR"]
            cr_display = all_credits[[c for c in cr_cols if c in all_credits.columns]]
            cr_display.to_excel(writer, sheet_name="Ledger Credits (ref)", index=False, startrow=1)
            ws = writer.sheets["Ledger Credits (ref)"]
            ws.cell(row=1, column=1).value = "Credit entries in ledger (bulk transfers-out, reference only)"

    output.seek(0)
    return output


# ---------------- UI ----------------

st.title("🏦 Bank Reconciliation Tool")
st.caption("Muller & Phipps Pakistan — matches ledger (GL) collections against bank MIS files by amount, flags exceptions for manual review")

st.divider()

col1, col2 = st.columns([1, 2])

with col1:
    depot_code = st.selectbox("Select Depot Code", DEPOT_CODES)

st.subheader("1. Upload GL File")
gl_file = st.file_uploader("GL file for this depot (.xlsx)", type=["xlsx"], key="gl_upload")

st.subheader("2. Upload Bank MIS Files")
st.caption("Upload whichever bank files you have this month — combined files covering all depots are fine, the tool filters by depot code automatically.")

bank_files = {}
bcol1, bcol2, bcol3, bcol4 = st.columns(4)
with bcol1:
    bank_files["MCB"] = st.file_uploader("MCB", type=["xlsx"], key="mcb_upload")
with bcol2:
    bank_files["HBL"] = st.file_uploader("HBL", type=["xlsx"], key="hbl_upload")
with bcol3:
    bank_files["UBL"] = st.file_uploader("UBL", type=["xlsx"], key="ubl_upload")
with bcol4:
    bank_files["FAYSAL"] = st.file_uploader("Faysal", type=["xlsx"], key="faysal_upload")

st.divider()

if st.button("🔄 Run Reconciliation"):

    if gl_file is None:
        st.warning("Please upload the GL file first.")
    else:
        try:
            gl_df = load_gl(gl_file)
        except Exception as e:
            st.error(f"Could not read GL file: {e}")
            st.stop()

        results = []
        errors = []

        for bank_key, file in bank_files.items():
            if file is None:
                continue
            try:
                bank_df = load_bank(file, bank_key)
                result = reconcile_account(gl_df, bank_df, bank_key, depot_code)
                results.append(result)
            except Exception as e:
                errors.append(f"{bank_key}: {e}")

        for err in errors:
            st.error(f"Error processing {err}")

        if not results:
            st.warning("No bank files were successfully processed. Upload at least one bank file.")
        else:
            st.session_state["recon_results"] = results
            st.session_state["recon_depot"] = depot_code
            st.success(f"Reconciliation complete for depot {depot_code}")

if "recon_results" in st.session_state:

    results = st.session_state["recon_results"]
    depot_code = st.session_state["recon_depot"]

    st.subheader("Summary")

    summary_data = []
    for r in results:
        total = len(r["ledger_debits"])
        matched = (r["ledger_debits"]["Status"] == "MATCHED").sum()
        summary_data.append({
            "Bank": r["account_label"],
            "Ledger Items": total,
            "Matched": matched,
            "Needs Review": total - matched,
            "Unmatched Bank Items": len(r["unmatched_bank_items"]),
        })
    st.dataframe(pd.DataFrame(summary_data), use_container_width=True, hide_index=True)

    st.subheader("Details by Bank")

    for r in results:
        with st.expander(f"{r['account_label']} — {len(r['ledger_debits'])} ledger items"):

            dr_display = r["ledger_debits"][
                [c for c in ["Journal Line Number", "Journal Line Description", "Entered Amount DR",
                             "Status", "Bank Deposit Slip", "Bank Amount", "Bank Date"]
                 if c in r["ledger_debits"].columns]
            ]

            def highlight_status(row):
                if row["Status"] == "UNMATCHED":
                    return ["background-color: #fdeaea"] * len(row)
                return ["background-color: #e6f4ea"] * len(row)

            st.dataframe(
                dr_display.style.apply(highlight_status, axis=1),
                use_container_width=True, hide_index=True
            )

            if len(r["unmatched_bank_items"]) > 0:
                st.markdown("**Bank items for this depot with no matching ledger entry:**")
                st.dataframe(r["unmatched_bank_items"], use_container_width=True, hide_index=True)

    st.divider()

    excel_file = build_excel(results, depot_code)

    st.download_button(
        "⬇️ Download Full Reconciliation (Excel)",
        data=excel_file,
        file_name=f"Bank_Reconciliation_Depot_{depot_code}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
