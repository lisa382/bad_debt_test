"""
Bad Debt Reserve Dashboard
==========================

Läser automatiskt in det senaste NetSuite-utdraget som ligger i samma mapp
som denna fil, beräknar bad debt-reserv enligt reglerna i
"Bad debt regler .docx", och visar en interaktiv dashboard.

Byt bara ut Excel-filen i mappen - ingen kod behöver ändras.

Körs med:
    streamlit run app.py
"""

from __future__ import annotations

import glob
import io
import os
from datetime import date, datetime

import pandas as pd
import streamlit as st

# ---------------------------------------------------------------------------
# Konfiguration
# ---------------------------------------------------------------------------

APP_DIR = os.path.dirname(os.path.abspath(__file__))

# Prefix för filer som dashboarden själv genererar (exports).
# Dessa ska ALDRIG plockas upp som källdata.
EXPORT_PREFIX = "BadDebt_Export_"

# Mappar som aldrig ska genomsökas efter källdata.
EXCLUDED_DIR_NAMES = {"använd ej", "anvand ej"}

REQUIRED_COLUMNS = [
    "Internal ID",
    "Date",
    "Type",
    "Document Number",
    "Name",
    "Due Date/Receive By",
    "Amount",
    "Amount Remaining",
    "Currency",
]

# Reservtrappa: (max_days_inclusive, reserve_pct, bucket_label)
# Sorterad stigande. Sista bucketen (None) fångar allt över föregående gräns.
AGING_LADDER = [
    (30, 0.005, "0-30"),
    (60, 0.01, "31-60"),
    (90, 0.045, "61-90"),
    (180, 0.10, "91-180"),
    (300, 0.25, "181-300"),
    (500, 0.50, "301-500"),
    (800, 0.60, "501-800"),
    (None, 0.70, "801+"),
]
BUCKET_ORDER = [b for _, _, b in AGING_LADDER]

ACCOUNT_RESERVE = "6350"
ACCOUNT_BALANCE_SHEET = "1515"


# ---------------------------------------------------------------------------
# Källfil - hitta senaste relevanta NetSuite-utdrag
# ---------------------------------------------------------------------------

def _is_excel_lock_file(path: str) -> bool:
    return os.path.basename(path).startswith("~$")


def find_source_candidates() -> list[str]:
    """Xlsx-filer direkt i appens mapp som ser ut som NetSuite-utdrag."""
    candidates = []
    for path in glob.glob(os.path.join(APP_DIR, "*.xlsx")):
        name = os.path.basename(path)
        if _is_excel_lock_file(path):
            continue
        if name.startswith(EXPORT_PREFIX):
            continue
        candidates.append(path)
    return candidates


def file_looks_like_netsuite_export(path: str) -> bool:
    try:
        header_df = pd.read_excel(path, nrows=0)
    except Exception:
        return False
    cols = set(header_df.columns.astype(str))
    # Kräver att minst de kritiska kolumnerna finns.
    critical = {"Internal ID", "Date", "Type", "Amount Remaining"}
    return critical.issubset(cols)


def find_latest_source_file() -> str | None:
    candidates = [p for p in find_source_candidates() if file_looks_like_netsuite_export(p)]
    if not candidates:
        return None
    candidates.sort(key=lambda p: os.path.getmtime(p), reverse=True)
    return candidates[0]


# ---------------------------------------------------------------------------
# Inläsning och beräkning
# ---------------------------------------------------------------------------

@st.cache_data(show_spinner=False)
def load_raw_data(path: str, mtime: float) -> pd.DataFrame:
    """mtime skickas in enbart för att invalidera cachen när filen ändras."""
    df = pd.read_excel(path)
    for col in REQUIRED_COLUMNS:
        if col not in df.columns:
            df[col] = pd.NA
    return df


def bucket_for_days(days: float) -> tuple[str, float]:
    if pd.isna(days):
        return ("Okänd", 0.0)
    d = max(days, 0)
    for max_days, pct, label in AGING_LADDER:
        if max_days is None or d <= max_days:
            return (label, pct)
    return ("801+", 0.70)


def validate_and_prepare(df: pd.DataFrame, calc_date: date) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Returnerar (clean_df, warnings_df).
    clean_df innehåller endast rader som är godkända för beräkning, med
    samtliga beräknade kolumner tillagda.
    warnings_df listar rader som exkluderats och varför.
    """
    work = df.copy()
    work["_row_num"] = range(2, len(work) + 2)  # ungefärlig Excel-radnr (1 = header)

    warnings_rows = []
    excluded_idx = set()

    def flag(mask, reason: str):
        idx = work.index[mask]
        for i in idx:
            if i not in excluded_idx:
                warnings_rows.append(
                    {
                        "Internal ID": work.at[i, "Internal ID"],
                        "Date": work.at[i, "Date"],
                        "Type": work.at[i, "Type"],
                        "Document Number": work.at[i, "Document Number"],
                        "Name": work.at[i, "Name"],
                        "Amount Remaining": work.at[i, "Amount Remaining"],
                        "Reason": reason,
                    }
                )
        excluded_idx.update(idx)

    # 1. Date saknas
    missing_date = work["Date"].isna()
    flag(missing_date, "Date saknas")

    # 2. Amount Remaining ej numeriskt
    amt_numeric = pd.to_numeric(work["Amount Remaining"], errors="coerce")
    non_numeric_amt = amt_numeric.isna() & work["Amount Remaining"].notna()
    also_missing_amt = work["Amount Remaining"].isna()
    flag(non_numeric_amt, "Amount Remaining är inte numeriskt")
    flag(also_missing_amt, "Amount Remaining saknas")
    work["_Amount Remaining Num"] = amt_numeric

    # 3. Type ej Invoice/Credit Memo
    valid_types = {"Invoice", "Credit Memo"}
    bad_type = ~work["Type"].isin(valid_types)
    flag(bad_type, "Type är varken Invoice eller Credit Memo")

    # 4. Internal ID eller Document Number saknas
    missing_id = work["Internal ID"].isna() | work["Document Number"].isna()
    flag(missing_id, "Internal ID eller Document Number saknas")

    # 5. Dubbletter på Internal ID (samma dokument inläst flera gånger)
    dup_mask = work["Internal ID"].notna() & work.duplicated(subset=["Internal ID"], keep="first")
    flag(dup_mask, "Dubblett - Internal ID förekommer flera gånger")

    # 6. Fakturadatum efter beräkningsdatum
    calc_ts = pd.Timestamp(calc_date)
    date_ts = pd.to_datetime(work["Date"], errors="coerce")
    future_date = date_ts.notna() & (date_ts > calc_ts)
    flag(future_date, "Fakturadatum ligger efter beräkningsdatum")

    warnings_df = pd.DataFrame(
        warnings_rows,
        columns=["Internal ID", "Date", "Type", "Document Number", "Name", "Amount Remaining", "Reason"],
    )

    clean = work.drop(index=list(excluded_idx)).copy()
    clean["Date"] = pd.to_datetime(clean["Date"], errors="coerce")

    # Signed Open Amount
    def signed(row):
        amt = row["_Amount Remaining Num"]
        if row["Type"] == "Credit Memo":
            return -amt
        return amt

    clean["Signed Open Amount"] = clean.apply(signed, axis=1)
    clean["Amount Remaining"] = clean["_Amount Remaining Num"]

    # Days Open
    clean["Days Open"] = (calc_ts - clean["Date"]).dt.days

    # Aging Bucket + Reserve %
    bucket_pct = clean["Days Open"].apply(bucket_for_days)
    clean["Aging Bucket"] = bucket_pct.apply(lambda t: t[0])
    clean["Reserve %"] = bucket_pct.apply(lambda t: t[1])

    # Bad Debt Reserve
    clean["Bad Debt Reserve"] = clean["Signed Open Amount"] * clean["Reserve %"]

    clean = clean.drop(columns=["_Amount Remaining Num", "_row_num"], errors="ignore")

    return clean, warnings_df


# ---------------------------------------------------------------------------
# Aggregeringar
# ---------------------------------------------------------------------------

def build_aging_overview(clean: pd.DataFrame) -> pd.DataFrame:
    grp = clean.groupby("Aging Bucket").agg(
        **{
            "Open AR": ("Signed Open Amount", "sum"),
            "Bad Debt Reserve": ("Bad Debt Reserve", "sum"),
        }
    )
    grp = grp.reindex(BUCKET_ORDER).fillna(0.0)
    # Reserve % per bucket = viktad andel (reserve/open ar), inte den nominella satsen,
    # eftersom Open AR kan innehålla credit memos som drar ner basen.
    grp["Bad Debt %"] = grp.apply(
        lambda r: (r["Bad Debt Reserve"] / r["Open AR"]) if r["Open AR"] not in (0, 0.0) else 0.0,
        axis=1,
    )
    total_reserve = grp["Bad Debt Reserve"].sum()
    grp["Andel av total reserv"] = grp["Bad Debt Reserve"].apply(
        lambda v: (v / total_reserve) if total_reserve not in (0, 0.0) else 0.0
    )
    grp = grp.reset_index().rename(columns={"index": "Aging Bucket"})
    return grp[["Aging Bucket", "Open AR", "Bad Debt %", "Bad Debt Reserve", "Andel av total reserv"]]


def build_customer_analysis(clean: pd.DataFrame) -> pd.DataFrame:
    grp = clean.groupby("Name").agg(
        **{
            "Open AR": ("Signed Open Amount", "sum"),
            "Bad Debt Reserve": ("Bad Debt Reserve", "sum"),
            "Oldest Invoice": ("Date", "min"),
            "Max Days Open": ("Days Open", "max"),
        }
    )
    grp["Effective Reserve %"] = grp.apply(
        lambda r: (r["Bad Debt Reserve"] / r["Open AR"]) if r["Open AR"] not in (0, 0.0) else 0.0,
        axis=1,
    )
    grp = grp.reset_index().rename(columns={"Name": "Customer"})
    grp = grp.sort_values("Bad Debt Reserve", ascending=False)
    return grp[
        ["Customer", "Open AR", "Bad Debt Reserve", "Effective Reserve %", "Oldest Invoice", "Max Days Open"]
    ]


def build_invoice_detail(clean: pd.DataFrame) -> pd.DataFrame:
    cols = [
        "Internal ID",
        "Date",
        "Type",
        "Document Number",
        "Name",
        "Amount Remaining",
        "Signed Open Amount",
        "Currency",
        "Days Open",
        "Aging Bucket",
        "Reserve %",
        "Bad Debt Reserve",
    ]
    return clean[cols].sort_values("Bad Debt Reserve", ascending=False).reset_index(drop=True)


# ---------------------------------------------------------------------------
# Export
# ---------------------------------------------------------------------------

def build_underlag_excel(
    summary_kpis: dict,
    aging_df: pd.DataFrame,
    customer_df: pd.DataFrame,
    invoice_df: pd.DataFrame,
) -> bytes:
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        wb = writer.book
        pct_fmt = wb.add_format({"num_format": "0.00%"})
        money_fmt = wb.add_format({"num_format": "#,##0.00"})
        bold = wb.add_format({"bold": True})

        # --- Summary ---
        summary_rows = [[k, v] for k, v in summary_kpis.items()]
        summary_df = pd.DataFrame(summary_rows, columns=["Nyckeltal", "Värde"])
        summary_df.to_excel(writer, sheet_name="Summary", index=False, startrow=0)
        ws = writer.sheets["Summary"]
        ws.set_column("A:A", 32, bold)
        ws.set_column("B:B", 22)

        start = len(summary_df) + 3
        ws.write(start - 1, 0, "Aging Overview", bold)
        aging_df.to_excel(writer, sheet_name="Summary", index=False, startrow=start)
        for col_idx, col in enumerate(aging_df.columns):
            fmt = pct_fmt if "%" in col or "Andel" in col else (money_fmt if col != "Aging Bucket" else None)
            if fmt:
                ws.set_column(col_idx, col_idx, 18, fmt)
            else:
                ws.set_column(col_idx, col_idx, 18)

        # --- Customer Detail ---
        customer_df.to_excel(writer, sheet_name="Customer Detail", index=False)
        ws2 = writer.sheets["Customer Detail"]
        for col_idx, col in enumerate(customer_df.columns):
            if col == "Effective Reserve %":
                ws2.set_column(col_idx, col_idx, 20, pct_fmt)
            elif col in ("Open AR", "Bad Debt Reserve"):
                ws2.set_column(col_idx, col_idx, 20, money_fmt)
            else:
                ws2.set_column(col_idx, col_idx, 24)

        # --- Invoice Detail ---
        invoice_df.to_excel(writer, sheet_name="Invoice Detail", index=False)
        ws3 = writer.sheets["Invoice Detail"]
        for col_idx, col in enumerate(invoice_df.columns):
            if col == "Reserve %":
                ws3.set_column(col_idx, col_idx, 14, pct_fmt)
            elif col in ("Amount Remaining", "Signed Open Amount", "Bad Debt Reserve"):
                ws3.set_column(col_idx, col_idx, 20, money_fmt)
            else:
                ws3.set_column(col_idx, col_idx, 18)

    buffer.seek(0)
    return buffer.getvalue()


def build_posting_export(rows: list[dict]) -> bytes:
    buffer = io.BytesIO()
    df = pd.DataFrame(rows)
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        df.to_excel(writer, sheet_name="Bokföringsorder", index=False)
        ws = writer.sheets["Bokföringsorder"]
        wb = writer.book
        money_fmt = wb.add_format({"num_format": "#,##0.00"})
        ws.set_column("A:A", 14)
        ws.set_column("B:B", 12)
        ws.set_column("C:D", 16, money_fmt)
        ws.set_column("E:E", 32)
    buffer.seek(0)
    return buffer.getvalue()


# ---------------------------------------------------------------------------
# Formattering
# ---------------------------------------------------------------------------

def fmt_sek(v: float) -> str:
    try:
        return f"{v:,.0f}".replace(",", " ")
    except Exception:
        return str(v)


def fmt_pct(v: float) -> str:
    try:
        return f"{v * 100:.2f}%"
    except Exception:
        return str(v)


# ---------------------------------------------------------------------------
# Streamlit UI
# ---------------------------------------------------------------------------

st.set_page_config(page_title="Bad Debt Reserve Dashboard", layout="wide")

st.title("Bad Debt Reserve Dashboard")

source_path = find_latest_source_file()

if source_path is None:
    st.error(
        "Hittade ingen NetSuite-fil att läsa in. Lägg en .xlsx-fil med kolumnerna "
        "Internal ID, Date, Type, Document Number, Name, Due Date/Receive By, "
        "Amount, Amount Remaining, Currency direkt i den här mappen."
    )
    st.stop()

mtime = os.path.getmtime(source_path)
raw_df = load_raw_data(source_path, mtime)

# --- Sidopanel ---
st.sidebar.header("Inställningar")
calc_date = st.sidebar.date_input("Beräkningsdatum", value=date.today())
st.sidebar.caption("Ändra datumet t.ex. till ett månadsslut. All beräkning uppdateras direkt.")

currently_booked = st.sidebar.number_input(
    "Currently booked reserve (SEK/redovisningsvaluta)",
    value=0.0,
    step=1000.0,
    format="%.2f",
)

clean, warnings_df = validate_and_prepare(raw_df, calc_date)

total_open_ar = clean["Signed Open Amount"].sum()
total_reserve = clean["Bad Debt Reserve"].sum()
reserve_pct = (total_reserve / total_open_ar) if total_open_ar not in (0, 0.0) else 0.0
doc_count = len(clean)

# --- KPI-rad ---
c1, c2, c3, c4 = st.columns(4)
c1.metric("Total Open AR", f"{fmt_sek(total_open_ar)}")
c2.metric("Required Bad Debt Reserve", f"{fmt_sek(total_reserve)}")
c3.metric("Reserve %", fmt_pct(reserve_pct))
c4.metric("Number of Open Documents", f"{doc_count:,}".replace(",", " "))

meta1, meta2 = st.columns(2)
meta1.markdown(f"**Calculation date:** {calc_date.strftime('%Y-%m-%d')}")
meta2.markdown(f"**Source file:** {os.path.basename(source_path)}")

if not warnings_df.empty:
    st.warning(
        f"{len(warnings_df)} rad(er) exkluderades från beräkningen på grund av datakvalitetsproblem. "
        "Se sektionen 'Data Quality / Warnings' längst ner."
    )

st.divider()

# --- Bokföringsorder ---
st.header("Bokföringsorder")

required_closing_reserve = total_reserve
required_adjustment = required_closing_reserve - currently_booked

bc1, bc2, bc3 = st.columns(3)
bc1.metric("Currently booked reserve", fmt_sek(currently_booked))
bc2.metric("Required closing reserve", fmt_sek(required_closing_reserve))
bc3.metric("Adjustment to book", fmt_sek(required_adjustment))

posting_date = calc_date.strftime("%Y-%m-%d")
description = f"Bad debt reserve {posting_date}"

if abs(required_adjustment) < 0.005:
    st.success("Ingen justering krävs - bokförd reserv matchar redan required closing reserve.")
    posting_rows = []
else:
    if required_adjustment > 0:
        posting_rows = [
            {
                "Posting Date": posting_date,
                "Account": ACCOUNT_RESERVE,
                "Debit": round(required_adjustment, 2),
                "Credit": 0.0,
                "Description": description,
            },
            {
                "Posting Date": posting_date,
                "Account": ACCOUNT_BALANCE_SHEET,
                "Debit": 0.0,
                "Credit": round(required_adjustment, 2),
                "Description": description,
            },
        ]
    else:
        abs_adj = abs(required_adjustment)
        posting_rows = [
            {
                "Posting Date": posting_date,
                "Account": ACCOUNT_BALANCE_SHEET,
                "Debit": round(abs_adj, 2),
                "Credit": 0.0,
                "Description": description,
            },
            {
                "Posting Date": posting_date,
                "Account": ACCOUNT_RESERVE,
                "Debit": 0.0,
                "Credit": round(abs_adj, 2),
                "Description": description,
            },
        ]

    posting_display = pd.DataFrame(posting_rows)[["Account", "Debit", "Credit", "Description"]]
    posting_display["Debit"] = posting_display["Debit"].map(lambda v: fmt_sek(v) if v else "-")
    posting_display["Credit"] = posting_display["Credit"].map(lambda v: fmt_sek(v) if v else "-")
    st.table(posting_display)

    total_debit = sum(r["Debit"] for r in posting_rows)
    total_credit = sum(r["Credit"] for r in posting_rows)
    if abs(total_debit - total_credit) < 0.005:
        st.caption(f"Debet och kredit balanserar: {fmt_sek(total_debit)}")
    else:
        st.error("Debet och kredit balanserar INTE - kontrollera beräkningen.")

st.divider()

# --- Aging overview ---
st.header("Aging Overview")
aging_df = build_aging_overview(clean)

aging_display = aging_df.copy()
aging_display["Open AR"] = aging_display["Open AR"].map(fmt_sek)
aging_display["Bad Debt %"] = aging_display["Bad Debt %"].map(fmt_pct)
aging_display["Bad Debt Reserve"] = aging_display["Bad Debt Reserve"].map(fmt_sek)
aging_display["Andel av total reserv"] = aging_display["Andel av total reserv"].map(fmt_pct)

col_table, col_chart = st.columns([1.1, 1])
with col_table:
    st.dataframe(aging_display, hide_index=True, width="stretch")
with col_chart:
    chart_df = aging_df.set_index("Aging Bucket")[["Bad Debt Reserve"]]
    st.bar_chart(chart_df)

st.divider()

# --- Customer analysis ---
st.header("Bad Debt by Customer")
customer_df = build_customer_analysis(clean)

cust_filter = st.multiselect(
    "Filtrera på kund",
    options=sorted(customer_df["Customer"].dropna().unique().tolist()),
)
cust_view = customer_df if not cust_filter else customer_df[customer_df["Customer"].isin(cust_filter)]

cust_display = cust_view.copy()
cust_display["Open AR"] = cust_display["Open AR"].map(fmt_sek)
cust_display["Bad Debt Reserve"] = cust_display["Bad Debt Reserve"].map(fmt_sek)
cust_display["Effective Reserve %"] = cust_display["Effective Reserve %"].map(fmt_pct)
cust_display["Oldest Invoice"] = pd.to_datetime(cust_display["Oldest Invoice"]).dt.strftime("%Y-%m-%d")

st.dataframe(cust_display, hide_index=True, width="stretch")
st.caption("Klicka på en kolumnrubrik för att sortera tabellen.")

st.divider()

# --- Invoice / transaction detail ---
st.header("Invoice / Transaction Detail")

f1, f2, f3, f4 = st.columns(4)
with f1:
    cust_opts = st.multiselect("Customer", sorted(clean["Name"].dropna().unique().tolist()), key="f_cust")
with f2:
    type_opts = st.multiselect("Type", sorted(clean["Type"].dropna().unique().tolist()), key="f_type")
with f3:
    curr_opts = st.multiselect("Currency", sorted(clean["Currency"].dropna().unique().tolist()), key="f_curr")
with f4:
    bucket_opts = st.multiselect("Aging Bucket", BUCKET_ORDER, key="f_bucket")

invoice_df = build_invoice_detail(clean)
filtered = invoice_df.copy()
if cust_opts:
    filtered = filtered[filtered["Name"].isin(cust_opts)]
if type_opts:
    filtered = filtered[filtered["Type"].isin(type_opts)]
if curr_opts:
    filtered = filtered[filtered["Currency"].isin(curr_opts)]
if bucket_opts:
    filtered = filtered[filtered["Aging Bucket"].isin(bucket_opts)]

detail_display = filtered.copy()
detail_display["Date"] = pd.to_datetime(detail_display["Date"]).dt.strftime("%Y-%m-%d")
detail_display["Amount Remaining"] = detail_display["Amount Remaining"].map(fmt_sek)
detail_display["Signed Open Amount"] = detail_display["Signed Open Amount"].map(fmt_sek)
detail_display["Reserve %"] = detail_display["Reserve %"].map(fmt_pct)
detail_display["Bad Debt Reserve"] = detail_display["Bad Debt Reserve"].map(fmt_sek)

st.dataframe(detail_display, hide_index=True, width="stretch", height=420)
st.caption(f"Visar {len(filtered)} av {len(invoice_df)} dokument. Klicka på kolumnrubrik för att sortera.")

st.divider()

# --- Exports ---
st.header("Export")

exp1, exp2 = st.columns(2)

with exp1:
    st.subheader("Bokföringsorder")
    if posting_rows:
        posting_bytes = build_posting_export(posting_rows)
        st.download_button(
            "Ladda ner bokföringsorder (Excel)",
            data=posting_bytes,
            file_name=f"{EXPORT_PREFIX}Bokforingsorder_{posting_date}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        posting_csv = pd.DataFrame(posting_rows)[
            ["Posting Date", "Account", "Debit", "Credit", "Description"]
        ].to_csv(index=False).encode("utf-8-sig")
        st.download_button(
            "Ladda ner bokföringsorder (CSV)",
            data=posting_csv,
            file_name=f"{EXPORT_PREFIX}Bokforingsorder_{posting_date}.csv",
            mime="text/csv",
        )
    else:
        st.caption("Ingen justering att exportera.")

with exp2:
    st.subheader("Fullständigt underlag")
    summary_kpis = {
        "Calculation date": posting_date,
        "Source file": os.path.basename(source_path),
        "Total Open AR": total_open_ar,
        "Required Bad Debt Reserve": total_reserve,
        "Reserve %": reserve_pct,
        "Number of Open Documents": doc_count,
        "Currently booked reserve": currently_booked,
        "Required adjustment": required_adjustment,
    }
    underlag_bytes = build_underlag_excel(summary_kpis, aging_df, customer_df, invoice_df)
    st.download_button(
        "Ladda ner underlag (Excel, 3 flikar)",
        data=underlag_bytes,
        file_name=f"{EXPORT_PREFIX}Underlag_{posting_date}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

st.divider()

# --- Kontrollsummeringar ---
st.header("Kontrollsummeringar")
raw_amount_remaining_total = pd.to_numeric(raw_df["Amount Remaining"], errors="coerce").sum()

k1, k2, k3 = st.columns(3)
k1.metric("Total source Amount Remaining", fmt_sek(raw_amount_remaining_total))
k2.metric("Total Signed Open Amount", fmt_sek(total_open_ar))
k3.metric("Total Bad Debt Reserve", fmt_sek(total_reserve))

st.caption(
    "Total source Amount Remaining är summan direkt från källfilen (utan sign-justering för credit memos "
    "och utan exkludering av rader med datakvalitetsproblem) - används enbart som avstämningskontroll."
)

st.divider()

# --- Data Quality / Warnings ---
st.header("Data Quality / Warnings")
if warnings_df.empty:
    st.success("Inga datakvalitetsproblem hittades. Samtliga rader i källfilen ingår i beräkningen.")
else:
    st.error(f"{len(warnings_df)} rad(er) exkluderades från beräkningen. Granska nedan.")
    wdf = warnings_df.copy()
    wdf["Date"] = pd.to_datetime(wdf["Date"], errors="coerce").dt.strftime("%Y-%m-%d")
    st.dataframe(wdf, hide_index=True, width="stretch")
