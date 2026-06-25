import datetime
import glob
import os
import re
import numpy as np
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
import pandas as pd

print("--> Memulai pemrosesan data...")

ptm_files = glob.glob("PTM*.xlsx")
monitoring_files = glob.glob("Monitoring*.xlsx")

if not ptm_files or not monitoring_files:
    print("--> File tidak ditemukan.")
    exit()

ptm_file = ptm_files[0]
monitoring_file = monitoring_files[0]

xl = pd.ExcelFile(ptm_file)
sheet_names = xl.sheet_names

wb = openpyxl.load_workbook(monitoring_file)


def clean_key(s):
    if s is None:
        return ""
    return re.sub(r"[^a-zA-Z0-9]", "", str(s)).lower()


ws = None
header_row = 3
for sheet_name in wb.sheetnames:
    test_ws = wb[sheet_name]
    for r in range(1, 11):
        for c in range(1, 5):
            if clean_key(test_ws.cell(row=r, column=c).value) == "keterangan":
                ws = test_ws
                header_row = r
                break
        if ws:
            break
    if ws:
        break

if not ws:
    ws = wb.active

col_map = {}
for i in range(1, ws.max_column + 1):
    val = ws.cell(row=header_row, column=i).value
    if val:
        if isinstance(val, datetime.datetime):
            val_str = val.strftime("%b%y").lower()
        else:
            val_str = clean_key(val)
        col_map[val_str] = i

row_map = {}
for row in range(header_row + 1, ws.max_row + 1):
    val = ws.cell(row=row, column=1).value
    if val:
        row_map[clean_key(val)] = row

alias_map = {
    "0725": "jul25",
    "0825": "aug25",
    "0925": "sep25",
    "1025": "oct25",
    "1125": "nov25",
    "1225": "dec25",
    "0126": "jan26",
    "0226": "feb26",
    "0326": "mar26",
    "0426": "apr26",
    "0526": "may26",
}

sisa_piutang_col = clean_key("Sisa Piutang")
umur_japo_col = clean_key("Umur AR Base on Tgl Japo")
nama_penjual_col = clean_key("Nama Penjual")
nama_pelanggan_col = clean_key("Nama Pelanggan")
negara_pelanggan_col = clean_key("Negara Pelanggan")

for sheet in sheet_names:
    alias = alias_map.get(clean_key(sheet))
    if not alias or alias not in col_map:
        continue

    df_all = pd.read_excel(ptm_file, sheet_name=sheet, header=None)
    header_idx = None
    for idx, row in df_all.iterrows():
        row_clean = [clean_key(x) for x in row.dropna()]
        if sisa_piutang_col in row_clean or nama_penjual_col in row_clean:
            header_idx = idx
            break

    if header_idx is None:
        header_idx = 0

    df = pd.read_excel(ptm_file, sheet_name=sheet, skiprows=header_idx)
    df_out = df.copy()
    df.columns = [clean_key(c) for c in df.columns]

    if not {
        sisa_piutang_col,
        umur_japo_col,
        nama_penjual_col,
        nama_pelanggan_col,
        negara_pelanggan_col,
    }.issubset(df.columns):
        continue

    df["days"] = (
        df[umur_japo_col].astype(str).str.extract(r"(-?\d+)").astype(float)
    )

    neg_pel_series = df[negara_pelanggan_col].fillna("").astype(str).str.strip()
    nam_pel_series = df[nama_pelanggan_col].fillna("").astype(str).str.strip()

    is_empty_neg = neg_pel_series.str.lower().isin(["", "nan", "none", "nat"])

    df["cust_id"] = neg_pel_series
    df.loc[is_empty_neg, "cust_id"] = nam_pel_series.loc[is_empty_neg]

    is_empty_final = (
        df["cust_id"].astype(str).str.strip().str.lower().isin(["", "nan", "none", "nat"])
    )
    df.loc[is_empty_final, "cust_id"] = np.nan

    not_fraud = ~df[nama_penjual_col].astype(str).str.contains(
        "FRAUD", case=False, na=False
    )
    is_60_364 = df["days"].between(60, 364)
    is_365_up = df["days"] >= 365

    tot_ar = pd.to_numeric(df[sisa_piutang_col], errors="coerce").sum()
    ar_60 = pd.to_numeric(
        df[is_60_364 & not_fraud][sisa_piutang_col], errors="coerce"
    ).sum()
    ar_365 = pd.to_numeric(
        df[is_365_up & not_fraud][sisa_piutang_col], errors="coerce"
    ).sum()

    pct_60_amt = ar_60 / tot_ar if tot_ar else 0
    pct_365_amt = ar_365 / tot_ar if tot_ar else 0

    cust_tot = df["cust_id"].nunique(dropna=True)
    cust_60 = df[is_60_364 & not_fraud]["cust_id"].nunique(dropna=True)
    cust_365 = df[is_365_up & not_fraud]["cust_id"].nunique(dropna=True)

    pct_60_cust = cust_60 / cust_tot if cust_tot else 0
    pct_365_cust = cust_365 / cust_tot if cust_tot else 0

    metrics = {
        "TOTAL AR OUTSTANDING (AMOUNT)": tot_ar,
        "AR 60 HARI UP (AMOUNT)": ar_60,
        "AR BAD DEBT (365 UP ) AMOUNT": ar_365,
        "AR 60 HARI UP IN AMOUNT (%)": pct_60_amt,
        "AR BAD DEBT IN AMOUNT (365 UP ) %": pct_365_amt,
        "JUMLAH CUSTOMER (DISTINCT COUNT)": cust_tot,
        "JUMLAH CUST AR 60 HARI UP (DISTINCT COUNT)": cust_60,
        "JUMLAH CUST AR BAD DEBT (365 UP) DISTINCT COUNT": cust_365,
        "AR 60 HARI UP (%)": pct_60_cust,
        "AR BAD DEBT (365 UP ) %": pct_365_cust,
    }

    col_idx = col_map[alias]
    for label, val in metrics.items():
        r_idx = row_map.get(clean_key(label))
        if r_idx:
            ws.cell(row=r_idx, column=col_idx, value=val)

    df_out["Calculated_Days"] = df["days"]
    df_out["Calculated_Cust_ID"] = df["cust_id"]
    df_out["Is_Fraud"] = ~not_fraud

    df_60 = df_out[is_60_364 & not_fraud]
    df_365 = df_out[is_365_up & not_fraud]
    df_distinct_cust = df_out.dropna(subset=["Calculated_Cust_ID"]).drop_duplicates(
        subset=["Calculated_Cust_ID"]
    )

    detail_sheet_title = f"Detail_{sheet}"
    if detail_sheet_title in wb.sheetnames:
        del wb[detail_sheet_title]
    new_ws = wb.create_sheet(title=detail_sheet_title)

    def write_table(title, sub_df):
        new_ws.append([title])
        sub_df = sub_df.astype(object).where(pd.notnull(sub_df), None)
        for r in dataframe_to_rows(sub_df, index=False, header=True):
            new_ws.append(r)
        new_ws.append([])
        new_ws.append([])

    write_table(
        "BARIS DATA FILTER: AR 60 HARI UP (AMOUNT) & JUMLAH CUST AR 60 HARI UP",
        df_60,
    )
    write_table(
        "BARIS DATA FILTER: AR BAD DEBT (365 UP) AMOUNT & JUMLAH CUST AR BAD DEBT",
        df_365,
    )

wb.save(monitoring_file)
print(
    "--> Proses selesai. Masalah sinkronisasi Nama/Negara Pelanggan berhasil diperbaiki."
)
