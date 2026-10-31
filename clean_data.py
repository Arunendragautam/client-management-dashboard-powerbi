"""Cleans raw_client_data.xlsx -> clean_client_data.xlsx (load this file into Power BI)."""
import re
import pandas as pd

raw = pd.read_excel("raw_client_data.xlsx", sheet_name=None, dtype=str)
clients, txns = raw["Clients"].copy(), raw["Transactions"].copy()
log = []   # data-quality log: what was fixed, and how many rows

# ---------- generic: trim whitespace, turn placeholder text into real blanks ----------
PLACEHOLDERS = {"", "n/a", "na", "none", "null", "-"}
def blank_out(df):
    for col in df.columns:
        df[col] = df[col].astype("string").str.strip()
        df.loc[df[col].str.lower().isin(PLACEHOLDERS), col] = pd.NA
    return df
clients, txns = blank_out(clients), blank_out(txns)

# ---------- CLIENTS ----------
clients["ClientName"] = clients["ClientName"].str.title()
clients["Email"] = clients["Email"].str.lower()

def clean_phone(p):
    if pd.isna(p): return pd.NA
    digits = re.sub(r"\D", "", p)[-10:]                 # keep last 10 digits
    return digits if len(digits) == 10 else pd.NA
clients["Phone"] = clients["Phone"].map(clean_phone)

city_map = {"pune": "Pune", "mumbai": "Mumbai", "nagpur": "Nagpur", "nashik": "Nashik",
            "pimpri-chinchwad": "Pimpri-Chinchwad", "pimpri chinchwad": "Pimpri-Chinchwad"}
clients["City"] = clients["City"].str.lower().map(city_map)

status_map = {"active": "Active", "inactive": "Inactive", "inactv": "Inactive"}
clients["Status"] = clients["Status"].str.lower().map(status_map)

clients["SignupDate"] = pd.to_datetime(clients["SignupDate"], dayfirst=True, errors="coerce", format="mixed")

# duplicates: keep the most complete record per ClientID
before = len(clients)
clients["_filled"] = clients.notna().sum(axis=1)
clients = (clients.sort_values("_filled", ascending=False)
                  .drop_duplicates("ClientID").drop(columns="_filled").sort_values("ClientID"))
log.append(("Duplicate client rows removed", before - len(clients)))

# missing values: label categories, never invent numbers/contacts
for col in ["Industry", "City", "Status"]:
    log.append((f"Missing {col} -> 'Unknown'", int(clients[col].isna().sum())))
    clients[col] = clients[col].fillna("Unknown")
log.append(("Missing/invalid emails left blank", int(clients["Email"].isna().sum())))
log.append(("Missing/invalid phones left blank", int(clients["Phone"].isna().sum())))

# ---------- TRANSACTIONS ----------
before = len(txns)
txns = txns.drop_duplicates("TxnID")
log.append(("Duplicate transaction rows removed", before - len(txns)))

txns["Amount"] = pd.to_numeric(txns["Amount"].str.replace(r"[^\d.]", "", regex=True), errors="coerce")
log.append(("Transactions dropped: missing amount", int(txns["Amount"].isna().sum())))
txns = txns.dropna(subset=["Amount"])

orphans = ~txns["ClientID"].isin(clients["ClientID"])
log.append(("Transactions dropped: unknown ClientID", int(orphans.sum())))
txns = txns[~orphans]
txns["TxnDate"] = pd.to_datetime(txns["TxnDate"], errors="coerce")

# ---------- save ----------
with pd.ExcelWriter("clean_client_data.xlsx", datetime_format="yyyy-mm-dd") as xw:
    clients.to_excel(xw, sheet_name="Clients", index=False)
    txns.to_excel(xw, sheet_name="Transactions", index=False)
    pd.DataFrame(log, columns=["Check", "Rows"]).to_excel(xw, sheet_name="DataQualityLog", index=False)

print(pd.DataFrame(log, columns=["Check", "Rows"]).to_string(index=False))
print(f"\nFinal: {len(clients)} clients, {len(txns)} transactions -> clean_client_data.xlsx")
