"""
build_labels.py: parse and label M-Pesa transactions by merchant.

Command line:
  python build_labels.py make  mpesa_statement.ods   # writes merchant_labels.csv
  python build_labels.py apply mpesa_statement.ods   # writes mpesa_labelled.csv

Notebook:
  from build_labels import add_groups, CATEGORIES
  df = add_groups(df)
  # adds: tx_group, name, merchant_key, group_id
  # (and category, if merchant_labels.csv exists)
"""
import os
import re
import sys

import pandas as pd

LABELS_CSV = "merchant_labels.csv"
OUT_CSV = "mpesa_labelled.csv"

CATEGORIES = [
    "Business - Stock",
    "Business - Operating costs",
    "Business - Sales income",
    "Personal - Food & Household",
    "Personal - Bills & Utilities",
    "Personal - Lifestyle & Entertainment",
    "Savings",
    "Loans & Fuliza",
    "Fees & Charges",
    "Cash Withdrawal",
    "Person-to-Person",
]

# tx_group -> suggested category
GROUP_DEFAULTS = {
    "fee": "Fees & Charges",
    "withdrawal": "Cash Withdrawal",
    "airtime": "Personal - Bills & Utilities",
    "fuliza": "Loans & Fuliza",
    "person": "Person-to-Person",
}

# Keyword -> category suggestions for merchants/paybills. Extend while reviewing.
RULES = [
    (r"KPLC|TOKEN|WATER|SAFARICOM|ZUKU|DSTV|GOTV|WIFI|FIBRE|RENT", "Personal - Bills & Utilities"),
    (r"NAIVAS|QUICKMART|CARREFOUR|CHANDARANA|SUPERMARKET|BUTCHERY|MILK|GROCER|MART|SHOP", "Personal - Food & Household"),
    (r"KFC|JAVA|UBER|BOLT|NETFLIX|SHOWMAX|BET|SPORTPESA|ODIBETS|CINEMA|BAR |RESTAURANT|CAFE|PIZZA", "Personal - Lifestyle & Entertainment"),
    (r"FULIZA|LOAN|OKOA|TALA|BRANCH", "Loans & Fuliza"),
    (r"M-SHWARI|MSHWARI|LOCK SAVINGS|SACCO|CHAMA|SAVINGS|NCBA|KCB M-PESA", "Savings"),
]

# Rows that are not spending: the loan half of a Fuliza payment
EXCLUDE_PATTERN = r"OVERDRAFT OF CREDIT PARTY"

# Groups where the merchant/person name identifies the group
NAMED_GROUPS = ["merchant", "paybill", "person", "other"]


# ---------------------------------------------------------------- parsing
def find_details_col(df: pd.DataFrame) -> str:
    if "Details" in df.columns:
        return "Details"
    for c in df.columns:
        if "detail" in str(c).lower():
            return c
    raise KeyError(f"No Details-like column found. Columns: {list(df.columns)}")


def clean_name(s) -> str:
    """Strip phone/account numbers and punctuation; merge spelling variants."""
    s = str(s).upper()
    s = re.sub(r"\bACC(?:OUNT|\.)?\b.*$", "", s)      # drop "Acc. 1234" tails
    s = re.sub(r"[\d*]{4,}", " ", s)                  # long numbers, masked phones
    s = re.sub(r"[^A-Z& ]", " ", s)                   # punctuation
    s = re.sub(r"\s+", " ", s).strip()
    # "KPLC PREPAID" and "KPLCPREPAID" become the same key
    return s.replace(" ", "") if len(s.replace(" ", "")) <= 12 else s


def parse_details(text) -> tuple:
    """Return (tx_group, name) for one Details string."""
    t = re.sub(r"\s+", " ", str(text)).strip()
    u = t.upper()

    if re.search(EXCLUDE_PATTERN, u):
        return "fuliza_loan", t
    if "OVERDRAFT" in u or "FULIZA" in u:
        return "fuliza", "FULIZA"
    if "CHARGE" in u or "FEE" in u:
        return "fee", "FEES"
    if "AIRTIME" in u:
        return "airtime", "AIRTIME"
    if "WITHDRAWAL" in u or "WITHDRAW" in u:
        return "withdrawal", "WITHDRAWAL"

    tail = r"(?:TO|FROM)\s*-?\s*(.*)$"
    if re.search(r"PAY ?BILL", u):
        m = re.search(tail, u)
        return "paybill", (m.group(1) if m else u)
    if re.search(r"MERCHANT PAYMENT|BUY GOODS|SMALL BUSINESS|TILL", u):
        m = re.search(tail, u)
        return "merchant", (m.group(1) if m else u)
    if re.search(r"CUSTOMER TRANSFER|FUNDS RECEIVED|RECEIVED|SEND MONEY", u):
        m = re.search(tail, u)
        return "person", (m.group(1) if m else u)
    return "other", u


def add_groups(df: pd.DataFrame, labels_csv: str = LABELS_CSV) -> pd.DataFrame:
    """Add tx_group, name, merchant_key, group_id (and category if a labels
    sheet exists). Does not drop any rows."""
    out = df.copy()
    parsed = out[find_details_col(out)].map(parse_details)
    out["tx_group"] = parsed.map(lambda p: p[0])
    out["name"] = parsed.map(lambda p: p[1])

    key = out["name"].map(clean_name)
    # fee/airtime/withdrawal/fuliza: one key per group
    out["merchant_key"] = key.where(out["tx_group"].isin(NAMED_GROUPS), out["name"])
    out["group_id"] = out["merchant_key"]

    if os.path.exists(labels_csv):
        sheet = pd.read_csv(labels_csv)
        key_col = next((c for c in ("group_id", "merchant_key", "merchant")
                        if c in sheet.columns), None)
        if key_col is None or "final_category" not in sheet.columns:
            raise KeyError(f"{labels_csv} needs group_id and final_category columns. "
                           f"Re-run: python build_labels.py make <statement>")
        labels = (sheet.drop_duplicates(key_col)
                       .set_index(key_col)["final_category"].fillna(""))
        out["category"] = out["group_id"].map(labels)
        bad = set(out["category"].dropna()) - set(CATEGORIES) - {""}
        if bad:
            raise ValueError(f"Unknown categories in {labels_csv}: {sorted(bad)}")
    return out


# ---------------------------------------------------------------- CLI
def load_statement(path: str) -> pd.DataFrame:
    low = path.lower()
    if low.endswith(".ods"):
        df = pd.read_excel(path, engine="odf")
    elif low.endswith((".xlsx", ".xls")):
        df = pd.read_excel(path)
    else:
        df = pd.read_csv(path)
    df.columns = [str(c).strip() for c in df.columns]
    return df


def suggest(tx_group: str, key: str) -> str:
    if tx_group in GROUP_DEFAULTS:
        return GROUP_DEFAULTS[tx_group]
    for pattern, cat in RULES:
        if re.search(pattern, key):
            return cat
    return ""


def make(path: str) -> None:
    df = add_groups(load_statement(path), labels_csv="__none__")
    df = df[df["tx_group"] != "fuliza_loan"]
    sheet = (df.groupby("group_id")
               .agg(count=("group_id", "size"),
                    tx_group=("tx_group", lambda s: s.mode().iat[0]),
                    example=("name", "first"))
               .reset_index()
               .sort_values("count", ascending=False))
    sheet["suggested_category"] = [suggest(g, k) for g, k in zip(sheet["tx_group"], sheet["group_id"])]
    sheet["final_category"] = sheet["suggested_category"]
    sheet["needs_review"] = sheet["suggested_category"].eq("").map({True: "YES", False: ""})
    sheet.to_csv(LABELS_CSV, index=False)
    print(f"Wrote {LABELS_CSV}: {len(sheet)} groups "
          f"({(sheet['needs_review'] == 'YES').sum()} need a label).")
    print("Fill in final_category using only these categories:")
    for c in CATEGORIES:
        print("  ", c)


def apply(path: str) -> None:
    if not os.path.exists(LABELS_CSV):
        sys.exit(f"{LABELS_CSV} not found. Run: python build_labels.py make {path}")
    df = add_groups(load_statement(path))
    df = df[df["tx_group"] != "fuliza_loan"]
    df = df[df["category"].fillna("").astype(str).str.strip() != ""]
    df.to_csv(OUT_CSV, index=False)
    print(f"Wrote {OUT_CSV}: {len(df)} labelled rows")
    print(df["category"].value_counts())


if __name__ == "__main__":
    if len(sys.argv) != 3 or sys.argv[1] not in ("make", "apply"):
        sys.exit(__doc__)
    {"make": make, "apply": apply}[sys.argv[1]](sys.argv[2])