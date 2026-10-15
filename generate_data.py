"""
generate_data.py
Creates three deliberately messy raw CSV files (orders, customers, products)
that mimic data collected from multiple systems.

NOTE: This is SYNTHETIC data. Replace data/raw/*.csv with a real Kaggle
dataset (and adjust the column mapping in run_pipeline.py) to get real numbers.
"""
from pathlib import Path

import numpy as np
import pandas as pd

SEED = 42
rng = np.random.default_rng(SEED)
RAW_DIR = Path(__file__).parent / "data" / "raw"

N_ORDERS = 20_000
N_CUSTOMERS = 6_000
N_PRODUCTS = 60

CITIES = ["Kolkata", "Mumbai", "Delhi", "Bengaluru", "Chennai", "Hyderabad", "Pune", "Ahmedabad"]
CITY_WEIGHTS = [0.20, 0.17, 0.17, 0.15, 0.09, 0.09, 0.07, 0.06]
CITY_DIRTY = {
    "Kolkata": ["kolkata", "KOLKATA", "Calcutta", "Kolkatta", " Kolkata "],
    "Mumbai": ["mumbai", "MUMBAI", " Mumbai"],
    "Delhi": ["delhi", "DELHI", "Delhi "],
    "Bengaluru": ["Bangalore", "bangalore", "bengaluru", "BENGALURU"],
    "Chennai": ["chennai", "CHENNAI "],
    "Hyderabad": ["hyderabad", "HYDERABAD"],
    "Pune": ["pune", "PUNE"],
    "Ahmedabad": ["ahmedabad", "Ahmadabad"],
}
CATEGORIES = {
    "Electronics": (1500, 30000),
    "Home & Kitchen": (300, 8000),
    "Fashion": (400, 4500),
    "Books": (150, 900),
    "Beauty": (200, 3000),
    "Sports": (350, 9000),
}
CATEGORY_DIRTY = {
    "Electronics": ["electronics", "Electronics "],
    "Home & Kitchen": ["Home and Kitchen", "home & kitchen"],
    "Fashion": ["fashion", "FASHION"],
    "Books": ["books", "Books "],
    "Beauty": ["beauty"],
    "Sports": ["sports", "Sports "],
}
STATUSES = ["Delivered", "Cancelled", "Returned", "Pending"]
STATUS_P = [0.85, 0.07, 0.05, 0.03]
PAYMENTS = ["UPI", "Credit Card", "Debit Card", "Net Banking", "Cash on Delivery"]
PAY_P = [0.38, 0.22, 0.17, 0.10, 0.13]
PAY_DIRTY = {
    "UPI": ["upi", "UPI "],
    "Credit Card": ["credit card", "CC", "Credit card"],
    "Debit Card": ["debit card", "DC"],
    "Net Banking": ["net banking", "NetBanking"],
    "Cash on Delivery": ["COD", "cash on delivery"],
}


def dirty(value, mapping, rate):
    """Replace a clean label with a dirty variant with probability `rate`."""
    if rng.random() < rate:
        return str(rng.choice(mapping[value]))
    return value


def make_products() -> pd.DataFrame:
    rows = []
    cats = list(CATEGORIES)
    for i in range(1, N_PRODUCTS + 1):
        cat = cats[(i - 1) % len(cats)]
        lo, hi = CATEGORIES[cat]
        price = round(float(rng.uniform(lo, hi)), -1) - 1
        rows.append({"product_id": f"P{i:03d}", "product_name": f"{cat} Item {i}",
                     "category": cat, "list_price": price})
    df = pd.DataFrame(rows)
    # keep a clean copy for generating orders
    clean = df.copy()
    df["category"] = [dirty(c, CATEGORY_DIRTY, 0.30) for c in df["category"]]
    # two products with a missing master price
    df.loc[[7, 33], "list_price"] = np.nan
    return clean, df


def make_customers() -> pd.DataFrame:
    rows = []
    for i in range(1, N_CUSTOMERS + 1):
        city = str(rng.choice(CITIES, p=CITY_WEIGHTS))
        signup = pd.Timestamp("2022-01-01") + pd.Timedelta(days=int(rng.integers(0, 730)))
        rows.append({"customer_id": f"C{i:04d}", "customer_name": f"Customer {i}",
                     "email": f"customer{i}@example.com", "city": city,
                     "signup_date": signup.strftime("%Y-%m-%d")})
    clean = pd.DataFrame(rows)
    df = clean.copy()
    df["city"] = [dirty(c, CITY_DIRTY, 0.25) for c in df["city"]]
    # bad / missing emails
    bad = rng.choice(N_CUSTOMERS, size=int(N_CUSTOMERS * 0.04), replace=False)
    df.loc[bad, "email"] = [f"customer{i+1}at example.com" if k % 2 else f"customer{i+1}@@example"
                            for k, i in enumerate(bad)]
    miss = rng.choice(N_CUSTOMERS, size=int(N_CUSTOMERS * 0.03), replace=False)
    df.loc[miss, "email"] = np.nan
    # duplicate customer rows (with stray whitespace in the id)
    dup = df.sample(40, random_state=SEED).copy()
    dup["customer_id"] = dup["customer_id"] + " "
    df = pd.concat([df, dup], ignore_index=True)
    return clean, df


def make_orders(products_clean, customers_clean) -> pd.DataFrame:
    n = N_ORDERS
    # customer activity is skewed; ~10% of customers never order
    cust_w = rng.lognormal(0, 1.0, N_CUSTOMERS)
    cust_w[rng.choice(N_CUSTOMERS, int(N_CUSTOMERS * 0.10), replace=False)] = 0
    cust_w = cust_w / cust_w.sum()
    days = pd.date_range("2024-01-01", "2025-06-30")
    # seasonality: heavier in Oct-Dec
    w = np.array([1.6 if d.month in (10, 11, 12) else 1.0 for d in days])
    w = w / w.sum()
    dates = rng.choice(days, size=n, p=w)
    prod_idx = rng.integers(0, len(products_clean), size=n)
    qty = rng.choice([1, 2, 3, 4, 5], size=n, p=[0.55, 0.22, 0.12, 0.07, 0.04])
    disc = rng.choice([0, 0.05, 0.10, 0.15, 0.20], size=n, p=[0.40, 0.20, 0.20, 0.12, 0.08])

    df = pd.DataFrame({
        "order_id": [f"ORD-{i:06d}" for i in range(1, n + 1)],
        "customer_id": [f"C{int(c):04d}" for c in rng.choice(np.arange(1, N_CUSTOMERS + 1), size=n, p=cust_w)],
        "product_id": products_clean["product_id"].to_numpy()[prod_idx],
        "order_date": pd.to_datetime(dates),
        "quantity": qty,
        "unit_price": products_clean["list_price"].to_numpy()[prod_idx],
        "discount": disc,
        "status": rng.choice(STATUSES, size=n, p=STATUS_P),
        "payment_method": rng.choice(PAYMENTS, size=n, p=PAY_P),
    })

    # ---- turn everything into dirty strings ----
    out = pd.DataFrame(index=df.index)
    out["order_id"] = df["order_id"]
    out["customer_id"] = df["customer_id"]
    out["product_id"] = df["product_id"]

    # mixed date formats
    fmt_choice = rng.choice(["iso", "dmy", "long"], size=n, p=[0.60, 0.25, 0.15])
    def fmt_date(d, f):
        return {"iso": d.strftime("%Y-%m-%d"), "dmy": d.strftime("%d/%m/%Y"),
                "long": d.strftime("%b %d, %Y")}[f]
    out["order_date"] = [fmt_date(d, f) for d, f in zip(df["order_date"], fmt_choice)]

    out["quantity"] = df["quantity"].astype(object)
    # prices as strings with currency formatting sometimes
    price_str = []
    for p in df["unit_price"]:
        r = rng.random()
        price_str.append(f"{p:.2f}" if r < 0.90 else (f"Rs. {p:,.2f}" if r < 0.95 else f"\u20b9{p:,.0f}"))
    out["unit_price"] = price_str
    out["discount"] = [f"{int(d * 100)}%" if rng.random() < 0.20 else str(d) for d in df["discount"]]
    out["status"] = [dirty(s, {s: [s.lower(), s.upper(), s + " "] for s in STATUSES}, 0.15) for s in df["status"]]
    out["payment_method"] = [dirty(p, PAY_DIRTY, 0.20) for p in df["payment_method"]]

    # ---- inject faults ----
    def pick(frac):
        return rng.choice(n, size=int(n * frac), replace=False)

    out.loc[pick(0.015), "quantity"] = np.nan                  # missing quantity
    out.loc[pick(0.008), "quantity"] = -1                      # negative quantity
    out.loc[pick(0.003), "quantity"] = 999                     # absurd outlier
    out.loc[pick(0.020), "unit_price"] = np.nan                # missing price
    out.loc[pick(0.030), "discount"] = np.nan                  # missing discount
    out.loc[pick(0.003), "discount"] = "95%"                   # impossible discount
    out.loc[pick(0.005), "product_id"] = "P999"                # orphan product
    out.loc[pick(0.005), "customer_id"] = "C9999"              # orphan customer
    out.loc[pick(0.003), "order_date"] = "N/A"                 # unparseable date
    out.loc[pick(0.002), "order_date"] = "31/02/2024"          # impossible date
    out.loc[pick(0.003), "order_date"] = "2031-01-15"          # future date
    # stray whitespace / casing in keys
    idx = pick(0.03)
    out.loc[idx, "customer_id"] = out.loc[idx, "customer_id"].str.lower() + " "
    idx = pick(0.02)
    out.loc[idx, "product_id"] = " " + out.loc[idx, "product_id"].astype(str)

    # duplicate order rows (3%), some with whitespace in the key
    dups = out.sample(int(n * 0.03), random_state=SEED).copy()
    ws = rng.random(len(dups)) < 0.5
    dups.loc[ws, "order_id"] = " " + dups.loc[ws, "order_id"]
    out = pd.concat([out, dups], ignore_index=True).sample(frac=1, random_state=SEED).reset_index(drop=True)
    return out


def main():
    RAW_DIR.mkdir(parents=True, exist_ok=True)
    products_clean, products_raw = make_products()
    customers_clean, customers_raw = make_customers()
    orders_raw = make_orders(products_clean, customers_clean)

    orders_raw.to_csv(RAW_DIR / "orders_raw.csv", index=False)
    customers_raw.to_csv(RAW_DIR / "customers_raw.csv", index=False)
    products_raw.to_csv(RAW_DIR / "products_raw.csv", index=False)
    print(f"orders_raw.csv    : {len(orders_raw):,} rows")
    print(f"customers_raw.csv : {len(customers_raw):,} rows")
    print(f"products_raw.csv  : {len(products_raw):,} rows")


if __name__ == "__main__":
    main()
