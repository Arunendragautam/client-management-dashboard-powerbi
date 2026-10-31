"""Creates a deliberately MESSY Excel file (raw_client_data.xlsx) so the cleaning step has real work to do."""
import random
import pandas as pd

random.seed(42)
first = ["Aarav","Vihaan","Isha","Priya","Rohan","Sneha","Kabir","Ananya","Aditya","Meera","Nikhil","Pooja","Rahul","Divya","Sahil"]
last = ["Sharma","Patil","Deshmukh","Joshi","Kulkarni","Mehta","Iyer","Nair","Singh","Gupta"]
cities = ["Pune","Mumbai","Nagpur","Nashik","Pimpri-Chinchwad"]
industries = ["Retail","IT Services","Healthcare","Education","Manufacturing"]
products = ["Consulting","Software License","Support Plan","Training","Hardware"]

clients = []
for i in range(1, 61):
    name = f"{random.choice(first)} {random.choice(last)}"
    clients.append({
        "ClientID": f"C{i:03d}",
        "ClientName": name,
        "Email": f"{name.lower().replace(' ', '.')}{i}@example.com",
        "Phone": f"+91 9{random.randint(100000000, 999999999)}",
        "City": random.choice(cities),
        "Industry": random.choice(industries),
        "SignupDate": f"{random.randint(2022, 2025)}-{random.randint(1, 12):02d}-{random.randint(1, 28):02d}",
        "Status": random.choice(["Active", "Inactive"]),
    })

# ---- inject mess ----
messy = []
for c in clients:
    c = c.copy()
    r = random.random()
    if r < 0.15: c["ClientName"] = "  " + c["ClientName"].upper() + " "          # case + spaces
    elif r < 0.30: c["ClientName"] = c["ClientName"].lower()
    if random.random() < 0.2: c["City"] = random.choice([c["City"].lower(), c["City"].upper() + " ", "Pimpri Chinchwad", "pune "])
    if random.random() < 0.2: c["Status"] = random.choice(["active", "ACTIVE ", "inactive", "Inactv"])
    if random.random() < 0.15: c["Industry"] = random.choice([None, "N/A", ""])      # missing
    if random.random() < 0.10: c["Email"] = None
    if random.random() < 0.10: c["Phone"] = random.choice([c["Phone"].replace("+91 ", ""), "98-76-54-32-10", "N/A"])
    if random.random() < 0.3:                                                       # mixed date formats
        y, m, d = c["SignupDate"].split("-")
        c["SignupDate"] = random.choice([f"{d}/{m}/{y}", f"{d}-{m}-{y}", f"{int(d)} {['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'][int(m)-1]} {y}"])
    messy.append(c)
# duplicates (same client entered twice, slightly different)
for c in random.sample(messy, 8):
    d = c.copy(); d["ClientName"] = d["ClientName"].strip().title() + " "; messy.append(d)
random.shuffle(messy)

txns = []
for t in range(1, 301):
    cid = f"C{random.randint(1, 60):03d}"
    if t % 97 == 0: cid = "C999"                                                     # orphan client
    amt = random.randint(5, 250) * 1000
    amt_txt = random.choice([f"₹{amt:,}", f"{amt}", f"Rs. {amt:,}", amt])
    if t % 41 == 0: amt_txt = None                                                   # missing amount
    txns.append({
        "TxnID": f"T{t:04d}", "ClientID": cid,
        "TxnDate": f"{random.randint(2024, 2026)}-{random.randint(1, 12):02d}-{random.randint(1, 28):02d}",
        "Product": random.choice(products), "Amount": amt_txt,
    })
txns += random.sample(txns, 5)                                                      # duplicate transactions

with pd.ExcelWriter("raw_client_data.xlsx") as xw:
    pd.DataFrame(messy).to_excel(xw, sheet_name="Clients", index=False)
    pd.DataFrame(txns).to_excel(xw, sheet_name="Transactions", index=False)
print("raw_client_data.xlsx created:", len(messy), "client rows,", len(txns), "transaction rows")
