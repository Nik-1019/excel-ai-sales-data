"""Build the sales files for the NextWork Excel + Claude projects.

Outputs (all reproducible from SEED):
  messy-sales.xlsx  200 rows, January to June 2026, with planted problems
  clean-sales.xlsx  the same orders cleaned (answer key, and the starter file
                    for "Build a Sales Dashboard with Claude in Excel")
  july-sales.xlsx   one new month in the clean layout (dashboard Secret Mission)
  ANSWER-KEY.md     the numbers a correct clean-up and analysis should find

Run: python make_data.py   (needs openpyxl)
"""
import datetime as dt
import random
from collections import defaultdict

from openpyxl import Workbook
from openpyxl.styles import Font

SEED = 20261117
rng = random.Random(SEED)

PRODUCTS = {
    "Wireless Mouse": 25,
    "Mechanical Keyboard": 89,
    "USB-C Hub": 45,
    "27-inch Monitor": 249,
    "Laptop Stand": 39,
    "Noise-Cancelling Headphones": 179,
}
PRODUCT_WEIGHTS = [26, 16, 20, 12, 16, 10]
REGIONS = ["North", "South", "East", "West"]
# Orders per region per month, January to June. West grows fastest; South dips.
MONTH_PLAN = {
    "North": [8, 8, 8, 8, 8, 8],
    "South": [9, 8, 8, 8, 7, 7],
    "East": [8, 8, 7, 8, 7, 6],
    "West": [6, 6, 8, 9, 10, 12],
}
QTY_CHOICES = [1, 1, 1, 2, 2, 3, 4, 5]


def make_orders(year, months, start_id, plan_for_month):
    orders = []
    for m in months:
        for region in REGIONS:
            for _ in range(plan_for_month(region, m)):
                product = rng.choices(list(PRODUCTS), PRODUCT_WEIGHTS)[0]
                # June: push monitors so the month's top product by sales is clear.
                if m == 6 and rng.random() < 0.10:
                    product = "27-inch Monitor"
                last_day = (dt.date(year, m % 12 + 1, 1) - dt.timedelta(days=1)).day if m < 12 else 31
                day = rng.randint(1, last_day)
                orders.append({
                    "date": dt.date(year, m, day),
                    "product": product,
                    "region": region,
                    "qty": rng.choice([1, 1, 2] if PRODUCTS[product] > 150 else QTY_CHOICES),
                    "price": PRODUCTS[product],
                })
    orders.sort(key=lambda o: o["date"])
    for i, o in enumerate(orders):
        o["id"] = f"SO-{start_id + i}"
    return orders


orders = make_orders(2026, range(1, 7), 1001, lambda r, m: MONTH_PLAN[r][m - 1])
assert len(orders) == 190, len(orders)

# ---- Plant the problems -------------------------------------------------------
messy = [dict(o) for o in orders]
idx = list(range(len(messy)))
rng.shuffle(idx)
take = lambda n: [idx.pop() for _ in range(n)]

date_style = {}
for i in take(30):
    date_style[i] = "mdy_text"   # "Mar 5, 2026"
for i in take(22):
    date_style[i] = "dmy_text"   # "05-Mar-2026"
for i in take(12):
    date_style[i] = "iso_text"   # "2026/03/05"

PRODUCT_VARIANTS = {
    "Wireless Mouse": ["wireless mouse", "Wireless mouse", "Wireless  Mouse"],
    "Mechanical Keyboard": ["Mech Keyboard", "mechanical keyboard"],
    "USB-C Hub": ["USB C Hub", "Usb-C hub", "USBC Hub"],
    "27-inch Monitor": ["27in Monitor", "27\" Monitor"],
    "Laptop Stand": ["laptop stand", "Laptop stand"],
    "Noise-Cancelling Headphones": ["Noise Cancelling Headphones", "NC Headphones"],
}
product_messy = set()
for i in take(26):
    messy[i]["product"] = rng.choice(PRODUCT_VARIANTS[messy[i]["product"]])
    product_messy.add(i)

spaces = set()
for i in take(14):
    if rng.random() < 0.5:
        messy[i]["region"] = " " + messy[i]["region"]
    else:
        messy[i]["region"] = messy[i]["region"] + "  "
    spaces.add(i)
for i in take(4):
    messy[i]["product"] = " " + messy[i]["product"] + " "
    spaces.add(i)

blank_region = take(4)
for i in blank_region:
    messy[i]["region"] = None
blank_price = take(5)
for i in blank_price:
    messy[i]["price"] = None
text_qty = take(4)
for i in text_qty:
    messy[i]["qty"] = str(messy[i]["qty"])

# 10 exact duplicates, inserted right after their original row.
dup_sources = sorted(rng.sample(range(len(messy)), 10))
rows = []
for i, o in enumerate(messy):
    rows.append((i, o))
    if i in dup_sources:
        rows.append((i, o))
assert len(rows) == 200


def fmt_date(i, d):
    style = date_style.get(i)
    if style == "mdy_text":
        return d.strftime("%b ") + str(d.day) + d.strftime(", %Y")
    if style == "dmy_text":
        return d.strftime("%d-%b-%Y")
    if style == "iso_text":
        return d.strftime("%Y/%m/%d")
    return d


HEADERS = ["Order ID", "Order Date", "Product", "Region", "Quantity", "Unit Price"]


def write_sheet(path, header, data_rows, date_col=1, money_cols=(), title="Sales"):
    wb = Workbook()
    ws = wb.active
    ws.title = title
    ws.append(header)
    for c in ws[1]:
        c.font = Font(bold=True)
    for r in data_rows:
        ws.append(r)
    for row in ws.iter_rows(min_row=2):
        cell = row[date_col]
        if isinstance(cell.value, (dt.date, dt.datetime)):
            cell.number_format = "yyyy-mm-dd"
        for mc in money_cols:
            row[mc].number_format = '"$"#,##0.00'
    widths = [11, 13, 28, 9, 10, 11, 12]
    for col, w in zip("ABCDEFG", widths):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A2"
    wb.save(path)


write_sheet(
    "messy-sales.xlsx",
    HEADERS,
    [[o["id"], fmt_date(i, o["date"]), o["product"], o["region"], o["qty"], o["price"]] for i, o in rows],
    money_cols=(5,),
)

# ---- Clean version (answer key) ------------------------------------------------
clean_rows = []
for i, o in enumerate(orders):
    region = "Unknown" if i in blank_region else o["region"]
    clean_rows.append([o["id"], o["date"], o["product"], region, o["qty"], o["price"], o["qty"] * o["price"]])
write_sheet("clean-sales.xlsx", HEADERS + ["Sales"], clean_rows, money_cols=(5, 6))

# ---- July, the new month -------------------------------------------------------
july = make_orders(2026, [7], 1191, lambda r, m: {"North": 9, "South": 7, "East": 8, "West": 15}[r])
write_sheet(
    "july-sales.xlsx",
    HEADERS + ["Sales"],
    [[o["id"], o["date"], o["product"], o["region"], o["qty"], o["price"], o["qty"] * o["price"]] for o in july],
    money_cols=(5, 6),
)

# ---- Answer key ----------------------------------------------------------------
def summarise(rows_):
    by_month, by_region_month, by_product, by_product_june = (defaultdict(int) for _ in range(4))
    for r in rows_:
        m = r[1].month
        by_month[m] += r[6]
        by_region_month[(r[3], m)] += r[6]
        by_product[r[2]] += r[6]
        if m == 6:
            by_product_june[r[2]] += r[6]
    return by_month, by_region_month, by_product, by_product_june


by_month, by_rm, by_product, by_product_june = summarise(clean_rows)
total = sum(r[6] for r in clean_rows)
MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun"]
lines = [
    "# Answer key",
    "",
    f"Seed {SEED}. The messy file has 200 rows; the clean file has {len(clean_rows)}.",
    "",
    "## Planted problems in messy-sales.xlsx",
    "",
    "- 10 exact duplicate rows (same Order ID, right below the original).",
    f"- Mixed date formats: {len(orders) - len(date_style)} real dates, {sum(v=='mdy_text' for v in date_style.values())} like \"Mar 5, 2026\", "
    f"{sum(v=='dmy_text' for v in date_style.values())} like \"05-Mar-2026\", {sum(v=='iso_text' for v in date_style.values())} like \"2026/03/05\" (text, not dates).",
    f"- {len(product_messy)} inconsistent product names (case, spacing, short forms such as \"Mech Keyboard\", \"NC Headphones\", \"27in Monitor\").",
    f"- {len(spaces)} cells with extra spaces before or after the value (Region or Product).",
    "- 4 blank Region cells (can't be worked out: label them Unknown).",
    "- 5 blank Unit Price cells (can be filled from the product's price on other rows).",
    "- 4 Quantity cells stored as text.",
    "",
    "## Prices",
    "",
] + [f"- {p}: ${v}" for p, v in PRODUCTS.items()] + [
    "",
    "## Results on the clean data",
    "",
    f"- Total sales: ${total:,}",
    f"- Orders: {len(clean_rows)}",
    f"- Average order value: ${total / len(clean_rows):,.2f}",
    "- Sales by month: " + ", ".join(f"{MONTHS[m-1]} ${by_month[m]:,}" for m in range(1, 7)),
    "- Sales by product (all months): " + ", ".join(f"{p} ${v:,}" for p, v in sorted(by_product.items(), key=lambda x: -x[1])),
    "- June (last month) by product: " + ", ".join(f"{p} ${v:,}" for p, v in sorted(by_product_june.items(), key=lambda x: -x[1])),
    "",
    "### Region by month",
    "",
    "| Region | " + " | ".join(MONTHS) + " | Jan to Jun change |",
    "|---|" + "---|" * 7,
]
for reg in REGIONS + ["Unknown"]:
    vals = [by_rm[(reg, m)] for m in range(1, 7)]
    change = f"{(vals[5] - vals[0]) / vals[0]:+.0%}" if vals[0] else "n/a"
    lines.append(f"| {reg} | " + " | ".join(f"${v:,}" for v in vals) + f" | {change} |")
july_total = sum(o["qty"] * o["price"] for o in july)
lines += ["", f"## July (july-sales.xlsx)", "", f"- {len(july)} orders, Order IDs SO-1191 onward, total ${july_total:,}"]
open("ANSWER-KEY.md", "w").write("\n".join(lines) + "\n")
print("\n".join(lines))
