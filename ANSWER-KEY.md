# Answer key

Seed 20261117. The messy file has 200 rows; the clean file has 190.

## Planted problems in messy-sales.xlsx

- 10 exact duplicate rows (same Order ID, right below the original).
- Mixed date formats: 126 real dates, 30 like "Mar 5, 2026", 22 like "05-Mar-2026", 12 like "2026/03/05" (text, not dates).
- 26 inconsistent product names (case, spacing, short forms such as "Mech Keyboard", "NC Headphones", "27in Monitor").
- 18 cells with extra spaces before or after the value (Region or Product).
- 4 blank Region cells (can't be worked out: label them Unknown).
- 5 blank Unit Price cells (can be filled from the product's price on other rows).
- 4 Quantity cells stored as text.

## Prices

- Wireless Mouse: $25
- Mechanical Keyboard: $89
- USB-C Hub: $45
- 27-inch Monitor: $249
- Laptop Stand: $39
- Noise-Cancelling Headphones: $179

## Results on the clean data

- Total sales: $26,261
- Orders: 190
- Average order value: $138.22
- Sales by month: Jan $3,355, Feb $4,138, Mar $4,025, Apr $4,383, May $4,869, Jun $5,491
- Sales by product (all months): 27-inch Monitor $6,474, Mechanical Keyboard $6,319, USB-C Hub $4,050, Noise-Cancelling Headphones $3,938, Wireless Mouse $2,750, Laptop Stand $2,730
- June (last month) by product: 27-inch Monitor $2,988, Mechanical Keyboard $712, Wireless Mouse $650, Laptop Stand $468, Noise-Cancelling Headphones $358, USB-C Hub $315

### Region by month

| Region | Jan | Feb | Mar | Apr | May | Jun | Jan to Jun change |
|---|---|---|---|---|---|---|---|
| North | $1,084 | $607 | $969 | $1,033 | $416 | $1,605 | +48% |
| South | $775 | $1,214 | $1,455 | $1,153 | $937 | $1,171 | +51% |
| East | $1,141 | $1,400 | $788 | $1,324 | $1,187 | $768 | -33% |
| West | $355 | $593 | $813 | $873 | $1,971 | $1,698 | +378% |
| Unknown | $0 | $324 | $0 | $0 | $358 | $249 | n/a |

## Dashboard update test (clean-sales.xlsx)

- Change SO-1001 (Laptop Stand) Quantity in cell E2 from 2 to 20: total sales goes from $26,261 to $26,963.

## July (july-sales.xlsx)

- 39 orders, Order IDs SO-1191 onward, total $6,004
- January to July together: 229 orders, total $32,265
