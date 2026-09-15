# Sales Tracker — Daily / Weekly / Monthly Auto-Aggregation

A lightweight sales tracker spreadsheet that automatically calculates daily, weekly and monthly totals from raw sale entries — built as part of the **Data Analytics Track internship at Veda Technology**.

Built so a non-technical person can keep adding rows of data and get correct roll-ups without ever touching a formula.

---

## Objective

Practice structuring a spreadsheet for ongoing data entry and automatic aggregation, with a clean separation between the data layer and the calculation layer.

## Tools

- Google Sheets / Microsoft Excel (`.xlsx`, compatible with both)
- Functions: `SUMIFS`, `EOMONTH`, `SUMPRODUCT`, `ROUND`, `IF`
- Data Validation (dropdown list, date rule, numeric rules)
- Freeze panes and colour-coded input cells

---

## Structure

The workbook has two tabs, deliberately separated:

### 1. `Raw Entries` — input layer
| Order Date | Category | Sales | Quantity |
|---|---|---|---|

- Instruction legend at the top of the sheet
- Header row frozen, data starts at row 5
- Shaded entry rows so the user knows exactly where data goes
- **Zero formulas on this tab** — it is purely raw data

### 2. `Summary` — calculation layer

| Block | Rows | Period covered |
|---|---|---|
| Daily Totals | 1,242 | every date present in the data |
| Weekly Totals | 209 | Mon–Sun weeks, Jan 2023 – Jan 2027 |
| Monthly Totals | 48 | Jan 2023 – Dec 2026 |

---

## Formulas

**Daily total** — exact date match:
```
=SUMIFS('Raw Entries'!$C$5:$C$10238, 'Raw Entries'!$A$5:$A$10238, A6)
```

**Weekly total** — between week start and week end:
```
=SUMIFS('Raw Entries'!$C$5:$C$10238,
        'Raw Entries'!$A$5:$A$10238, ">="&D6,
        'Raw Entries'!$A$5:$A$10238, "<="&E6)
```

**Monthly total** — `EOMONTH` handles month length and leap years automatically:
```
=SUMIFS('Raw Entries'!$C$5:$C$10238,
        'Raw Entries'!$A$5:$A$10238, ">="&H6,
        'Raw Entries'!$A$5:$A$10238, "<="&EOMONTH(H6,0))
```

All ranges are absolute (`$`) and extend past the current last row, so newly added entries are picked up automatically.

---

## Data Validation

| Column | Rule | Prevents |
|---|---|---|
| Order Date | valid date, on or after 01-01-2020 | text dates, impossible dates |
| Category | dropdown: Office Supplies / Furniture / Technology | spelling variants that break `SUMIFS` matching |
| Sales | decimal ≥ 0 | negative values, text in a numeric column |
| Quantity | whole number ≥ 0 | fractional / negative quantities |

---

## Verification

Totals were verified two ways.

**1. Independent cross-check** — each sample period is recomputed with `SUMPRODUCT` (a different function from `SUMIFS`) and flagged Match / Mismatch:

| Check | Period | SUMIFS | SUMPRODUCT | Status |
|---|---|---|---|---|
| Sample day | 03-02-2025 | 866.40 | 866.40 | Match |
| Sample week | 30-12-2024 → 05-01-2025 | 8,459.56 | 8,459.56 | Match |
| Sample month | January 2025 | 18,830.33 | 18,830.33 | Match |

**2. Full reconciliation** — all three roll-ups sum to the same grand total as the raw data:

```
Raw entries total    2,326,534.35
Daily totals sum     2,326,534.35
Weekly totals sum    2,326,534.35
Monthly totals sum   2,326,534.35
```

1,508 formulas, zero formula errors.

---

## Sample dataset

- 10,194 sale entries, 3 Jan 2023 → 30 Dec 2026
- No blanks, no negative sales or quantities
- Total value 2,326,534.35 across 38,654 units

**Observations:** Technology leads on revenue (839,893.28 from only 1,865 orders), while Office Supplies leads on order count (6,128 orders, 731,893.31) — high frequency, low ticket size versus low frequency, high ticket size.

---

## Design notes

**Why non-technical staff can use this safely:** one clearly marked place to type, inputs constrained by validation rather than trusted, instructions written inside the file, header row frozen, and totals that update themselves so nobody has to "recalculate" anything by hand.

**Why raw data and formulas are on separate tabs:** inserting, deleting or sorting rows on a mixed sheet silently breaks formula references; formulas get overwritten during ordinary data entry; and exports carry totals mixed in with records. Keeping the raw tab append-only and the summary tab derived is the same principle as separating a database table from a view.

---

## Files

| File | Description |
|---|---|
| `sales_tracker.xlsx` | The tracker (both tabs, formulas, validation, sample data) |
| `Sales_Tracker_Project_Report.pdf` | Full project report |

---

## Limitations / future scope

- Daily Totals covers dates present in the current dataset; a brand-new date needs one row dragged down (weekly and monthly blocks already extend to Jan 2027)
- Category-wise and quantity-wise roll-ups could be added
- A dashboard tab with trend and category charts
- Sheet protection on `Summary` for a true multi-user setup

---

Google sheet link : https://docs.google.com/spreadsheets/d/1NCOgUxx3c8qpibB6H996Gg_YYqExtrEN/edit?usp=drivesdk&ouid=102009503898049232079&rtpof=true&sd=true

**Author:** Sneha Dubey — Integrated MCA, IIPS DAVV Indore
