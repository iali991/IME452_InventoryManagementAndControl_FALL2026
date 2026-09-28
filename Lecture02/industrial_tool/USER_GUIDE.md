# Lot-Sizing Decision Tool — User Guide

**Models:** EOQ (Wilson), EPQ (finite production rate), all-units quantity discounts
**Companion to the lecture:** *Order Quantities When Demand is Approximately Level* (Dr. Islam Ali)
**Version:** 1.0.0

---

## 1. What problem does this tool solve?

For every stocked item (SKU) the tool answers two questions from the lecture:

| Question | Output column |
|---|---|
| **How much** should we order or produce each time? | `Recommended order qty` |
| **When** should we place the order? | `Reorder point (units)` |

It also shows **how much money the current policy wastes** (`Annual savings`, `Current policy PCP (%)`) and **how much freedom you have** to round the quantity (`Flex band: low Q` to `Flex band: high Q`).

Typical uses during an internship or a first job:

* Reviewing the order quantities stored in the ERP/MRP system (SAP, Oracle, Odoo, …) for purchased parts.
* Setting production batch sizes (lot sizes) for items made in-house.
* Evaluating a supplier's quantity-discount offer: *"Should we buy 1,000 units to get 4% off?"*
* Building a savings business case for management (the PDF report is ready to present).

---

## 2. Folder contents

```
industrial_tool/
├── lot_sizing_tool.py                 ← the tool (one file, no installation needed)
├── USER_GUIDE.md                      ← this guide
├── input/
│   ├── blank_input_template.xlsx      ← empty template to fill with your own data
│   └── example_company_input.xlsx     ← complete worked example (23 items, fictional company)
└── output/
    ├── example_company_results.xlsx   ← results produced from the example input
    └── example_company_report.pdf     ← management report produced from the example input
```

---

## 3. Installation (one time)

You need Python 3.9+ and four common packages:

```bash
pip install numpy pandas matplotlib openpyxl
```

---

## 4. Quick start (3 steps)

```bash
# 1) Create a blank template (or copy input/blank_input_template.xlsx)
python lot_sizing_tool.py --template input/my_plant_input.xlsx

# 2) Fill the template in Excel (see Section 5), save it.

# 3) Run the analysis
python lot_sizing_tool.py --input input/my_plant_input.xlsx --outdir output
```

Output: `output/my_plant_results.xlsx` and `output/my_plant_report.pdf`.

Other options:

| Option | Effect |
|---|---|
| `--no-pdf` | Only the Excel results (fast, for very large item lists) |
| `--sku-pages 20` | PDF detail pages only for the 20 SKUs with the largest savings |
| `--version` | Print the tool version |

From Jupyter / Python:

```python
import lot_sizing_tool as lst
res = lst.run("input/example_company_input.xlsx", "output")
res["results"].head()      # pandas DataFrame with every number
res["kpis"]                # portfolio totals
```

---

## 5. Filling the input workbook

The workbook has five sheets. **Blue headers are required, grey headers are optional.** Hover over an `Items` header in Excel to see its explanation.

### 5.1 `Settings`

| setting | example | meaning |
|---|---|---|
| company_name | Delta Pumps & Valves | Printed on the report |
| report_title | Lot-Sizing Recommendations FY2027 | Printed on the report |
| analyst | Your name | Printed on the report |
| currency | $ | Symbol used in outputs |
| days_per_year | 250 | Working days per year; converts lead times and cycle times |
| default_carrying_rate | 0.24 | r used for items with an empty `carrying_rate` |
| flex_tolerance_pct | 2 | Cost tolerance (%) that defines the flexibility band |

### 5.2 `Items` — one row per SKU

| column | required | unit | notes |
|---|---|---|---|
| sku | ✔ | – | Unique code |
| description | | – | |
| category | | – | Any grouping (raw material, spare part, …) |
| supply_type | ✔ | Purchased / Produced | Purchased → EOQ; Produced → EPQ |
| annual_demand | ✔ | units/year | **D**. Use a forecast or the last 12 months of consumption |
| order_cost | ✔ | currency/order | **A**. Fixed cost per purchase order or per production setup |
| unit_cost | ✔ | currency/unit | **v** (base price **v₀** when price breaks exist) |
| carrying_rate | | $/$/year | **r**. Blank = default. Enter 0.24, not 24 |
| production_rate | Produced only | units/year | **m** — output rate while the line runs; must exceed D |
| lead_time_days | | working days | Used for the reorder point |
| pack_size | | units | Q must be a multiple of it (carton, pallet, batch) |
| min_order_qty | | units | Supplier or process MOQ |
| max_order_qty | | units | Storage space, shelf life or budget limit |
| current_order_qty | | units | Today's order size — enables the savings comparison |

### 5.3 `PriceBreaks` — optional, one row per break

| sku | min_qty | unit_cost |
|---|---|---|
| RM-1002 | 500 | 36.50 |
| RM-1002 | 1000 | 35.20 |

*All-units discounts:* if Q ≥ `min_qty`, the **whole order** is charged `unit_cost`. The base price (from `Items`) applies below the first break. Several breaks per SKU are allowed.

### 5.4 Where to find the numbers in a real company

| Parameter | Where to look | Tip |
|---|---|---|
| D | ERP consumption history, sales forecast, MRP gross requirements | Check that demand is roughly level; a strongly seasonal or declining item needs a time-varying lot-sizing method (Silver–Meal, Wagner–Whitin) |
| A (purchased) | Purchasing + receiving + inspection + accounts-payable labour per PO, fixed freight charge | Typically $20–$150 per PO |
| A (produced) | Setup/changeover time × (labour + machine hourly rate) + scrap during setup | Setup reduction projects (SMED) reduce A and therefore Q |
| v | Standard cost (produced) or purchase price (purchased) | |
| r | Finance: cost of capital (WACC) + storage + insurance + obsolescence | 0.15–0.35 per year is typical |
| m | Line or machine output rate × available hours per year | Must be larger than D |
| L | Supplier's confirmed lead time or the production lead time | |

---

## 6. What the tool calculates

**Total annual cost** for any order size Q:

```
Total(Q) = A·D/Q            (ordering / setup)
         + (Q/2)·k·v(Q)·r   (carrying,  k = 1 for EOQ,  k = 1 − D/m for EPQ)
         + D·v(Q)           (purchase, matters only with quantity discounts)
```

**Formulas from the lecture**

| Quantity | Formula |
|---|---|
| EOQ | √(2AD / vr) |
| EPQ | √(2AD / (vr(1 − D/m))) = EOQ / √(1 − D/m) |
| TRC(EOQ) | √(2ADvr) |
| Months of supply | T = 12·Q/D = √(288A / Dvr) at the EOQ |
| Turnover ratio | TR = D / (Q/2) = √(2vrD / A) at the EOQ |
| Percentage cost penalty | PCP = 50·α² / (1 + α), with Q′ = (1 + α)·EOQ |

**Quantity-discount procedure (generalised from the lecture's 3-step procedure)**

1. For each price tier compute the EOQ/EPQ at that tier's price.
2. If it lies inside the tier it is a candidate; if it is below the tier's break quantity, the break quantity is the candidate (the cost curve is convex inside a tier).
3. Compare all candidates on **total** cost (including D·v) and keep the cheapest.

With a single break this gives exactly the lecture's **Case 1** (order Q₁), **Case 2** (ignore the discount) and **Case 3** (EOQ at the discounted price).

**Practical rules applied after the theory:** the best quantity is rounded up and down to multiples of `pack_size`, kept between `min_order_qty` and `max_order_qty`, and all feasible candidates are compared on total cost.

**Reorder point** = D × L / days_per_year: place an order when the *inventory position* (on hand + on order − backorders) reaches this level.

**Flexibility band**: all Q around the optimum whose cost is at most `flex_tolerance_pct` % above the optimal TRC. Because the cost curve is flat near the optimum, rounding inside the band to full pallets, truckloads or shifts costs almost nothing.

---

## 7. Reading the outputs

### 7.1 Excel results (`*_results.xlsx`)

| Sheet | Content |
|---|---|
| README | How to act on the results |
| Summary | Portfolio KPIs: cost, savings, inventory investment, number of orders |
| Recommendations | One row per SKU (key columns highlighted green). Filters are enabled |
| Column_Guide | Meaning and formula of every column |
| Validation_Log | Input problems. **ERROR** = item skipped, **WARNING** = item analysed with an assumption |
| Input_Items, Input_PriceBreaks | Copy of the analysed data (audit trail) |

### 7.2 PDF report (`*_report.pdf`)

1. **Cover**: KPI tiles (savings, inventory investment, orders per year).
2. **How to read this report**: method and limitations in plain language.
3. **Portfolio charts**: savings by SKU, PCP of current policies, order frequency vs inventory trade-off, inventory investment, cycle-time distribution, cost components by model.
4. **Summary table** of all recommendations.
5. **Validation log.**
6. **One page per SKU**: cost curve with price tiers (solid = applicable price, dotted = extension, as in the lecture figures), current vs recommended point, flexibility band, inventory sawtooth with reorder point, TRC breakdown, key figures.

---

## 8. Worked example (included)

`input/example_company_input.xlsx` describes the fictional company **Delta Pumps & Valves Co.** with 23 items:
14 purchased raw materials and parts (8 with supplier price breaks), 7 produced finished goods and sub-assemblies (EPQ), and 2 rows with deliberate data mistakes that show the validation log at work:

* `RM-1013` has `carrying_rate = 22` (a percentage typed by mistake). The tool interprets it as 0.22 and logs a WARNING.
* `FG-3005` has a production rate lower than its demand. That is impossible for EPQ, so the tool skips the item and logs an ERROR.

Things to look for in the results:

* **RM-1008** (electric motor): the theoretical optimum is the 200-unit price break, but the storage limit `max_order_qty = 150` wins, so the practical rule `Max qty` is applied.
* **RM-1003 / RM-1004**: lecture **Case 1**. It pays to raise the order to the break quantity.
* **RM-1001 / RM-1002 / RM-1005**: several price breaks; the tool picks the best tier.
* **FG-3003**: an EPQ item where today's batch (2,000) is far too large; the PCP is about 50%.

---

## 9. Assumptions and when NOT to use the tool

The tool inherits the EOQ assumptions from the lecture:

1. Demand is level and known. Do not use it for highly seasonal, promotional, trending or lumpy demand; use time-varying lot sizing for those.
2. Costs do not change over time (low inflation, no planned price changes).
3. Items are independent. Joint replenishment (ordering many items from one supplier together) is not modelled.
4. No shortages are planned. The reorder point contains **no safety stock**; add safety stock separately for uncertain demand or lead time.
5. The discount is all-units. For *incremental* discounts, the price schedule has to be converted first.

**Rule of thumb:** because of the flat cost curve (PCP), a ±20% error in any one of D, A, v or r raises the cost by less than 1%. Spend effort on getting the data roughly right, not perfect.

---

## 10. Troubleshooting

| Message | Cause | Fix |
|---|---|---|
| `Items sheet is missing required column(s)` | Header renamed or deleted | Start again from the blank template |
| `Missing or non-positive [...] - item skipped` | Empty D, A or v | Fill the value |
| `production_rate must exceed annual_demand` | m ≤ D | Check units (both per year) or switch to Purchased |
| `carrying_rate > 1` | Percentage typed | Enter it as a fraction (0.24) |
| `Break ... not lower than previous tier` | Price breaks out of order | Prices must fall as `min_qty` increases |
| `max_order_qty is below MOQ` | Conflicting limits | Correct one of them |
