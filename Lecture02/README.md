# Order Quantities When Demand is Approximately Level — Lecture Package
Dr. Islam Ali · Inventory and Production Management

| Item | Purpose |
|---|---|
| `02_Order_Quantities_Level_Demand.ipynb` | Interactive lecture notebook (theory, derivations, visualisations, 🎛️ demos, PDF export, practice problems, industrial tool) |
| `data/` | Input data: `ex1_eoq_basic.csv`, `ex2_epq_production.csv`, `ex3_quantity_discount_cases.csv`, `ex4_sensitivity_scenarios.csv`, `practice_problems.csv` |
| `reports/` | PDF reports exported by the notebook (regenerated every run) |
| `industrial_tool/` | Lot-Sizing Decision Tool: `lot_sizing_tool.py`, `USER_GUIDE.md`, Excel input templates, example Excel + PDF outputs |

## Getting started
```bash
pip install -r requirements.txt
jupyter lab 02_Order_Quantities_Level_Demand.ipynb     # then Run All
```
Google Colab: upload the whole folder (keep the structure), open the notebook, run all cells.

## Notebook map (follows slides 1–29)
0 Setup & PDF engine · 1 Inventory decisions · 2 Deterministic single-item models · 3 EOQ assumptions ·
4 Notation & sawtooth · 5 Cost components / TRC · 6 Derivation (sympy) · 7 T_EOQ & turnover ·
8 Worked example + PDF · 9 Sensitivity (PCP) · 10 EPQ · 11 Quantity discounts (cases 1–3, decision map) ·
12 Practice problems (auto-checker) · 13 Custom solver + full lecture PDF · 14 Industrial tool · 15 Formula sheet

Interactive demos: sawtooth, cost trade-off, EOQ landscape, sensitivity, EPQ, discount negotiator, quiz, universal solver.
Every demo that has a 📄 button writes a PDF to `reports/`.
