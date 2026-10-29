# Food Price Trends & Inflation: Nigeria, Nov 2024 – Apr 2025

A pandas + matplotlib version of the Excel workbook `food_trend_dashboard_2024_2025.xlsx`.
It rebuilds the workbook's pivots, summary sheets, Read Me findings and dashboard charts in
Python, so the analysis can be re-run, audited and extended without Excel.

Data source: Selected Food Price Watch, standardized across the January 2025 item-list change.

## Quick start

```bash
pip install pandas numpy matplotlib openpyxl
python food_trend_analysis.py food_trend_dashboard_2024_2025.xlsx
```

If no path is given, the script looks for `food_trend_dashboard_2024_2025.xlsx` in the current folder.
It prints the key tables to the terminal and writes:

- `charts/`: PNG charts
- `tables/`: CSV summary tables

The script also runs in Jupyter or VS Code. Each section starts with a `# %%` cell marker.

## Project layout

```
food_trend_analysis.py                    main script
food_trend_dashboard_2024_2025.xlsx       input workbook
charts/                                   generated PNGs
tables/                                   generated CSVs
```

## How the workbook maps to the code

Only the `Standardized Data` sheet is read. Every other sheet (pivot tables, Item Summary,
State Summary, Top Movers) is a formula over it, so pandas recomputes them directly.

| Excel | pandas |
|---|---|
| Standardized Data | `df` (254 rows, 52 items, 6 months) |
| Item Summary (AVERAGEIFS / SUMIFS) | `item_summary` via `groupby("item")` and `pivot_table` |
| State Summary (COUNTIF / AVERAGEIF) | `state` via `groupby("high_state" / "low_state")` |
| Pivot: price trend by item | `price_wide` |
| Pivot: MoM / YoY by item | `mom_wide`, `yoy_wide` |
| Top Movers | `top_yoy`, `top_mom` and the state tables |
| Dashboard item slicer | `plot_item("Item name")` |

## Outputs

### Charts (`charts/`)

| File | Shows |
|---|---|
| `01_national_basket_trend.png` | Basket total for the 33 continuous items, by month |
| `02_inflation_by_food_group.png` | Average YoY inflation by food group (staples highlighted) |
| `03_top10_yoy_inflation.png` | Top 10 items by average YoY % |
| `04_top10_mom_increase.png` | Top 10 items by average MoM % |
| `05_staples_price_index.png` | Rice, beans, plantain and yam, indexed to Nov 2024 = 100 |
| `06_state_price_gap.png` | Items with the widest cheapest-to-dearest state gap |
| `07_states_highest_lowest.png` | States most often the highest and lowest priced |
| `08_mom_heatmap.png` | MoM % change for every continuous item |
| `item_<name>.png` | Price, MoM and YoY panels for one item (3 examples generated) |

### Tables (`tables/`)

`basket_trend.csv`, `category_yoy.csv`, `item_summary.csv`, `top10_yoy.csv`, `top10_mom.csv`,
`volatility.csv`, `state_gap_by_item.csv`, `state_summary.csv`

## Chart any item

This replaces the Excel slicer. In a notebook or Python session:

```python
plot_item("Yam Tuber")                       # displays and saves charts/item_Yam_Tuber.png
plot_item("Onions, fresh", save_it=False)    # display only
```

Item names must match the `Standardized Item` column exactly. A wrong name raises a `KeyError`
that suggests valid names. `list(price_wide.index)` shows all 52.

## Verified results

The script reproduces the figures in the workbook's Read Me:

| Check | Result |
|---|---|
| Continuous items | 33 |
| Basket, Nov 2024 to Apr 2025 | ₦84,423 to ₦95,788 (+13.5%) |
| April dip | ₦97,158 (Mar) to ₦95,788 (Apr), the only MoM fall |
| Staples avg YoY | 99.7% |
| Fish, dairy, meat, oils, veg/fruit | 87.5%, 74.4%, 73.8%, 71.9%, 62.1% |
| State price gap | mean 123%, median 100% |
| North-east states | lowest-price 91 times, highest-price 34 times |
| Most volatile items | Onions (std 26.7%), egg crate (17.2%) |

## Methodology notes

- **Continuous items.** Only items present in all six months go into the basket total. Items
  discontinued or introduced in January 2025 are excluded so the basket is like for like.
- **Food groups are inferred.** The workbook has no category column, so `CATEGORY_RULES`
  assigns groups by keywords in the item name. This reproduces the Read Me values, but review
  the rules if you add items. Unmatched items print as "Items without a category".
- **Food-group averages cover all months.** The Read Me describes them as "as of April 2025",
  but its figures are the mean YoY over all monthly rows, Nov 2024 to Apr 2025. The code
  matches the workbook, and the chart title says so.
- **Monthly compounding.** Growth of 13.5% over five month-to-month steps is about 2.6% per
  month. The Read Me rounds this to "roughly 2.5%".
- **State price gap** = highest-state price ÷ lowest-state price − 1, per item per month.
  1.0 means the dearest state costs double the cheapest.

## Data caveats

- **Item list change (Jan 2025).** Ten items were discontinued and nine introduced. Only the 27
  items marked `Confirmed` are fully reliable for YoY comparisons beyond January.
- **Missing YoY / MoM.** Several `Likely` items have no YoY after January (for example Local
  Rice (Broken)). They appear as "n/a" in the per-item charts rather than as zeros.
- **January 2025 spikes.** Large jumps for plantain, tilapia, catfish, mackerel and titus
  (+15% to +37% MoM) probably reflect the unit or brand change, not pure price movement.
  Footnote this in any external use.
- **April dip.** One month does not make a trend. Describe it as "stabilizing" until May and
  June data confirm the direction.
- **Match confidence.** Labels are `Confirmed`, `Likely`, `New (2025)` or `Discontinued`.
  Treat `Likely` matches with more caution.

## Extending

- **New months:** append rows to `Standardized Data` and add the labels to `MONTHS`.
- **New items:** add keywords to `CATEGORY_RULES` so they land in a food group.
- **Other states or regions:** change `NORTH_EAST` in the geography section.
