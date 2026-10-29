"""
Food Price Trends & Inflation: Nigeria, Nov 2024 - Apr 2025
pandas + matplotlib version of food_trend_dashboard_2024_2025.xlsx

Run:   python food_trend_analysis.py [path/to/workbook.xlsx]
Needs: pandas, numpy, matplotlib, openpyxl
Output: printed tables + PNG charts in ./charts/ + CSV tables in ./tables/

The script works in Jupyter / VS Code too (cells are marked with  # %% ).
"""

# %% 1. Setup -----------------------------------------------------------------
import sys
from pathlib import Path

import matplotlib.pyplot as plt
import matplotlib.ticker as mtick
import numpy as np
import pandas as pd

XLSX = sys.argv[1] if len(sys.argv) > 1 else "food_trend_dashboard_2024_2025.xlsx"
CHART_DIR, TABLE_DIR = Path("charts"), Path("tables")
CHART_DIR.mkdir(exist_ok=True)
TABLE_DIR.mkdir(exist_ok=True)

MONTHS = ["Nov-2024", "Dec-2024", "Jan-2025", "Feb-2025", "Mar-2025", "Apr-2025"]

# Chart styling: one accent for emphasis, grey for context, light recessive grid.
BLUE, ORANGE, GREY = "#2a78d6", "#eb6834", "#b8b7b2"
INK, INK_2 = "#0b0b0b", "#52514e"
plt.rcParams.update({
    "figure.facecolor": "#fcfcfb", "axes.facecolor": "#fcfcfb",
    "axes.edgecolor": "#d9d8d3", "axes.labelcolor": INK_2,
    "xtick.color": INK_2, "ytick.color": INK_2, "text.color": INK,
    "axes.spines.top": False, "axes.spines.right": False,
    "axes.grid": True, "grid.color": "#ececE8", "grid.linewidth": 0.8,
    "axes.axisbelow": True, "axes.titleweight": "bold", "axes.titlesize": 12,
    "axes.titlelocation": "left", "font.size": 10, "figure.dpi": 110,
})


def save(fig, name):
    fig.tight_layout()
    fig.savefig(CHART_DIR / f"{name}.png", dpi=160, bbox_inches="tight")
    plt.close(fig)


def naira(x, _=None):
    return f"₦{x:,.0f}"


# %% 2. Load & clean -----------------------------------------------------------
# 'Standardized Data' is the base table; every other sheet in the workbook
# (pivots, Item Summary, State Summary, Top Movers) is a formula over it.
df = pd.read_excel(XLSX, sheet_name="Standardized Data").rename(columns={
    "Standardized Item": "item",
    "Match Confidence": "confidence",
    "Avg (Same Month Last Year)": "price_ly",
    "Avg (Previous Month)": "price_prev",
    "Avg (Current Month)": "price",
    "MoM": "mom",
    "YoY": "yoy",
    "Highest Price": "high_price",
    "Highest State": "high_state",
    "Lowest Price": "low_price",
    "Lowest State": "low_state",
})
df["month"] = pd.Categorical(df["Month"], categories=MONTHS, ordered=True)
df["month_dt"] = pd.to_datetime(df["Month"], format="%b-%Y")
df = df.drop(columns=["Month", "Highest", "Lowest", "Source File"])

print(f"{len(df)} rows | {df['item'].nunique()} items | {df['month'].nunique()} months")
print(df["confidence"].value_counts().to_string(), "\n")

# Items present in all 6 months (the workbook's 33 "continuous" items).
months_covered = df.groupby("item")["month"].nunique()
continuous = months_covered[months_covered == len(MONTHS)].index
print(f"Continuously tracked items: {len(continuous)}\n")

# %% 3. Item summary (sheet 'Item Summary') ------------------------------------
price_wide = df.pivot_table(index="item", columns="month", values="price", observed=False)
item_summary = (
    df.groupby("item")
    .agg(confidence=("confidence", "first"),
         months_covered=("month", "nunique"),
         avg_mom_pct=("mom", "mean"),
         avg_yoy_pct=("yoy", "mean"))
    .join(price_wide.add_prefix("price_"))
)
item_summary["latest_price"] = price_wide.ffill(axis=1).iloc[:, -1]
item_summary.to_csv(TABLE_DIR / "item_summary.csv")

# %% 4. National basket trend (Read Me section 1) ------------------------------
basket = (
    df[df["item"].isin(continuous)]
    .groupby("month", observed=True)["price"].sum()
    .to_frame("basket_total")
)
basket["mom_pct"] = basket["basket_total"].pct_change() * 100
total_change = basket["basket_total"].iloc[-1] / basket["basket_total"].iloc[0] - 1
monthly_cagr = (1 + total_change) ** (1 / (len(MONTHS) - 1)) - 1
print("=== 1. National basket (33 continuous items) ===")
print(basket.round(1).to_string())
print(f"Nov -> Apr: {total_change:.1%}  (~{monthly_cagr:.1%} per month compounded, five month-to-month steps)\n")
basket.to_csv(TABLE_DIR / "basket_trend.csv")

fig, ax = plt.subplots(figsize=(8, 4.2))
ax.plot(basket.index.astype(str), basket["basket_total"], color=BLUE, lw=2, marker="o", ms=6,
        markeredgecolor="#fcfcfb", markeredgewidth=2)
for m, v in basket["basket_total"].items():
    ax.annotate(naira(v), (str(m), v), textcoords="offset points", xytext=(0, 9),
                ha="center", fontsize=9, color=INK)
ax.set_ylim(basket["basket_total"].min() * 0.93, basket["basket_total"].max() * 1.06)
ax.yaxis.set_major_formatter(mtick.FuncFormatter(naira))
ax.set_title(f"Food basket up {total_change:.1%} in five months, with a dip in April")
ax.set_ylabel("Basket total, 33 items")
save(fig, "01_national_basket_trend")

# %% 5. Inflation by food group (Read Me section 2) ----------------------------
# The workbook doesn't store food groups; this mapping reproduces the Read Me groups.
CATEGORY_RULES = [
    ("Staples (rice, beans, garri, yam, plantain)",
     ["rice", "beans", "garri", "yam", "plantain", "maize", "semovita", "wheat flour", "potato"]),
    ("Fish / seafood",
     ["fish", "mackerel", "sardine", "titus", "tilapia", "catfish", "mudfish", "cray"]),
    ("Meat / poultry / eggs",
     ["beef", "chicken", "goat", "egg"]),
    ("Cooking oils", ["oil"]),
    ("Processed / dairy", ["milk", "bread", "groundnuts, roasted"]),
    ("Vegetables / fruit",
     ["onion", "tomato", "carrot", "ginger", "watermelon"]),
]


def categorise(name: str) -> str:
    n = name.lower()
    for cat, keys in CATEGORY_RULES:
        if any(k in n for k in keys):
            return cat
    return "Other"


df["category"] = df["item"].map(categorise)
unmapped = df.loc[df["category"] == "Other", "item"].unique()
if len(unmapped):
    print("Items without a category:", list(unmapped))

# NOTE: the Read Me says "as of April 2025", but its figures (99.7, 87.5, 74.4 ...)
# are the mean YoY over ALL monthly rows, Nov-2024 to Apr-2025. Reproduced here.
cat_yoy = (df.groupby("category")["yoy"].mean().sort_values(ascending=False)
           .rename("avg_yoy_pct").to_frame())
print("=== 2. Avg YoY inflation by food group, Nov-2024 to Apr-2025 ===")
print(cat_yoy.round(1).to_string(), "\n")
cat_yoy.to_csv(TABLE_DIR / "category_yoy.csv")

fig, ax = plt.subplots(figsize=(8, 3.8))
order = cat_yoy.sort_values("avg_yoy_pct")
colors = [BLUE if "Staples" in c else GREY for c in order.index]
bars = ax.barh(order.index, order["avg_yoy_pct"], color=colors, height=0.62)
ax.bar_label(bars, fmt="%.0f%%", padding=4, fontsize=9)
ax.set_xlim(0, order["avg_yoy_pct"].max() * 1.15)
ax.xaxis.set_major_formatter(mtick.PercentFormatter(decimals=0))
ax.grid(axis="y", visible=False)
ax.set_title("Staples roughly doubled year-on-year (avg Nov 2024 – Apr 2025)")
save(fig, "02_inflation_by_food_group")

# %% 6. Which items drive inflation / volatility (Read Me section 3) -----------
top_yoy = item_summary["avg_yoy_pct"].dropna().nlargest(10).rename("avg_yoy_pct").to_frame()
top_mom = item_summary["avg_mom_pct"].nlargest(10).rename("avg_mom_pct").to_frame()
vol = (df.groupby("item")["mom"].std().dropna().sort_values(ascending=False)
       .rename("mom_std_pct").to_frame())
print("=== 3a. Top 10 items by average YoY % ===")
print(top_yoy.round(1).to_string())
print("\n=== 3b. Top 10 items by average MoM % ===")
print(top_mom.round(2).to_string())
print("\n=== 3c. Most volatile items (std of MoM %) ===")
print(vol.head(5).round(1).to_string(), "\n")
for name, t in [("top10_yoy", top_yoy), ("top10_mom", top_mom), ("volatility", vol)]:
    t.to_csv(TABLE_DIR / f"{name}.csv")


def ranked_bar(series, title, fname, color=BLUE, fmt="%.0f%%", highlight=None):
    s = series.sort_values()
    fig, ax = plt.subplots(figsize=(8, 4.6))
    cols = [color if (highlight is None or i in highlight) else GREY for i in s.index]
    bars = ax.barh(s.index, s.values, color=cols, height=0.62)
    ax.bar_label(bars, fmt=fmt, padding=4, fontsize=9)
    ax.set_xlim(0, s.max() * 1.15)
    ax.grid(axis="y", visible=False)
    ax.set_title(title)
    save(fig, fname)


ranked_bar(top_yoy["avg_yoy_pct"], "Top 10 items by average year-on-year inflation (%)",
           "03_top10_yoy_inflation")
ranked_bar(top_mom["avg_mom_pct"], "Top 10 items by average month-on-month increase (%)",
           "04_top10_mom_increase", color=ORANGE, fmt="%.1f%%")

# %% 7. MoM / YoY trend for each item (Dashboard "MOM/YOY BY ITEMS") -----------
mom_wide = df.pivot_table(index="item", columns="month", values="mom", observed=False)
yoy_wide = df.pivot_table(index="item", columns="month", values="yoy", observed=False)


def plot_item(item: str, save_it: bool = True):
    """Price, MoM and YoY for one item: replaces the Excel slicer.
    Three panels (not one shared axis) because MoM is a few % and YoY is 100%+."""
    if item not in price_wide.index:
        raise KeyError(f"{item!r} not found. Try one of: {list(price_wide.index)[:5]} ...")
    x = np.arange(len(MONTHS))
    fig, axes = plt.subplots(1, 3, figsize=(13, 3.9))

    a = axes[0]
    a.plot(x, price_wide.loc[item].values, color=BLUE, lw=2, marker="o", ms=6,
           markeredgecolor="#fcfcfb", markeredgewidth=2)
    a.yaxis.set_major_formatter(mtick.FuncFormatter(naira))
    a.set_title("Average price")

    for a, series, colour, title in [
        (axes[1], mom_wide.loc[item], ORANGE, "Month-on-month change"),
        (axes[2], yoy_wide.loc[item], BLUE, "Year-on-year change"),
    ]:
        vals = series.values.astype(float)
        a.bar(x, np.nan_to_num(vals), 0.6, color=colour)
        a.axhline(0, color=INK_2, lw=0.8)
        a.yaxis.set_major_formatter(mtick.PercentFormatter(decimals=0))
        a.set_title(title)
        for xi, v in zip(x, vals):
            if np.isnan(v):
                a.text(xi, 0, "n/a", ha="center", va="bottom", fontsize=8, color=INK_2)

    for a in axes:
        a.set_xticks(x, MONTHS, rotation=35, ha="right")
    fig.suptitle(item, x=0.01, ha="left", fontweight="bold", fontsize=13)
    if yoy_wide.loc[item].isna().any():
        fig.text(0.01, -0.02, "n/a = no comparable prior-year or prior-month price "
                 "(item list changed in Jan-2025).", fontsize=8, color=INK_2)
    if save_it:
        safe = "".join(c if c.isalnum() else "_" for c in item).strip("_")
        save(fig, f"item_{safe}")
    return fig


for it in ["Local Rice (Broken)", "Yam Tuber", "Onions, fresh"]:
    plot_item(it)

# Price index (Nov-2024 = 100) for the headline staples: comparable across items
staples = ["Local Rice (Broken)", "Beans Brown", "Plantain (unripe)", "Yam Tuber"]
index100 = price_wide.loc[staples].div(price_wide.loc[staples, "Nov-2024"], axis=0) * 100
fig, ax = plt.subplots(figsize=(8.5, 4.4))
palette = [BLUE, ORANGE, "#1baf7a", "#6250d6"]
for (name, row), c in zip(index100.iterrows(), palette):
    ax.plot(MONTHS, row.values, color=c, lw=2, marker="o", ms=5,
            markeredgecolor="#fcfcfb", markeredgewidth=1.5, label=name)
    ax.annotate(name, (MONTHS[-1], row.values[-1]), textcoords="offset points",
                xytext=(8, 0), va="center", fontsize=9, color=INK)
ax.axhline(100, color=INK_2, lw=0.8)
ax.set_xlim(-0.2, len(MONTHS) - 1 + 1.6)
ax.set_ylabel("Price index (Nov 2024 = 100)")
ax.set_title("Headline staples: price index since November")
ax.legend(frameon=False, loc="upper left")
save(fig, "05_staples_price_index")

# %% 8. Geography: state price gaps (Read Me section 4) ------------------------
df["state_gap"] = df["high_price"] / df["low_price"] - 1          # 1.0 = 100% gap
gap_by_item = (df.groupby("item")["state_gap"].mean().sort_values(ascending=False)
               .rename("avg_state_gap").to_frame())
print("=== 4a. State price gap (highest-price state vs lowest-price state) ===")
print(f"Mean gap {df['state_gap'].mean():.0%} | median gap {df['state_gap'].median():.0%}")
print(gap_by_item.head(5).round(2).to_string(), "\n")
gap_by_item.to_csv(TABLE_DIR / "state_gap_by_item.csv")

ranked_bar(gap_by_item["avg_state_gap"].head(10) * 100,
           "Items with the widest gap between cheapest and dearest state (avg %)",
           "06_state_price_gap", color=BLUE)

state = (
    pd.concat([
        df.groupby("high_state")["high_price"].agg(times_highest="size", avg_price_when_highest="mean"),
        df.groupby("low_state")["low_price"].agg(times_lowest="size", avg_price_when_lowest="mean"),
    ], axis=1)
    .fillna({"times_highest": 0, "times_lowest": 0})
    .astype({"times_highest": int, "times_lowest": int})
    .rename_axis("state")
)
state.to_csv(TABLE_DIR / "state_summary.csv")

NORTH_EAST = ["Yobe", "Borno", "Adamawa", "Taraba", "Gombe", "Bauchi"]
ne = state.loc[state.index.isin(NORTH_EAST)]
print("=== 4b. North-east states ===")
print(f"Lowest-price state {ne['times_lowest'].sum()} times | "
      f"highest-price state {ne['times_highest'].sum()} times\n")

top_high = state["times_highest"].nlargest(10)
top_low = state["times_lowest"].nlargest(10)

fig, axes = plt.subplots(1, 2, figsize=(11, 4.6), sharex=True)
for ax, s, title, col in [
    (axes[0], top_high.sort_values(), "Most often the highest-price state", ORANGE),
    (axes[1], top_low.sort_values(), "Most often the lowest-price state", BLUE),
]:
    bars = ax.barh(s.index, s.values, color=col, height=0.62)
    ax.bar_label(bars, padding=3, fontsize=9)
    ax.grid(axis="y", visible=False)
    ax.set_title(title)
save(fig, "07_states_highest_lowest")

# %% 9. Month-on-month heatmap (all continuous items) --------------------------
heat = mom_wide.loc[continuous, MONTHS[1:]]
heat = heat.loc[heat.abs().max(axis=1).sort_values(ascending=False).index]
fig, ax = plt.subplots(figsize=(7.5, 9))
lim = 20
im = ax.imshow(heat.values, cmap="RdBu_r", vmin=-lim, vmax=lim, aspect="auto")
ax.set_xticks(range(heat.shape[1]), heat.columns)
ax.set_yticks(range(heat.shape[0]), heat.index, fontsize=8)
ax.grid(False)
for i in range(heat.shape[0]):
    for j in range(heat.shape[1]):
        v = heat.values[i, j]
        if not np.isnan(v):
            ax.text(j, i, f"{v:.1f}", ha="center", va="center", fontsize=7,
                    color="white" if abs(v) > 12 else INK)
cb = fig.colorbar(im, ax=ax, shrink=0.5, pad=0.02)
cb.set_label("MoM change (%)")
ax.set_title("Month-on-month price change by item")
save(fig, "08_mom_heatmap")

# %% 10. Data-quality note (Read Me section 5) ---------------------------------
print("=== Data caveat ===")
print(df.drop_duplicates("item")["confidence"].value_counts().to_string())
print("\nDiscontinued/new items break YoY comparability after Jan-2025; "
      "footnote any external use of MoM/YoY for 'Likely' and 'New (2025)' items.")
print(f"\nCharts saved to {CHART_DIR.resolve()}\nTables saved to {TABLE_DIR.resolve()}")
