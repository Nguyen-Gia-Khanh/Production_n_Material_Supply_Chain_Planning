# Publishing notebook tables in the README

Export the **DataFrame behind the IPython display** as an HTML table. GitHub can render the table inside a collapsible `<details>` section, keeping the numbers selectable and searchable. No screenshot is needed.

GitHub sanitizes rendered HTML, including custom inline styles and CSS identifiers, so the notebook's colors and layout styling do not transfer reliably. Use the notebook for the styled view and the README for the table's content. See [GitHub's rendering pipeline](https://github.com/github/markup) and [collapsed-section documentation](https://docs.github.com/en/get-started/writing-on-github/working-with-advanced-formatting/organizing-information-with-collapsed-sections).

## Use the helper

`export_readme_table` in `utils/util_display.py` replaces only the section between an existing pair of table markers. It exports every row, keeps row and column index labels, and splits wide tables into separate collapsible parts. Each part repeats the row index so it can be read independently.

`columns_per_table=5` means five **data columns** per part. Index levels do not count toward that limit. For a date/shift MultiIndex, each date/shift combination is one data column. Reduce the number when headers are long.

Run these examples from the repository root **after the corresponding report cells have completed with a usable solution**. They update your local `README.md`; commit and push that file to update GitHub. Keep all exported sections from the same planning scenario.

### From notebook 1: regular production

```python
from utils.util_display import export_readme_table

run_note = (
    f"Regular horizon: {calendar_dates[0].date()} to "
    f"{calendar_dates[-1].date()}; solver status: {solve_status}; "
    f"configured relative gap: {SOLVER_RELATIVE_GAP:.1%}."
)

export_readme_table(
    forecast_attainment_table,
    section_id="regular-attainment",
    title="Regular forecast attainment",
    formats={
        "Forecast": "{:,.0f}",
        "Finished": "{:,.0f}",
        "Unscheduled": "{:,.0f}",
        "Attainment": "{:.1%}",
        "Profit per Product": "{:,.2f}",
    },
    note=run_note,
)

export_readme_table(
    regular_schedule_table,
    section_id="regular-schedule",
    title="Regular production by line, shift and day",
    formats="{:,.0f}",
    columns_per_table=5,
    note=run_note,
)
```

### From notebook 2: overtime

```python
from utils.util_display import export_readme_table

export_readme_table(
    ot_attainment_table,
    section_id="ot-attainment",
    title="Overtime forecast recovery",
    formats={
        "Original Forecast": "{:,.0f}",
        "Solved Regular": "{:,.0f}",
        "For OT": "{:,.0f}",
        "Solved OT": "{:,.0f}",
        "Remaining After OT": "{:,.0f}",
        "OT Attainment": "{:.1%}",
    },
    note=(
        f"OT horizon: {calendar_dates[0].date()} to "
        f"{calendar_dates[-1].date()}; solver status: {solve_status}; "
        f"configured relative gap: {SOLVER_RELATIVE_GAP:.1%}."
    ),
)
```

### From notebook 3: material orders and warehouse occupation

```python
from utils.util_display import export_readme_table, order_format

export_readme_table(
    order_material_first[[
        "Arrival Day", "Order Quantity", "Containers",
        "Discount", "Discounted Unit Price", "Purchase Subtotal",
    ]],
    section_id="material-orders",
    title="Material order plan",
    formats=order_format,
    note=f"Material planning: {num_days} days; solver status: {pl.LpStatus[model.status]}.",
)

export_readme_table(
    df_warehouse,
    section_id="warehouse",
    title="Warehouse occupation at the end of each day",
    note=(
        "This report shows ending inventory occupation. The model checks "
        "warehouse capacity after arrivals, before daily consumption."
    ),
)
```

The material-order example deliberately selects six summary columns. Remove that selection to export all columns; the helper will split them across parts.

After publishing the first snapshots, edit the introductory note under **Selected result tables** to describe the scenario and remove the statement that no numerical snapshot has been published.

## Add another table

Add a unique marker pair at the desired location in `README.md`:

```markdown
<!-- BEGIN TABLE:regular-wip -->
<!-- END TABLE:regular-wip -->
```

Then export the relevant DataFrame from its notebook:

```python
export_readme_table(
    wip_carryover_table,
    section_id="regular-wip",
    title="Regular WIP carryover",
    formats="{:,.0f}",
    columns_per_table=4,
)
```

Existing report variables include:

| Notebook | DataFrame | Content |
|---|---|---|
| Regular | `line_utilization_table` | Utilization and capacity signals by line and shift. |
| Regular | `remaining_time_table` | Remaining time by line, date and shift. |
| Regular | `wip_carryover_table` | Movement through process queues. |
| Regular | `ending_wip_snapshot` | Closing WIP for the next planning period. |
| OT | `overtime_schedule_table` | OT quantity by line, product and date. |
| OT | `ot_labor_table` | Worker-time, utilization and labor cost. |
| OT | `ot_wip_carryover_table` | WIP in the separate OT flow. |
| Material | `inventory_material_first` | Arrivals, requirements and inventory by material and day. |

The helper accepts a DataFrame. For a variable holding a pandas Styler, use its `.data` attribute. Do not pass `display(...)`, which returns no table data. Re-running the helper replaces the same marked section; it does not append duplicate snapshots or execute a solver.

## One-off export without the helper

For a flat table, [pandas `to_markdown`](https://pandas.pydata.org/docs/reference/api/pandas.DataFrame.to_markdown.html) is enough. It requires the optional `tabulate` package:

```bash
python -m pip install tabulate
```

```python
print(forecast_attainment_table.to_markdown())
```

Copy that output into the README. For hierarchical rows or columns, [pandas `to_html`](https://pandas.pydata.org/docs/reference/api/pandas.DataFrame.to_html.html) retains the table structure:

```python
print(regular_schedule_table.iloc[:, :5].to_html(sparsify=False))
```

Paste the resulting `<table>...</table>` directly inside a `<details>` block, with blank lines around it. Do not wrap the actual table in a fenced code block, or GitHub will display its source instead of rendering it.
