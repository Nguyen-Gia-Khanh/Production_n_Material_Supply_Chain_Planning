# Production & Material Supply Chain Planning

An Excel-driven planning project that converts a product forecast into a daily production schedule, evaluates additional output through overtime, and plans material purchases against a fixed production plan. The optimization models use **Python, pandas, PuLP and HiGHS**, with Jupyter notebooks for analysis and reporting.

The project connects practical factory planning decisions: which products to prioritize when capacity is limited, where work in progress accumulates, whether overtime is worthwhile, and when materials should arrive without exceeding warehouse capacity.

## Planning workflow

| Stage | Notebook | Main decisions and outputs |
|---|---|---|
| 1. Regular production | [1_production_scheduling.ipynb](1_production_scheduling.ipynb) | Product quantities by line, shift and day; completed output; unmet forecast; utilization; process-level WIP. |
| 2. Overtime | [2_overtime_scheduling.ipynb](2_overtime_scheduling.ipynb) | Additional finished output, production by line and day, workers per 30-minute OT block, and OT labor cost. |
| 3. Material planning | [3_material_planning.ipynb](3_material_planning.ipynb) | Order dates, quantities, containers, supplier discount tiers, arrivals, inventory and warehouse occupation. |

The regular notebook writes `Solved_Regular` and `For_OT` to the production workbook. The overtime notebook uses `For_OT = Forecast_Amount - Solved_Regular` as its demand limit.

**The material handoff is currently manual:** notebook 3 reads `Production_Schedule` from `xlconfigs/material_input.xlsm`. Prepare that dated plan from the production scenario you want to supply, including OT if applicable, before running material planning. The three stages are solved sequentially; material feasibility does not automatically revise the production schedule.

## What the models consider

### Regular production

- Integer production quantities, capped by each product's forecast.
- Product routes defined by the ordered process columns in `POL_Matrix`; zero processing time skips a process.
- Several physical lines can serve one process, each with its own staffed shifts and dated capacity.
- Person-time per unit, stored in milliseconds, matched against available worker-time. A date with zero capacity cannot receive production.
- Minimum product batches by line and day. Any unmet forecast must be zero or at least the batch threshold derived from the route's available lines.
- WIP balances between consecutive processes and an optional opening WIP snapshot from `Last_MIP`.
- A soft 70% time-utilization policy: below the threshold, the objective charges the salary value of all unused available worker-time.

**Objective:** maximize finished-product profit minus the under-utilization salary penalty.

### Overtime

- Demand is limited to the quantity left after the regular solve.
- Only staffed `Day` lines with registered availability in `OT_Total` can use overtime.
- Integer worker counts are selected for 30-minute blocks within the registered duration. Counts may decrease across consecutive blocks.
- Processing must use at least 70% of the worker-time actually scheduled for OT.
- Minimum batches, process routes and WIP balances still apply.
- OT starts with zero opening WIP and processes its output through the full route. It does not consume regular-production WIP.

**Objective:** maximize additional finished-product profit minus OT labor cost. The configured OT pay multiplier is `1.5`.

### Material planning

- A bill of materials converts the dated product plan into daily material requirements.
- Supplier lead times determine order arrival days.
- Order activation, minimum/maximum quantities, integer container counts and one selected discount tier per order.
- Inventory balances and minimum safety stock for every material and day.
- Warehouse capacity checked after receiving arrivals, before that day's consumption.
- Input safeguards check a single material's daily space requirement against an 80% warehouse-share allowance and check opening stock coverage before the earliest supplier arrival.

**Objective:** minimize purchasing, container shipping and inventory holding costs. Material order and inventory quantities are continuous; production quantities and container counts are integer.

## Files and inputs

| File | Purpose |
|---|---|
| [xlconfigs/production_input.xlsm](xlconfigs/production_input.xlsm) | Product forecasts, routes, process times, line settings and opening WIP. |
| [xlconfigs/labor_input.xlsm](xlconfigs/labor_input.xlsm) | Staffing, shift information, regular availability and registered OT. |
| [xlconfigs/material_input.xlsm](xlconfigs/material_input.xlsm) | Material settings, BOM, supplier discounts and the dated production plan to supply. |
| [utils/util_display.py](utils/util_display.py) | Notebook table formatting and export of result tables into this README. |
| [utils/util_material_input_validation.py](utils/util_material_input_validation.py) | Material input safeguards and readable explanations of violations. |
| [utils/util_export_schedule_2xl.py](utils/util_export_schedule_2xl.py) | Helper for writing a solved finished-production schedule to an existing Excel template. |

<details>
<summary>Workbook sheets and input conventions</summary>

| Workbook | Sheets read by the notebooks |
|---|---|
| Production | `Master_List`, `Product_Forecast`, `POL_Matrix`, `Line_Config`, `Product_Config`; optional `Last_MIP`. |
| Labor | `Expanded_Details`, `Work_Total`, `OT_Total`. |
| Material | `Master_Lists`, `Production_Schedule`, `BOM_Matrix`, `Material_Config`, `Sales_Program`. |

`POL_Matrix` holds **person-milliseconds per unit**, indexed by product and process. A line such as `Pri1` takes its processing standard from process `Pri`. Excel availability stays in hours; Python converts it to milliseconds.

The current notebooks read the salary column named `Hrly_Sal` and divide it by `Max_Hr × 3,600,000`. Despite the column name, this calculation treats the value as pay per worker for the full `Max_Hr` period. Match the input's units to that calculation when preparing a scenario.

`Work_Total` and `OT_Total` use Excel row 3 as their header, with nine descriptive columns before the dated columns. Preserve the supplied layout and use chronological dates.

`Last_MIP` has four columns: `As_Of_Date`, `Product_List`, `Process`, `WIP_Units`. Its snapshot date must be the day before the regular horizon starts. WIP belongs to a product/process queue; lines supply capacity.

The material schedule uses dates on row 3, products in column A from row 4, and quantities starting at B4. Keep its product order aligned with `Master_Lists`. Material lead times advance through schedule columns, so the calendar must include the days that should count toward lead time.

</details>

## Getting started

Create an environment from the repository root:

```bash
git clone https://github.com/Nguyen-Gia-Khanh/Production_n_Material_Supply_Chain_Planning.git
cd Production_n_Material_Supply_Chain_Planning
python -m venv .venv
```

Activate it with `.venv\Scripts\Activate.ps1` in Windows PowerShell, or `source .venv/bin/activate` on macOS/Linux. Then install the notebook dependencies:

```bash
python -m pip install pandas pulp highspy openpyxl jupyterlab ipykernel jinja2
python -m jupyterlab
```

1. Prepare the three workbooks in `xlconfigs/`. Recalculate and save formula-driven inputs in Excel before loading them; `openpyxl` does not calculate Excel formulas. Close the production workbook before Python writes to it.
2. Open notebook 1 with the project root as the working directory. Review its settings, run cells in order, and inspect the solver status and reports.
3. Run notebook 2 after notebook 1 has written `Solved_Regular` and `For_OT`. Check that `OT_Total` represents the intended planning horizon.
4. Prepare the approved dated production plan in `material_input.xlsm` and run notebook 3. Resolve any input-safeguard failure before solving.
5. Use the [table export guide](docs/display-tables.md) to publish selected results below as text tables.

Notebook 1 also writes its ending WIP back to `Last_MIP`. For repeated comparisons of the **same horizon**, preserve the original opening snapshot and skip the final WIP writeback cell until ready to advance the plan. That final cell currently sets `WRITE_MIP_SNAPSHOT = True` itself.

The regular and OT notebooks each allow up to **1,800 seconds** with a **1% relative MIP gap**. Material planning uses **300 seconds** and **0.5%**. These are solver settings, not measured runtimes or guarantees of exact optimality; read the solver log when interpreting a result.

## Selected result tables

The notebooks produce forecast attainment, detailed schedules, utilization, WIP, OT labor, material orders and warehouse reports. The committed notebook outputs are cleared, so no solved numerical snapshot is published here yet. Run the [export examples](docs/display-tables.md) after solving to fill these sections with actual results.

<!-- BEGIN TABLE:regular-attainment -->
<details>
<summary>Regular forecast attainment</summary>

A solved snapshot can be exported from `forecast_attainment_table`.

</details>
<!-- END TABLE:regular-attainment -->

<!-- BEGIN TABLE:regular-schedule -->
<details>
<summary>Regular production by line, shift and day</summary>

A solved snapshot can be exported from `regular_schedule_table`.

</details>
<!-- END TABLE:regular-schedule -->

<!-- BEGIN TABLE:ot-attainment -->
<details>
<summary>Overtime forecast recovery</summary>

A solved snapshot can be exported from `ot_attainment_table`.

</details>
<!-- END TABLE:ot-attainment -->

<!-- BEGIN TABLE:material-orders -->
<details>
<summary>Material order plan</summary>

A solved snapshot can be exported from `order_material_first`.

</details>
<!-- END TABLE:material-orders -->

<!-- BEGIN TABLE:warehouse -->
<details>
<summary>Warehouse occupation</summary>

A solved snapshot can be exported from `df_warehouse`. This report shows ending occupation; the model's capacity constraint checks opening occupation after arrivals.

</details>
<!-- END TABLE:warehouse -->

## Scope and assumptions

This is a planning prototype with deterministic forecast inputs and processing standards. Forecasting itself is outside this repository. `Profit_Per_Product` is a supplied planning parameter rather than a complete accounting calculation.

Production flow is aggregated by day. Same-day transfer permits downstream processing against that day's upstream output; it does not establish exact within-day start times or a minute-by-minute executable sequence. Sequence-dependent setup times and machine breakdown uncertainty are not modeled.

Ending WIP is allowed by default in both production stages. The OT model keeps its WIP separate, and the material model consumes BOM quantities on the dates in its supplied schedule. Stage-specific material consumption and a joint optimization of production, OT and procurement would require extending the current models.
