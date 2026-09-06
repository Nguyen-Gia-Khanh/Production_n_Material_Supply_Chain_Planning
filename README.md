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
5. Review the [saved report exports](reports/). To refresh them, save the executed notebooks with their outputs, export their displayed tables, and replace the matching report files and README sections.

Notebook 1 also writes its ending WIP back to `Last_MIP`. For repeated comparisons of the **same horizon**, preserve the original opening snapshot and skip the final WIP writeback cell until ready to advance the plan. That final cell currently sets `WRITE_MIP_SNAPSHOT = True` itself.

The regular and OT notebooks each allow up to **1,800 seconds** with a **1% relative MIP gap**. Material planning uses **300 seconds** and **0.5%**. These are solver settings, not measured runtimes or guarantees of exact optimality; read the solver log when interpreting a result.

## Selected result tables

These are saved notebook displays from the supplied planning runs. Table contents, date columns, grouped index labels and number formats are retained from the exports. Expand a section to view a complete table.

| Full report | Included displays |
|---|---|
| [Regular production](reports/1_production_scheduling.tables.md) | 13 tables: opening WIP, objective, finished output, forecast attainment, schedules, capacity, utilization and closing WIP. |
| [Overtime](reports/2_overtime_scheduling.tables.md) | 7 tables: objective, attainment, finished output, line schedule, labor requirements and WIP. |
| [Material planning](reports/3_material_planning.tables.md) | 5 tables: orders by material and by day, inventory by material and by day, and warehouse occupation. |

The complete exported files are also available in [reports/](reports/). Their included styling is retained in the source; GitHub applies its own display styling.

The full regular report's **finished-product daily schedule** already contains `...` in the supplied export. That display hides some middle dates. The detailed regular line schedule below contains all 30 date columns and its Total column.

<!-- BEGIN TABLE:regular-attainment -->

<details>
<summary>FORECAST ATTAINMENT</summary>

<style type="text/css">
</style>
<table id="T_3c6ee">
  <thead>
    <tr>
      <th class="blank level0" >&nbsp;</th>
      <th id="T_3c6ee_level0_col0" class="col_heading level0 col0" >Forecast</th>
      <th id="T_3c6ee_level0_col1" class="col_heading level0 col1" >Finished</th>
      <th id="T_3c6ee_level0_col2" class="col_heading level0 col2" >Unscheduled</th>
      <th id="T_3c6ee_level0_col3" class="col_heading level0 col3" >Attainment</th>
      <th id="T_3c6ee_level0_col4" class="col_heading level0 col4" >Profit per Product</th>
    </tr>
    <tr>
      <th class="index_name level0" >Product</th>
      <th class="blank col0" >&nbsp;</th>
      <th class="blank col1" >&nbsp;</th>
      <th class="blank col2" >&nbsp;</th>
      <th class="blank col3" >&nbsp;</th>
      <th class="blank col4" >&nbsp;</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <th id="T_3c6ee_level0_row0" class="row_heading level0 row0" >A</th>
      <td id="T_3c6ee_row0_col0" class="data row0 col0" >3,000,000</td>
      <td id="T_3c6ee_row0_col1" class="data row0 col1" >46,590</td>
      <td id="T_3c6ee_row0_col2" class="data row0 col2" >2,953,410</td>
      <td id="T_3c6ee_row0_col3" class="data row0 col3" >1.6%</td>
      <td id="T_3c6ee_row0_col4" class="data row0 col4" >500.00</td>
    </tr>
    <tr>
      <th id="T_3c6ee_level0_row1" class="row_heading level0 row1" >B</th>
      <td id="T_3c6ee_row1_col0" class="data row1 col0" >3,100,000</td>
      <td id="T_3c6ee_row1_col1" class="data row1 col1" >3,100,000</td>
      <td id="T_3c6ee_row1_col2" class="data row1 col2" >0</td>
      <td id="T_3c6ee_row1_col3" class="data row1 col3" >100.0%</td>
      <td id="T_3c6ee_row1_col4" class="data row1 col4" >700.00</td>
    </tr>
    <tr>
      <th id="T_3c6ee_level0_row2" class="row_heading level0 row2" >C</th>
      <td id="T_3c6ee_row2_col0" class="data row2 col0" >4,566,000</td>
      <td id="T_3c6ee_row2_col1" class="data row2 col1" >4,566,000</td>
      <td id="T_3c6ee_row2_col2" class="data row2 col2" >0</td>
      <td id="T_3c6ee_row2_col3" class="data row2 col3" >100.0%</td>
      <td id="T_3c6ee_row2_col4" class="data row2 col4" >900.00</td>
    </tr>
    <tr>
      <th id="T_3c6ee_level0_row3" class="row_heading level0 row3" >D</th>
      <td id="T_3c6ee_row3_col0" class="data row3 col0" >295,000</td>
      <td id="T_3c6ee_row3_col1" class="data row3 col1" >295,000</td>
      <td id="T_3c6ee_row3_col2" class="data row3 col2" >0</td>
      <td id="T_3c6ee_row3_col3" class="data row3 col3" >100.0%</td>
      <td id="T_3c6ee_row3_col4" class="data row3 col4" >3,000.00</td>
    </tr>
    <tr>
      <th id="T_3c6ee_level0_row4" class="row_heading level0 row4" >E</th>
      <td id="T_3c6ee_row4_col0" class="data row4 col0" >1,233,000</td>
      <td id="T_3c6ee_row4_col1" class="data row4 col1" >209,891</td>
      <td id="T_3c6ee_row4_col2" class="data row4 col2" >1,023,109</td>
      <td id="T_3c6ee_row4_col3" class="data row4 col3" >17.0%</td>
      <td id="T_3c6ee_row4_col4" class="data row4 col4" >600.00</td>
    </tr>
    <tr>
      <th id="T_3c6ee_level0_row5" class="row_heading level0 row5" >F</th>
      <td id="T_3c6ee_row5_col0" class="data row5 col0" >1,234,000</td>
      <td id="T_3c6ee_row5_col1" class="data row5 col1" >1,232,000</td>
      <td id="T_3c6ee_row5_col2" class="data row5 col2" >2,000</td>
      <td id="T_3c6ee_row5_col3" class="data row5 col3" >99.8%</td>
      <td id="T_3c6ee_row5_col4" class="data row5 col4" >800.00</td>
    </tr>
    <tr>
      <th id="T_3c6ee_level0_row6" class="row_heading level0 row6" >G</th>
      <td id="T_3c6ee_row6_col0" class="data row6 col0" >134,500</td>
      <td id="T_3c6ee_row6_col1" class="data row6 col1" >0</td>
      <td id="T_3c6ee_row6_col2" class="data row6 col2" >134,500</td>
      <td id="T_3c6ee_row6_col3" class="data row6 col3" >0.0%</td>
      <td id="T_3c6ee_row6_col4" class="data row6 col4" >450.00</td>
    </tr>
    <tr>
      <th id="T_3c6ee_level0_row7" class="row_heading level0 row7" >H</th>
      <td id="T_3c6ee_row7_col0" class="data row7 col0" >1,234,000</td>
      <td id="T_3c6ee_row7_col1" class="data row7 col1" >0</td>
      <td id="T_3c6ee_row7_col2" class="data row7 col2" >1,234,000</td>
      <td id="T_3c6ee_row7_col3" class="data row7 col3" >0.0%</td>
      <td id="T_3c6ee_row7_col4" class="data row7 col4" >120.00</td>
    </tr>
    <tr>
      <th id="T_3c6ee_level0_row8" class="row_heading level0 row8" >I</th>
      <td id="T_3c6ee_row8_col0" class="data row8 col0" >8,765,000</td>
      <td id="T_3c6ee_row8_col1" class="data row8 col1" >8,763,000</td>
      <td id="T_3c6ee_row8_col2" class="data row8 col2" >2,000</td>
      <td id="T_3c6ee_row8_col3" class="data row8 col3" >100.0%</td>
      <td id="T_3c6ee_row8_col4" class="data row8 col4" >1,000.00</td>
    </tr>
  </tbody>
</table>


</details>

<!-- END TABLE:regular-attainment -->

<!-- BEGIN TABLE:regular-schedule -->

<details>
<summary>REGULAR PRODUCTION SCHEDULE</summary>

<style type="text/css">
#T_1666d th.col_heading {
  background-color: #D9E1F2;
  color: #000000;
  font-weight: bold;
  padding: 6px;
  border: 1px solid #A6A6A6;
}
#T_1666d th.row_heading.level0 {
  background-color: #FFD966;
  color: #000000;
  font-weight: bold;
  padding: 6px;
  border-right: 2px solid #7F6000;
}
#T_1666d td.row0 {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row0:not(.level0) {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d .row0 {
  border-top: 3px solid #595959;
}
#T_1666d td.row1 {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row1:not(.level0) {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row2 {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row2:not(.level0) {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row3 {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row3:not(.level0) {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row4 {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row4:not(.level0) {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row5 {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row5:not(.level0) {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row6 {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row6:not(.level0) {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row7 {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row7:not(.level0) {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row8 {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row8:not(.level0) {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row9 {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row9:not(.level0) {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d .row9 {
  border-top: 3px solid #595959;
}
#T_1666d td.row10 {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row10:not(.level0) {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row11 {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row11:not(.level0) {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row12 {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row12:not(.level0) {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row13 {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row13:not(.level0) {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row14 {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row14:not(.level0) {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row15 {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row15:not(.level0) {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row16 {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row16:not(.level0) {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row17 {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row17:not(.level0) {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d .row17 {
  border-top: 3px solid #595959;
}
#T_1666d td.row18 {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row18:not(.level0) {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row19 {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row19:not(.level0) {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row20 {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row20:not(.level0) {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row21 {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row21:not(.level0) {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row22 {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row22:not(.level0) {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row23 {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row23:not(.level0) {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row24 {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row24:not(.level0) {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row25 {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row25:not(.level0) {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row26 {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row26:not(.level0) {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d .row26 {
  border-top: 3px solid #595959;
}
#T_1666d td.row27 {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row27:not(.level0) {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row28 {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row28:not(.level0) {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row29 {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row29:not(.level0) {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row30 {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row30:not(.level0) {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d .row30 {
  border-top: 3px solid #595959;
}
#T_1666d td.row31 {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row31:not(.level0) {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row32 {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row32:not(.level0) {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row33 {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row33:not(.level0) {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row34 {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row34:not(.level0) {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row35 {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row35:not(.level0) {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row36 {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row36:not(.level0) {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row37 {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row37:not(.level0) {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row38 {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row38:not(.level0) {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row39 {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row39:not(.level0) {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d .row39 {
  border-top: 3px solid #595959;
}
#T_1666d td.row40 {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row40:not(.level0) {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row41 {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row41:not(.level0) {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row42 {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row42:not(.level0) {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row43 {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row43:not(.level0) {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row44 {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row44:not(.level0) {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row45 {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row45:not(.level0) {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row46 {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row46:not(.level0) {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row47 {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row47:not(.level0) {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d .row47 {
  border-top: 3px solid #595959;
}
#T_1666d td.row48 {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row48:not(.level0) {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row49 {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row49:not(.level0) {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row50 {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row50:not(.level0) {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row51 {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row51:not(.level0) {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row52 {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row52:not(.level0) {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d .row52 {
  border-top: 3px solid #595959;
}
#T_1666d td.row53 {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row53:not(.level0) {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row54 {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row54:not(.level0) {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row55 {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row55:not(.level0) {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row56 {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row56:not(.level0) {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row57 {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row57:not(.level0) {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row58 {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row58:not(.level0) {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d .row58 {
  border-top: 3px solid #595959;
}
#T_1666d td.row59 {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row59:not(.level0) {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row60 {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row60:not(.level0) {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row61 {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row61:not(.level0) {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row62 {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row62:not(.level0) {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row63 {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row63:not(.level0) {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row64 {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row64:not(.level0) {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d .row64 {
  border-top: 3px solid #595959;
}
#T_1666d td.row65 {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row65:not(.level0) {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row66 {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row66:not(.level0) {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row67 {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row67:not(.level0) {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row68 {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row68:not(.level0) {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d .row68 {
  border-top: 3px solid #595959;
}
#T_1666d td.row69 {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row69:not(.level0) {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row70 {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row70:not(.level0) {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row71 {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row71:not(.level0) {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row72 {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row72:not(.level0) {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row73 {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row73:not(.level0) {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row74 {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row74:not(.level0) {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row75 {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row75:not(.level0) {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row76 {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row76:not(.level0) {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d .row76 {
  border-top: 3px solid #595959;
}
#T_1666d td.row77 {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row77:not(.level0) {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row78 {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row78:not(.level0) {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row79 {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row79:not(.level0) {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row80 {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row80:not(.level0) {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row81 {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row81:not(.level0) {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row82 {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row82:not(.level0) {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row83 {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row83:not(.level0) {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row84 {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row84:not(.level0) {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row85 {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row85:not(.level0) {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row86 {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row86:not(.level0) {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row87 {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row87:not(.level0) {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row88 {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row88:not(.level0) {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row89 {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row89:not(.level0) {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row90 {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row90:not(.level0) {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row91 {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row91:not(.level0) {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row92 {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row92:not(.level0) {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row93 {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row93:not(.level0) {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d .row93 {
  border-top: 3px solid #595959;
}
#T_1666d td.row94 {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row94:not(.level0) {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row95 {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row95:not(.level0) {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row96 {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row96:not(.level0) {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row97 {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row97:not(.level0) {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row98 {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row98:not(.level0) {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row99 {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row99:not(.level0) {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row100 {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row100:not(.level0) {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row101 {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row101:not(.level0) {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row102 {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row102:not(.level0) {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row103 {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row103:not(.level0) {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row104 {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row104:not(.level0) {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row105 {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row105:not(.level0) {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row106 {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row106:not(.level0) {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row107 {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row107:not(.level0) {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row108 {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row108:not(.level0) {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row109 {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row109:not(.level0) {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row110 {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row110:not(.level0) {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row111 {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row111:not(.level0) {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d .row111 {
  border-top: 3px solid #595959;
}
#T_1666d td.row112 {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row112:not(.level0) {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row113 {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row113:not(.level0) {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row114 {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row114:not(.level0) {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row115 {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row115:not(.level0) {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row116 {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row116:not(.level0) {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d td.row117 {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_1666d  th.row_heading.row117:not(.level0) {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
</style>
<table id="T_1666d">
  <thead>
    <tr>
      <th class="blank" >&nbsp;</th>
      <th class="blank" >&nbsp;</th>
      <th class="blank level0" >&nbsp;</th>
      <th id="T_1666d_level0_col0" class="col_heading level0 col0" >2026-09-01</th>
      <th id="T_1666d_level0_col1" class="col_heading level0 col1" >2026-09-02</th>
      <th id="T_1666d_level0_col2" class="col_heading level0 col2" >2026-09-03</th>
      <th id="T_1666d_level0_col3" class="col_heading level0 col3" >2026-09-04</th>
      <th id="T_1666d_level0_col4" class="col_heading level0 col4" >2026-09-05</th>
      <th id="T_1666d_level0_col5" class="col_heading level0 col5" >2026-09-06</th>
      <th id="T_1666d_level0_col6" class="col_heading level0 col6" >2026-09-07</th>
      <th id="T_1666d_level0_col7" class="col_heading level0 col7" >2026-09-08</th>
      <th id="T_1666d_level0_col8" class="col_heading level0 col8" >2026-09-09</th>
      <th id="T_1666d_level0_col9" class="col_heading level0 col9" >2026-09-10</th>
      <th id="T_1666d_level0_col10" class="col_heading level0 col10" >2026-09-11</th>
      <th id="T_1666d_level0_col11" class="col_heading level0 col11" >2026-09-12</th>
      <th id="T_1666d_level0_col12" class="col_heading level0 col12" >2026-09-13</th>
      <th id="T_1666d_level0_col13" class="col_heading level0 col13" >2026-09-14</th>
      <th id="T_1666d_level0_col14" class="col_heading level0 col14" >2026-09-15</th>
      <th id="T_1666d_level0_col15" class="col_heading level0 col15" >2026-09-16</th>
      <th id="T_1666d_level0_col16" class="col_heading level0 col16" >2026-09-17</th>
      <th id="T_1666d_level0_col17" class="col_heading level0 col17" >2026-09-18</th>
      <th id="T_1666d_level0_col18" class="col_heading level0 col18" >2026-09-19</th>
      <th id="T_1666d_level0_col19" class="col_heading level0 col19" >2026-09-20</th>
      <th id="T_1666d_level0_col20" class="col_heading level0 col20" >2026-09-21</th>
      <th id="T_1666d_level0_col21" class="col_heading level0 col21" >2026-09-22</th>
      <th id="T_1666d_level0_col22" class="col_heading level0 col22" >2026-09-23</th>
      <th id="T_1666d_level0_col23" class="col_heading level0 col23" >2026-09-24</th>
      <th id="T_1666d_level0_col24" class="col_heading level0 col24" >2026-09-25</th>
      <th id="T_1666d_level0_col25" class="col_heading level0 col25" >2026-09-26</th>
      <th id="T_1666d_level0_col26" class="col_heading level0 col26" >2026-09-27</th>
      <th id="T_1666d_level0_col27" class="col_heading level0 col27" >2026-09-28</th>
      <th id="T_1666d_level0_col28" class="col_heading level0 col28" >2026-09-29</th>
      <th id="T_1666d_level0_col29" class="col_heading level0 col29" >2026-09-30</th>
      <th id="T_1666d_level0_col30" class="col_heading level0 col30" >Total</th>
    </tr>
    <tr>
      <th class="index_name level0" >Line</th>
      <th class="index_name level1" >Shift</th>
      <th class="index_name level2" >Product</th>
      <th class="blank col0" >&nbsp;</th>
      <th class="blank col1" >&nbsp;</th>
      <th class="blank col2" >&nbsp;</th>
      <th class="blank col3" >&nbsp;</th>
      <th class="blank col4" >&nbsp;</th>
      <th class="blank col5" >&nbsp;</th>
      <th class="blank col6" >&nbsp;</th>
      <th class="blank col7" >&nbsp;</th>
      <th class="blank col8" >&nbsp;</th>
      <th class="blank col9" >&nbsp;</th>
      <th class="blank col10" >&nbsp;</th>
      <th class="blank col11" >&nbsp;</th>
      <th class="blank col12" >&nbsp;</th>
      <th class="blank col13" >&nbsp;</th>
      <th class="blank col14" >&nbsp;</th>
      <th class="blank col15" >&nbsp;</th>
      <th class="blank col16" >&nbsp;</th>
      <th class="blank col17" >&nbsp;</th>
      <th class="blank col18" >&nbsp;</th>
      <th class="blank col19" >&nbsp;</th>
      <th class="blank col20" >&nbsp;</th>
      <th class="blank col21" >&nbsp;</th>
      <th class="blank col22" >&nbsp;</th>
      <th class="blank col23" >&nbsp;</th>
      <th class="blank col24" >&nbsp;</th>
      <th class="blank col25" >&nbsp;</th>
      <th class="blank col26" >&nbsp;</th>
      <th class="blank col27" >&nbsp;</th>
      <th class="blank col28" >&nbsp;</th>
      <th class="blank col29" >&nbsp;</th>
      <th class="blank col30" >&nbsp;</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <th id="T_1666d_level0_row0" class="row_heading level0 row0" rowspan="9">Pri1</th>
      <th id="T_1666d_level1_row0" class="row_heading level1 row0" rowspan="9">Day</th>
      <th id="T_1666d_level2_row0" class="row_heading level2 row0" >B</th>
      <td id="T_1666d_row0_col0" class="data row0 col0" >381,068</td>
      <td id="T_1666d_row0_col1" class="data row0 col1" >0</td>
      <td id="T_1666d_row0_col2" class="data row0 col2" >0</td>
      <td id="T_1666d_row0_col3" class="data row0 col3" >339,318</td>
      <td id="T_1666d_row0_col4" class="data row0 col4" >0</td>
      <td id="T_1666d_row0_col5" class="data row0 col5" >0</td>
      <td id="T_1666d_row0_col6" class="data row0 col6" >0</td>
      <td id="T_1666d_row0_col7" class="data row0 col7" >0</td>
      <td id="T_1666d_row0_col8" class="data row0 col8" >0</td>
      <td id="T_1666d_row0_col9" class="data row0 col9" >172,509</td>
      <td id="T_1666d_row0_col10" class="data row0 col10" >42,807</td>
      <td id="T_1666d_row0_col11" class="data row0 col11" >0</td>
      <td id="T_1666d_row0_col12" class="data row0 col12" >0</td>
      <td id="T_1666d_row0_col13" class="data row0 col13" >0</td>
      <td id="T_1666d_row0_col14" class="data row0 col14" >0</td>
      <td id="T_1666d_row0_col15" class="data row0 col15" >0</td>
      <td id="T_1666d_row0_col16" class="data row0 col16" >0</td>
      <td id="T_1666d_row0_col17" class="data row0 col17" >106,445</td>
      <td id="T_1666d_row0_col18" class="data row0 col18" >0</td>
      <td id="T_1666d_row0_col19" class="data row0 col19" >0</td>
      <td id="T_1666d_row0_col20" class="data row0 col20" >306,299</td>
      <td id="T_1666d_row0_col21" class="data row0 col21" >235,200</td>
      <td id="T_1666d_row0_col22" class="data row0 col22" >0</td>
      <td id="T_1666d_row0_col23" class="data row0 col23" >0</td>
      <td id="T_1666d_row0_col24" class="data row0 col24" >0</td>
      <td id="T_1666d_row0_col25" class="data row0 col25" >0</td>
      <td id="T_1666d_row0_col26" class="data row0 col26" >0</td>
      <td id="T_1666d_row0_col27" class="data row0 col27" >0</td>
      <td id="T_1666d_row0_col28" class="data row0 col28" >0</td>
      <td id="T_1666d_row0_col29" class="data row0 col29" >0</td>
      <td id="T_1666d_row0_col30" class="data row0 col30" >1,583,646</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row1" class="row_heading level2 row1" >H</th>
      <td id="T_1666d_row1_col0" class="data row1 col0" >194,932</td>
      <td id="T_1666d_row1_col1" class="data row1 col1" >0</td>
      <td id="T_1666d_row1_col2" class="data row1 col2" >0</td>
      <td id="T_1666d_row1_col3" class="data row1 col3" >0</td>
      <td id="T_1666d_row1_col4" class="data row1 col4" >0</td>
      <td id="T_1666d_row1_col5" class="data row1 col5" >0</td>
      <td id="T_1666d_row1_col6" class="data row1 col6" >0</td>
      <td id="T_1666d_row1_col7" class="data row1 col7" >0</td>
      <td id="T_1666d_row1_col8" class="data row1 col8" >0</td>
      <td id="T_1666d_row1_col9" class="data row1 col9" >0</td>
      <td id="T_1666d_row1_col10" class="data row1 col10" >0</td>
      <td id="T_1666d_row1_col11" class="data row1 col11" >2,000</td>
      <td id="T_1666d_row1_col12" class="data row1 col12" >0</td>
      <td id="T_1666d_row1_col13" class="data row1 col13" >0</td>
      <td id="T_1666d_row1_col14" class="data row1 col14" >0</td>
      <td id="T_1666d_row1_col15" class="data row1 col15" >2,000</td>
      <td id="T_1666d_row1_col16" class="data row1 col16" >0</td>
      <td id="T_1666d_row1_col17" class="data row1 col17" >49,364</td>
      <td id="T_1666d_row1_col18" class="data row1 col18" >0</td>
      <td id="T_1666d_row1_col19" class="data row1 col19" >0</td>
      <td id="T_1666d_row1_col20" class="data row1 col20" >0</td>
      <td id="T_1666d_row1_col21" class="data row1 col21" >0</td>
      <td id="T_1666d_row1_col22" class="data row1 col22" >0</td>
      <td id="T_1666d_row1_col23" class="data row1 col23" >0</td>
      <td id="T_1666d_row1_col24" class="data row1 col24" >0</td>
      <td id="T_1666d_row1_col25" class="data row1 col25" >0</td>
      <td id="T_1666d_row1_col26" class="data row1 col26" >0</td>
      <td id="T_1666d_row1_col27" class="data row1 col27" >0</td>
      <td id="T_1666d_row1_col28" class="data row1 col28" >0</td>
      <td id="T_1666d_row1_col29" class="data row1 col29" >0</td>
      <td id="T_1666d_row1_col30" class="data row1 col30" >248,296</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row2" class="row_heading level2 row2" >D</th>
      <td id="T_1666d_row2_col0" class="data row2 col0" >0</td>
      <td id="T_1666d_row2_col1" class="data row2 col1" >0</td>
      <td id="T_1666d_row2_col2" class="data row2 col2" >0</td>
      <td id="T_1666d_row2_col3" class="data row2 col3" >0</td>
      <td id="T_1666d_row2_col4" class="data row2 col4" >0</td>
      <td id="T_1666d_row2_col5" class="data row2 col5" >0</td>
      <td id="T_1666d_row2_col6" class="data row2 col6" >3,665</td>
      <td id="T_1666d_row2_col7" class="data row2 col7" >0</td>
      <td id="T_1666d_row2_col8" class="data row2 col8" >0</td>
      <td id="T_1666d_row2_col9" class="data row2 col9" >36,556</td>
      <td id="T_1666d_row2_col10" class="data row2 col10" >0</td>
      <td id="T_1666d_row2_col11" class="data row2 col11" >0</td>
      <td id="T_1666d_row2_col12" class="data row2 col12" >0</td>
      <td id="T_1666d_row2_col13" class="data row2 col13" >0</td>
      <td id="T_1666d_row2_col14" class="data row2 col14" >0</td>
      <td id="T_1666d_row2_col15" class="data row2 col15" >0</td>
      <td id="T_1666d_row2_col16" class="data row2 col16" >0</td>
      <td id="T_1666d_row2_col17" class="data row2 col17" >0</td>
      <td id="T_1666d_row2_col18" class="data row2 col18" >0</td>
      <td id="T_1666d_row2_col19" class="data row2 col19" >0</td>
      <td id="T_1666d_row2_col20" class="data row2 col20" >96,901</td>
      <td id="T_1666d_row2_col21" class="data row2 col21" >0</td>
      <td id="T_1666d_row2_col22" class="data row2 col22" >72,507</td>
      <td id="T_1666d_row2_col23" class="data row2 col23" >0</td>
      <td id="T_1666d_row2_col24" class="data row2 col24" >0</td>
      <td id="T_1666d_row2_col25" class="data row2 col25" >0</td>
      <td id="T_1666d_row2_col26" class="data row2 col26" >0</td>
      <td id="T_1666d_row2_col27" class="data row2 col27" >0</td>
      <td id="T_1666d_row2_col28" class="data row2 col28" >0</td>
      <td id="T_1666d_row2_col29" class="data row2 col29" >0</td>
      <td id="T_1666d_row2_col30" class="data row2 col30" >209,629</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row3" class="row_heading level2 row3" >G</th>
      <td id="T_1666d_row3_col0" class="data row3 col0" >0</td>
      <td id="T_1666d_row3_col1" class="data row3 col1" >0</td>
      <td id="T_1666d_row3_col2" class="data row3 col2" >0</td>
      <td id="T_1666d_row3_col3" class="data row3 col3" >0</td>
      <td id="T_1666d_row3_col4" class="data row3 col4" >0</td>
      <td id="T_1666d_row3_col5" class="data row3 col5" >0</td>
      <td id="T_1666d_row3_col6" class="data row3 col6" >0</td>
      <td id="T_1666d_row3_col7" class="data row3 col7" >0</td>
      <td id="T_1666d_row3_col8" class="data row3 col8" >0</td>
      <td id="T_1666d_row3_col9" class="data row3 col9" >0</td>
      <td id="T_1666d_row3_col10" class="data row3 col10" >200,758</td>
      <td id="T_1666d_row3_col11" class="data row3 col11" >0</td>
      <td id="T_1666d_row3_col12" class="data row3 col12" >0</td>
      <td id="T_1666d_row3_col13" class="data row3 col13" >0</td>
      <td id="T_1666d_row3_col14" class="data row3 col14" >0</td>
      <td id="T_1666d_row3_col15" class="data row3 col15" >2,000</td>
      <td id="T_1666d_row3_col16" class="data row3 col16" >0</td>
      <td id="T_1666d_row3_col17" class="data row3 col17" >0</td>
      <td id="T_1666d_row3_col18" class="data row3 col18" >0</td>
      <td id="T_1666d_row3_col19" class="data row3 col19" >0</td>
      <td id="T_1666d_row3_col20" class="data row3 col20" >0</td>
      <td id="T_1666d_row3_col21" class="data row3 col21" >0</td>
      <td id="T_1666d_row3_col22" class="data row3 col22" >3,368</td>
      <td id="T_1666d_row3_col23" class="data row3 col23" >0</td>
      <td id="T_1666d_row3_col24" class="data row3 col24" >0</td>
      <td id="T_1666d_row3_col25" class="data row3 col25" >0</td>
      <td id="T_1666d_row3_col26" class="data row3 col26" >0</td>
      <td id="T_1666d_row3_col27" class="data row3 col27" >0</td>
      <td id="T_1666d_row3_col28" class="data row3 col28" >39,745</td>
      <td id="T_1666d_row3_col29" class="data row3 col29" >141,829</td>
      <td id="T_1666d_row3_col30" class="data row3 col30" >387,700</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row4" class="row_heading level2 row4" >A</th>
      <td id="T_1666d_row4_col0" class="data row4 col0" >0</td>
      <td id="T_1666d_row4_col1" class="data row4 col1" >0</td>
      <td id="T_1666d_row4_col2" class="data row4 col2" >113,551</td>
      <td id="T_1666d_row4_col3" class="data row4 col3" >0</td>
      <td id="T_1666d_row4_col4" class="data row4 col4" >0</td>
      <td id="T_1666d_row4_col5" class="data row4 col5" >0</td>
      <td id="T_1666d_row4_col6" class="data row4 col6" >35,264</td>
      <td id="T_1666d_row4_col7" class="data row4 col7" >0</td>
      <td id="T_1666d_row4_col8" class="data row4 col8" >403,200</td>
      <td id="T_1666d_row4_col9" class="data row4 col9" >0</td>
      <td id="T_1666d_row4_col10" class="data row4 col10" >0</td>
      <td id="T_1666d_row4_col11" class="data row4 col11" >0</td>
      <td id="T_1666d_row4_col12" class="data row4 col12" >141,596</td>
      <td id="T_1666d_row4_col13" class="data row4 col13" >0</td>
      <td id="T_1666d_row4_col14" class="data row4 col14" >0</td>
      <td id="T_1666d_row4_col15" class="data row4 col15" >0</td>
      <td id="T_1666d_row4_col16" class="data row4 col16" >0</td>
      <td id="T_1666d_row4_col17" class="data row4 col17" >178,025</td>
      <td id="T_1666d_row4_col18" class="data row4 col18" >0</td>
      <td id="T_1666d_row4_col19" class="data row4 col19" >0</td>
      <td id="T_1666d_row4_col20" class="data row4 col20" >0</td>
      <td id="T_1666d_row4_col21" class="data row4 col21" >50,070</td>
      <td id="T_1666d_row4_col22" class="data row4 col22" >0</td>
      <td id="T_1666d_row4_col23" class="data row4 col23" >0</td>
      <td id="T_1666d_row4_col24" class="data row4 col24" >0</td>
      <td id="T_1666d_row4_col25" class="data row4 col25" >0</td>
      <td id="T_1666d_row4_col26" class="data row4 col26" >0</td>
      <td id="T_1666d_row4_col27" class="data row4 col27" >0</td>
      <td id="T_1666d_row4_col28" class="data row4 col28" >0</td>
      <td id="T_1666d_row4_col29" class="data row4 col29" >0</td>
      <td id="T_1666d_row4_col30" class="data row4 col30" >921,706</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row5" class="row_heading level2 row5" >I</th>
      <td id="T_1666d_row5_col0" class="data row5 col0" >0</td>
      <td id="T_1666d_row5_col1" class="data row5 col1" >145,408</td>
      <td id="T_1666d_row5_col2" class="data row5 col2" >462,449</td>
      <td id="T_1666d_row5_col3" class="data row5 col3" >236,682</td>
      <td id="T_1666d_row5_col4" class="data row5 col4" >0</td>
      <td id="T_1666d_row5_col5" class="data row5 col5" >0</td>
      <td id="T_1666d_row5_col6" class="data row5 col6" >364,271</td>
      <td id="T_1666d_row5_col7" class="data row5 col7" >9,689</td>
      <td id="T_1666d_row5_col8" class="data row5 col8" >17,459</td>
      <td id="T_1666d_row5_col9" class="data row5 col9" >0</td>
      <td id="T_1666d_row5_col10" class="data row5 col10" >322,583</td>
      <td id="T_1666d_row5_col11" class="data row5 col11" >403,151</td>
      <td id="T_1666d_row5_col12" class="data row5 col12" >344,097</td>
      <td id="T_1666d_row5_col13" class="data row5 col13" >425,089</td>
      <td id="T_1666d_row5_col14" class="data row5 col14" >264,750</td>
      <td id="T_1666d_row5_col15" class="data row5 col15" >142,582</td>
      <td id="T_1666d_row5_col16" class="data row5 col16" >237,177</td>
      <td id="T_1666d_row5_col17" class="data row5 col17" >69,366</td>
      <td id="T_1666d_row5_col18" class="data row5 col18" >0</td>
      <td id="T_1666d_row5_col19" class="data row5 col19" >0</td>
      <td id="T_1666d_row5_col20" class="data row5 col20" >0</td>
      <td id="T_1666d_row5_col21" class="data row5 col21" >186,395</td>
      <td id="T_1666d_row5_col22" class="data row5 col22" >113,116</td>
      <td id="T_1666d_row5_col23" class="data row5 col23" >212,118</td>
      <td id="T_1666d_row5_col24" class="data row5 col24" >0</td>
      <td id="T_1666d_row5_col25" class="data row5 col25" >0</td>
      <td id="T_1666d_row5_col26" class="data row5 col26" >0</td>
      <td id="T_1666d_row5_col27" class="data row5 col27" >390,581</td>
      <td id="T_1666d_row5_col28" class="data row5 col28" >0</td>
      <td id="T_1666d_row5_col29" class="data row5 col29" >0</td>
      <td id="T_1666d_row5_col30" class="data row5 col30" >4,346,963</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row6" class="row_heading level2 row6" >E</th>
      <td id="T_1666d_row6_col0" class="data row6 col0" >0</td>
      <td id="T_1666d_row6_col1" class="data row6 col1" >8,777</td>
      <td id="T_1666d_row6_col2" class="data row6 col2" >0</td>
      <td id="T_1666d_row6_col3" class="data row6 col3" >0</td>
      <td id="T_1666d_row6_col4" class="data row6 col4" >0</td>
      <td id="T_1666d_row6_col5" class="data row6 col5" >0</td>
      <td id="T_1666d_row6_col6" class="data row6 col6" >0</td>
      <td id="T_1666d_row6_col7" class="data row6 col7" >0</td>
      <td id="T_1666d_row6_col8" class="data row6 col8" >0</td>
      <td id="T_1666d_row6_col9" class="data row6 col9" >0</td>
      <td id="T_1666d_row6_col10" class="data row6 col10" >0</td>
      <td id="T_1666d_row6_col11" class="data row6 col11" >0</td>
      <td id="T_1666d_row6_col12" class="data row6 col12" >0</td>
      <td id="T_1666d_row6_col13" class="data row6 col13" >0</td>
      <td id="T_1666d_row6_col14" class="data row6 col14" >0</td>
      <td id="T_1666d_row6_col15" class="data row6 col15" >0</td>
      <td id="T_1666d_row6_col16" class="data row6 col16" >0</td>
      <td id="T_1666d_row6_col17" class="data row6 col17" >0</td>
      <td id="T_1666d_row6_col18" class="data row6 col18" >0</td>
      <td id="T_1666d_row6_col19" class="data row6 col19" >0</td>
      <td id="T_1666d_row6_col20" class="data row6 col20" >0</td>
      <td id="T_1666d_row6_col21" class="data row6 col21" >0</td>
      <td id="T_1666d_row6_col22" class="data row6 col22" >0</td>
      <td id="T_1666d_row6_col23" class="data row6 col23" >0</td>
      <td id="T_1666d_row6_col24" class="data row6 col24" >0</td>
      <td id="T_1666d_row6_col25" class="data row6 col25" >0</td>
      <td id="T_1666d_row6_col26" class="data row6 col26" >0</td>
      <td id="T_1666d_row6_col27" class="data row6 col27" >0</td>
      <td id="T_1666d_row6_col28" class="data row6 col28" >0</td>
      <td id="T_1666d_row6_col29" class="data row6 col29" >0</td>
      <td id="T_1666d_row6_col30" class="data row6 col30" >8,777</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row7" class="row_heading level2 row7" >C</th>
      <td id="T_1666d_row7_col0" class="data row7 col0" >0</td>
      <td id="T_1666d_row7_col1" class="data row7 col1" >161,252</td>
      <td id="T_1666d_row7_col2" class="data row7 col2" >0</td>
      <td id="T_1666d_row7_col3" class="data row7 col3" >0</td>
      <td id="T_1666d_row7_col4" class="data row7 col4" >0</td>
      <td id="T_1666d_row7_col5" class="data row7 col5" >0</td>
      <td id="T_1666d_row7_col6" class="data row7 col6" >0</td>
      <td id="T_1666d_row7_col7" class="data row7 col7" >393,511</td>
      <td id="T_1666d_row7_col8" class="data row7 col8" >0</td>
      <td id="T_1666d_row7_col9" class="data row7 col9" >194,135</td>
      <td id="T_1666d_row7_col10" class="data row7 col10" >0</td>
      <td id="T_1666d_row7_col11" class="data row7 col11" >0</td>
      <td id="T_1666d_row7_col12" class="data row7 col12" >0</td>
      <td id="T_1666d_row7_col13" class="data row7 col13" >150,911</td>
      <td id="T_1666d_row7_col14" class="data row7 col14" >0</td>
      <td id="T_1666d_row7_col15" class="data row7 col15" >257,111</td>
      <td id="T_1666d_row7_col16" class="data row7 col16" >0</td>
      <td id="T_1666d_row7_col17" class="data row7 col17" >0</td>
      <td id="T_1666d_row7_col18" class="data row7 col18" >0</td>
      <td id="T_1666d_row7_col19" class="data row7 col19" >0</td>
      <td id="T_1666d_row7_col20" class="data row7 col20" >0</td>
      <td id="T_1666d_row7_col21" class="data row7 col21" >0</td>
      <td id="T_1666d_row7_col22" class="data row7 col22" >0</td>
      <td id="T_1666d_row7_col23" class="data row7 col23" >191,082</td>
      <td id="T_1666d_row7_col24" class="data row7 col24" >353,665</td>
      <td id="T_1666d_row7_col25" class="data row7 col25" >0</td>
      <td id="T_1666d_row7_col26" class="data row7 col26" >0</td>
      <td id="T_1666d_row7_col27" class="data row7 col27" >12,619</td>
      <td id="T_1666d_row7_col28" class="data row7 col28" >403,200</td>
      <td id="T_1666d_row7_col29" class="data row7 col29" >434,171</td>
      <td id="T_1666d_row7_col30" class="data row7 col30" >2,551,657</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row8" class="row_heading level2 row8" >F</th>
      <td id="T_1666d_row8_col0" class="data row8 col0" >0</td>
      <td id="T_1666d_row8_col1" class="data row8 col1" >0</td>
      <td id="T_1666d_row8_col2" class="data row8 col2" >0</td>
      <td id="T_1666d_row8_col3" class="data row8 col3" >0</td>
      <td id="T_1666d_row8_col4" class="data row8 col4" >0</td>
      <td id="T_1666d_row8_col5" class="data row8 col5" >0</td>
      <td id="T_1666d_row8_col6" class="data row8 col6" >0</td>
      <td id="T_1666d_row8_col7" class="data row8 col7" >0</td>
      <td id="T_1666d_row8_col8" class="data row8 col8" >0</td>
      <td id="T_1666d_row8_col9" class="data row8 col9" >0</td>
      <td id="T_1666d_row8_col10" class="data row8 col10" >9,852</td>
      <td id="T_1666d_row8_col11" class="data row8 col11" >0</td>
      <td id="T_1666d_row8_col12" class="data row8 col12" >0</td>
      <td id="T_1666d_row8_col13" class="data row8 col13" >0</td>
      <td id="T_1666d_row8_col14" class="data row8 col14" >311,250</td>
      <td id="T_1666d_row8_col15" class="data row8 col15" >0</td>
      <td id="T_1666d_row8_col16" class="data row8 col16" >0</td>
      <td id="T_1666d_row8_col17" class="data row8 col17" >0</td>
      <td id="T_1666d_row8_col18" class="data row8 col18" >0</td>
      <td id="T_1666d_row8_col19" class="data row8 col19" >0</td>
      <td id="T_1666d_row8_col20" class="data row8 col20" >0</td>
      <td id="T_1666d_row8_col21" class="data row8 col21" >0</td>
      <td id="T_1666d_row8_col22" class="data row8 col22" >214,209</td>
      <td id="T_1666d_row8_col23" class="data row8 col23" >0</td>
      <td id="T_1666d_row8_col24" class="data row8 col24" >120,045</td>
      <td id="T_1666d_row8_col25" class="data row8 col25" >0</td>
      <td id="T_1666d_row8_col26" class="data row8 col26" >0</td>
      <td id="T_1666d_row8_col27" class="data row8 col27" >0</td>
      <td id="T_1666d_row8_col28" class="data row8 col28" >0</td>
      <td id="T_1666d_row8_col29" class="data row8 col29" >0</td>
      <td id="T_1666d_row8_col30" class="data row8 col30" >655,356</td>
    </tr>
    <tr>
      <th id="T_1666d_level0_row9" class="row_heading level0 row9" rowspan="8">Pri2</th>
      <th id="T_1666d_level1_row9" class="row_heading level1 row9" rowspan="8">Day</th>
      <th id="T_1666d_level2_row9" class="row_heading level2 row9" >B</th>
      <td id="T_1666d_row9_col0" class="data row9 col0" >141,668</td>
      <td id="T_1666d_row9_col1" class="data row9 col1" >0</td>
      <td id="T_1666d_row9_col2" class="data row9 col2" >403,200</td>
      <td id="T_1666d_row9_col3" class="data row9 col3" >61,361</td>
      <td id="T_1666d_row9_col4" class="data row9 col4" >0</td>
      <td id="T_1666d_row9_col5" class="data row9 col5" >0</td>
      <td id="T_1666d_row9_col6" class="data row9 col6" >0</td>
      <td id="T_1666d_row9_col7" class="data row9 col7" >134,284</td>
      <td id="T_1666d_row9_col8" class="data row9 col8" >0</td>
      <td id="T_1666d_row9_col9" class="data row9 col9" >0</td>
      <td id="T_1666d_row9_col10" class="data row9 col10" >0</td>
      <td id="T_1666d_row9_col11" class="data row9 col11" >0</td>
      <td id="T_1666d_row9_col12" class="data row9 col12" >0</td>
      <td id="T_1666d_row9_col13" class="data row9 col13" >363,621</td>
      <td id="T_1666d_row9_col14" class="data row9 col14" >310,903</td>
      <td id="T_1666d_row9_col15" class="data row9 col15" >0</td>
      <td id="T_1666d_row9_col16" class="data row9 col16" >0</td>
      <td id="T_1666d_row9_col17" class="data row9 col17" >23,822</td>
      <td id="T_1666d_row9_col18" class="data row9 col18" >0</td>
      <td id="T_1666d_row9_col19" class="data row9 col19" >0</td>
      <td id="T_1666d_row9_col20" class="data row9 col20" >81,177</td>
      <td id="T_1666d_row9_col21" class="data row9 col21" >0</td>
      <td id="T_1666d_row9_col22" class="data row9 col22" >0</td>
      <td id="T_1666d_row9_col23" class="data row9 col23" >0</td>
      <td id="T_1666d_row9_col24" class="data row9 col24" >99,509</td>
      <td id="T_1666d_row9_col25" class="data row9 col25" >0</td>
      <td id="T_1666d_row9_col26" class="data row9 col26" >0</td>
      <td id="T_1666d_row9_col27" class="data row9 col27" >0</td>
      <td id="T_1666d_row9_col28" class="data row9 col28" >0</td>
      <td id="T_1666d_row9_col29" class="data row9 col29" >168,749</td>
      <td id="T_1666d_row9_col30" class="data row9 col30" >1,788,294</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row10" class="row_heading level2 row10" >H</th>
      <td id="T_1666d_row10_col0" class="data row10 col0" >0</td>
      <td id="T_1666d_row10_col1" class="data row10 col1" >0</td>
      <td id="T_1666d_row10_col2" class="data row10 col2" >0</td>
      <td id="T_1666d_row10_col3" class="data row10 col3" >0</td>
      <td id="T_1666d_row10_col4" class="data row10 col4" >0</td>
      <td id="T_1666d_row10_col5" class="data row10 col5" >0</td>
      <td id="T_1666d_row10_col6" class="data row10 col6" >0</td>
      <td id="T_1666d_row10_col7" class="data row10 col7" >0</td>
      <td id="T_1666d_row10_col8" class="data row10 col8" >0</td>
      <td id="T_1666d_row10_col9" class="data row10 col9" >0</td>
      <td id="T_1666d_row10_col10" class="data row10 col10" >0</td>
      <td id="T_1666d_row10_col11" class="data row10 col11" >0</td>
      <td id="T_1666d_row10_col12" class="data row10 col12" >0</td>
      <td id="T_1666d_row10_col13" class="data row10 col13" >0</td>
      <td id="T_1666d_row10_col14" class="data row10 col14" >0</td>
      <td id="T_1666d_row10_col15" class="data row10 col15" >403,200</td>
      <td id="T_1666d_row10_col16" class="data row10 col16" >0</td>
      <td id="T_1666d_row10_col17" class="data row10 col17" >0</td>
      <td id="T_1666d_row10_col18" class="data row10 col18" >0</td>
      <td id="T_1666d_row10_col19" class="data row10 col19" >0</td>
      <td id="T_1666d_row10_col20" class="data row10 col20" >0</td>
      <td id="T_1666d_row10_col21" class="data row10 col21" >0</td>
      <td id="T_1666d_row10_col22" class="data row10 col22" >0</td>
      <td id="T_1666d_row10_col23" class="data row10 col23" >0</td>
      <td id="T_1666d_row10_col24" class="data row10 col24" >0</td>
      <td id="T_1666d_row10_col25" class="data row10 col25" >0</td>
      <td id="T_1666d_row10_col26" class="data row10 col26" >0</td>
      <td id="T_1666d_row10_col27" class="data row10 col27" >0</td>
      <td id="T_1666d_row10_col28" class="data row10 col28" >0</td>
      <td id="T_1666d_row10_col29" class="data row10 col29" >0</td>
      <td id="T_1666d_row10_col30" class="data row10 col30" >403,200</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row11" class="row_heading level2 row11" >D</th>
      <td id="T_1666d_row11_col0" class="data row11 col0" >54,475</td>
      <td id="T_1666d_row11_col1" class="data row11 col1" >0</td>
      <td id="T_1666d_row11_col2" class="data row11 col2" >0</td>
      <td id="T_1666d_row11_col3" class="data row11 col3" >4,566</td>
      <td id="T_1666d_row11_col4" class="data row11 col4" >0</td>
      <td id="T_1666d_row11_col5" class="data row11 col5" >0</td>
      <td id="T_1666d_row11_col6" class="data row11 col6" >0</td>
      <td id="T_1666d_row11_col7" class="data row11 col7" >13,819</td>
      <td id="T_1666d_row11_col8" class="data row11 col8" >0</td>
      <td id="T_1666d_row11_col9" class="data row11 col9" >0</td>
      <td id="T_1666d_row11_col10" class="data row11 col10" >0</td>
      <td id="T_1666d_row11_col11" class="data row11 col11" >2,857</td>
      <td id="T_1666d_row11_col12" class="data row11 col12" >0</td>
      <td id="T_1666d_row11_col13" class="data row11 col13" >10,705</td>
      <td id="T_1666d_row11_col14" class="data row11 col14" >0</td>
      <td id="T_1666d_row11_col15" class="data row11 col15" >0</td>
      <td id="T_1666d_row11_col16" class="data row11 col16" >0</td>
      <td id="T_1666d_row11_col17" class="data row11 col17" >0</td>
      <td id="T_1666d_row11_col18" class="data row11 col18" >0</td>
      <td id="T_1666d_row11_col19" class="data row11 col19" >0</td>
      <td id="T_1666d_row11_col20" class="data row11 col20" >0</td>
      <td id="T_1666d_row11_col21" class="data row11 col21" >0</td>
      <td id="T_1666d_row11_col22" class="data row11 col22" >0</td>
      <td id="T_1666d_row11_col23" class="data row11 col23" >0</td>
      <td id="T_1666d_row11_col24" class="data row11 col24" >0</td>
      <td id="T_1666d_row11_col25" class="data row11 col25" >0</td>
      <td id="T_1666d_row11_col26" class="data row11 col26" >0</td>
      <td id="T_1666d_row11_col27" class="data row11 col27" >0</td>
      <td id="T_1666d_row11_col28" class="data row11 col28" >0</td>
      <td id="T_1666d_row11_col29" class="data row11 col29" >0</td>
      <td id="T_1666d_row11_col30" class="data row11 col30" >86,422</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row12" class="row_heading level2 row12" >G</th>
      <td id="T_1666d_row12_col0" class="data row12 col0" >72,078</td>
      <td id="T_1666d_row12_col1" class="data row12 col1" >0</td>
      <td id="T_1666d_row12_col2" class="data row12 col2" >0</td>
      <td id="T_1666d_row12_col3" class="data row12 col3" >0</td>
      <td id="T_1666d_row12_col4" class="data row12 col4" >0</td>
      <td id="T_1666d_row12_col5" class="data row12 col5" >0</td>
      <td id="T_1666d_row12_col6" class="data row12 col6" >0</td>
      <td id="T_1666d_row12_col7" class="data row12 col7" >0</td>
      <td id="T_1666d_row12_col8" class="data row12 col8" >0</td>
      <td id="T_1666d_row12_col9" class="data row12 col9" >0</td>
      <td id="T_1666d_row12_col10" class="data row12 col10" >0</td>
      <td id="T_1666d_row12_col11" class="data row12 col11" >0</td>
      <td id="T_1666d_row12_col12" class="data row12 col12" >0</td>
      <td id="T_1666d_row12_col13" class="data row12 col13" >0</td>
      <td id="T_1666d_row12_col14" class="data row12 col14" >0</td>
      <td id="T_1666d_row12_col15" class="data row12 col15" >0</td>
      <td id="T_1666d_row12_col16" class="data row12 col16" >0</td>
      <td id="T_1666d_row12_col17" class="data row12 col17" >0</td>
      <td id="T_1666d_row12_col18" class="data row12 col18" >0</td>
      <td id="T_1666d_row12_col19" class="data row12 col19" >0</td>
      <td id="T_1666d_row12_col20" class="data row12 col20" >0</td>
      <td id="T_1666d_row12_col21" class="data row12 col21" >0</td>
      <td id="T_1666d_row12_col22" class="data row12 col22" >0</td>
      <td id="T_1666d_row12_col23" class="data row12 col23" >0</td>
      <td id="T_1666d_row12_col24" class="data row12 col24" >0</td>
      <td id="T_1666d_row12_col25" class="data row12 col25" >0</td>
      <td id="T_1666d_row12_col26" class="data row12 col26" >0</td>
      <td id="T_1666d_row12_col27" class="data row12 col27" >0</td>
      <td id="T_1666d_row12_col28" class="data row12 col28" >0</td>
      <td id="T_1666d_row12_col29" class="data row12 col29" >19,184</td>
      <td id="T_1666d_row12_col30" class="data row12 col30" >91,262</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row13" class="row_heading level2 row13" >A</th>
      <td id="T_1666d_row13_col0" class="data row13 col0" >0</td>
      <td id="T_1666d_row13_col1" class="data row13 col1" >0</td>
      <td id="T_1666d_row13_col2" class="data row13 col2" >0</td>
      <td id="T_1666d_row13_col3" class="data row13 col3" >0</td>
      <td id="T_1666d_row13_col4" class="data row13 col4" >0</td>
      <td id="T_1666d_row13_col5" class="data row13 col5" >0</td>
      <td id="T_1666d_row13_col6" class="data row13 col6" >0</td>
      <td id="T_1666d_row13_col7" class="data row13 col7" >0</td>
      <td id="T_1666d_row13_col8" class="data row13 col8" >0</td>
      <td id="T_1666d_row13_col9" class="data row13 col9" >0</td>
      <td id="T_1666d_row13_col10" class="data row13 col10" >172,800</td>
      <td id="T_1666d_row13_col11" class="data row13 col11" >0</td>
      <td id="T_1666d_row13_col12" class="data row13 col12" >0</td>
      <td id="T_1666d_row13_col13" class="data row13 col13" >0</td>
      <td id="T_1666d_row13_col14" class="data row13 col14" >0</td>
      <td id="T_1666d_row13_col15" class="data row13 col15" >0</td>
      <td id="T_1666d_row13_col16" class="data row13 col16" >0</td>
      <td id="T_1666d_row13_col17" class="data row13 col17" >0</td>
      <td id="T_1666d_row13_col18" class="data row13 col18" >0</td>
      <td id="T_1666d_row13_col19" class="data row13 col19" >0</td>
      <td id="T_1666d_row13_col20" class="data row13 col20" >0</td>
      <td id="T_1666d_row13_col21" class="data row13 col21" >0</td>
      <td id="T_1666d_row13_col22" class="data row13 col22" >102,722</td>
      <td id="T_1666d_row13_col23" class="data row13 col23" >63,876</td>
      <td id="T_1666d_row13_col24" class="data row13 col24" >0</td>
      <td id="T_1666d_row13_col25" class="data row13 col25" >0</td>
      <td id="T_1666d_row13_col26" class="data row13 col26" >0</td>
      <td id="T_1666d_row13_col27" class="data row13 col27" >0</td>
      <td id="T_1666d_row13_col28" class="data row13 col28" >0</td>
      <td id="T_1666d_row13_col29" class="data row13 col29" >0</td>
      <td id="T_1666d_row13_col30" class="data row13 col30" >339,398</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row14" class="row_heading level2 row14" >I</th>
      <td id="T_1666d_row14_col0" class="data row14 col0" >0</td>
      <td id="T_1666d_row14_col1" class="data row14 col1" >230,899</td>
      <td id="T_1666d_row14_col2" class="data row14 col2" >0</td>
      <td id="T_1666d_row14_col3" class="data row14 col3" >441,439</td>
      <td id="T_1666d_row14_col4" class="data row14 col4" >0</td>
      <td id="T_1666d_row14_col5" class="data row14 col5" >0</td>
      <td id="T_1666d_row14_col6" class="data row14 col6" >278,225</td>
      <td id="T_1666d_row14_col7" class="data row14 col7" >255,097</td>
      <td id="T_1666d_row14_col8" class="data row14 col8" >339,366</td>
      <td id="T_1666d_row14_col9" class="data row14 col9" >403,200</td>
      <td id="T_1666d_row14_col10" class="data row14 col10" >338,159</td>
      <td id="T_1666d_row14_col11" class="data row14 col11" >283,027</td>
      <td id="T_1666d_row14_col12" class="data row14 col12" >505,582</td>
      <td id="T_1666d_row14_col13" class="data row14 col13" >0</td>
      <td id="T_1666d_row14_col14" class="data row14 col14" >0</td>
      <td id="T_1666d_row14_col15" class="data row14 col15" >0</td>
      <td id="T_1666d_row14_col16" class="data row14 col16" >376,917</td>
      <td id="T_1666d_row14_col17" class="data row14 col17" >363,730</td>
      <td id="T_1666d_row14_col18" class="data row14 col18" >0</td>
      <td id="T_1666d_row14_col19" class="data row14 col19" >0</td>
      <td id="T_1666d_row14_col20" class="data row14 col20" >322,023</td>
      <td id="T_1666d_row14_col21" class="data row14 col21" >2,000</td>
      <td id="T_1666d_row14_col22" class="data row14 col22" >0</td>
      <td id="T_1666d_row14_col23" class="data row14 col23" >339,324</td>
      <td id="T_1666d_row14_col24" class="data row14 col24" >0</td>
      <td id="T_1666d_row14_col25" class="data row14 col25" >0</td>
      <td id="T_1666d_row14_col26" class="data row14 col26" >0</td>
      <td id="T_1666d_row14_col27" class="data row14 col27" >0</td>
      <td id="T_1666d_row14_col28" class="data row14 col28" >133,563</td>
      <td id="T_1666d_row14_col29" class="data row14 col29" >0</td>
      <td id="T_1666d_row14_col30" class="data row14 col30" >4,612,551</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row15" class="row_heading level2 row15" >C</th>
      <td id="T_1666d_row15_col0" class="data row15 col0" >0</td>
      <td id="T_1666d_row15_col1" class="data row15 col1" >0</td>
      <td id="T_1666d_row15_col2" class="data row15 col2" >0</td>
      <td id="T_1666d_row15_col3" class="data row15 col3" >9,775</td>
      <td id="T_1666d_row15_col4" class="data row15 col4" >0</td>
      <td id="T_1666d_row15_col5" class="data row15 col5" >0</td>
      <td id="T_1666d_row15_col6" class="data row15 col6" >242,114</td>
      <td id="T_1666d_row15_col7" class="data row15 col7" >0</td>
      <td id="T_1666d_row15_col8" class="data row15 col8" >63,834</td>
      <td id="T_1666d_row15_col9" class="data row15 col9" >0</td>
      <td id="T_1666d_row15_col10" class="data row15 col10" >65,041</td>
      <td id="T_1666d_row15_col11" class="data row15 col11" >204,869</td>
      <td id="T_1666d_row15_col12" class="data row15 col12" >0</td>
      <td id="T_1666d_row15_col13" class="data row15 col13" >0</td>
      <td id="T_1666d_row15_col14" class="data row15 col14" >129,481</td>
      <td id="T_1666d_row15_col15" class="data row15 col15" >0</td>
      <td id="T_1666d_row15_col16" class="data row15 col16" >26,283</td>
      <td id="T_1666d_row15_col17" class="data row15 col17" >0</td>
      <td id="T_1666d_row15_col18" class="data row15 col18" >0</td>
      <td id="T_1666d_row15_col19" class="data row15 col19" >0</td>
      <td id="T_1666d_row15_col20" class="data row15 col20" >0</td>
      <td id="T_1666d_row15_col21" class="data row15 col21" >402,883</td>
      <td id="T_1666d_row15_col22" class="data row15 col22" >300,478</td>
      <td id="T_1666d_row15_col23" class="data row15 col23" >0</td>
      <td id="T_1666d_row15_col24" class="data row15 col24" >0</td>
      <td id="T_1666d_row15_col25" class="data row15 col25" >0</td>
      <td id="T_1666d_row15_col26" class="data row15 col26" >0</td>
      <td id="T_1666d_row15_col27" class="data row15 col27" >403,200</td>
      <td id="T_1666d_row15_col28" class="data row15 col28" >0</td>
      <td id="T_1666d_row15_col29" class="data row15 col29" >218,629</td>
      <td id="T_1666d_row15_col30" class="data row15 col30" >2,066,587</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row16" class="row_heading level2 row16" >F</th>
      <td id="T_1666d_row16_col0" class="data row16 col0" >0</td>
      <td id="T_1666d_row16_col1" class="data row16 col1" >172,300</td>
      <td id="T_1666d_row16_col2" class="data row16 col2" >0</td>
      <td id="T_1666d_row16_col3" class="data row16 col3" >0</td>
      <td id="T_1666d_row16_col4" class="data row16 col4" >0</td>
      <td id="T_1666d_row16_col5" class="data row16 col5" >0</td>
      <td id="T_1666d_row16_col6" class="data row16 col6" >0</td>
      <td id="T_1666d_row16_col7" class="data row16 col7" >0</td>
      <td id="T_1666d_row16_col8" class="data row16 col8" >0</td>
      <td id="T_1666d_row16_col9" class="data row16 col9" >0</td>
      <td id="T_1666d_row16_col10" class="data row16 col10" >0</td>
      <td id="T_1666d_row16_col11" class="data row16 col11" >0</td>
      <td id="T_1666d_row16_col12" class="data row16 col12" >0</td>
      <td id="T_1666d_row16_col13" class="data row16 col13" >28,874</td>
      <td id="T_1666d_row16_col14" class="data row16 col14" >0</td>
      <td id="T_1666d_row16_col15" class="data row16 col15" >0</td>
      <td id="T_1666d_row16_col16" class="data row16 col16" >0</td>
      <td id="T_1666d_row16_col17" class="data row16 col17" >15,648</td>
      <td id="T_1666d_row16_col18" class="data row16 col18" >0</td>
      <td id="T_1666d_row16_col19" class="data row16 col19" >0</td>
      <td id="T_1666d_row16_col20" class="data row16 col20" >0</td>
      <td id="T_1666d_row16_col21" class="data row16 col21" >0</td>
      <td id="T_1666d_row16_col22" class="data row16 col22" >0</td>
      <td id="T_1666d_row16_col23" class="data row16 col23" >0</td>
      <td id="T_1666d_row16_col24" class="data row16 col24" >303,691</td>
      <td id="T_1666d_row16_col25" class="data row16 col25" >0</td>
      <td id="T_1666d_row16_col26" class="data row16 col26" >0</td>
      <td id="T_1666d_row16_col27" class="data row16 col27" >0</td>
      <td id="T_1666d_row16_col28" class="data row16 col28" >269,637</td>
      <td id="T_1666d_row16_col29" class="data row16 col29" >0</td>
      <td id="T_1666d_row16_col30" class="data row16 col30" >790,150</td>
    </tr>
    <tr>
      <th id="T_1666d_level0_row17" class="row_heading level0 row17" rowspan="9">Cut1</th>
      <th id="T_1666d_level1_row17" class="row_heading level1 row17" rowspan="9">Day</th>
      <th id="T_1666d_level2_row17" class="row_heading level2 row17" >B</th>
      <td id="T_1666d_row17_col0" class="data row17 col0" >522,736</td>
      <td id="T_1666d_row17_col1" class="data row17 col1" >0</td>
      <td id="T_1666d_row17_col2" class="data row17 col2" >88,525</td>
      <td id="T_1666d_row17_col3" class="data row17 col3" >339,318</td>
      <td id="T_1666d_row17_col4" class="data row17 col4" >0</td>
      <td id="T_1666d_row17_col5" class="data row17 col5" >0</td>
      <td id="T_1666d_row17_col6" class="data row17 col6" >0</td>
      <td id="T_1666d_row17_col7" class="data row17 col7" >50,790</td>
      <td id="T_1666d_row17_col8" class="data row17 col8" >459,530</td>
      <td id="T_1666d_row17_col9" class="data row17 col9" >172,509</td>
      <td id="T_1666d_row17_col10" class="data row17 col10" >42,807</td>
      <td id="T_1666d_row17_col11" class="data row17 col11" >0</td>
      <td id="T_1666d_row17_col12" class="data row17 col12" >0</td>
      <td id="T_1666d_row17_col13" class="data row17 col13" >31,123</td>
      <td id="T_1666d_row17_col14" class="data row17 col14" >310,903</td>
      <td id="T_1666d_row17_col15" class="data row17 col15" >165,518</td>
      <td id="T_1666d_row17_col16" class="data row17 col16" >67,482</td>
      <td id="T_1666d_row17_col17" class="data row17 col17" >229,765</td>
      <td id="T_1666d_row17_col18" class="data row17 col18" >0</td>
      <td id="T_1666d_row17_col19" class="data row17 col19" >0</td>
      <td id="T_1666d_row17_col20" class="data row17 col20" >387,476</td>
      <td id="T_1666d_row17_col21" class="data row17 col21" >235,200</td>
      <td id="T_1666d_row17_col22" class="data row17 col22" >0</td>
      <td id="T_1666d_row17_col23" class="data row17 col23" >0</td>
      <td id="T_1666d_row17_col24" class="data row17 col24" >0</td>
      <td id="T_1666d_row17_col25" class="data row17 col25" >0</td>
      <td id="T_1666d_row17_col26" class="data row17 col26" >0</td>
      <td id="T_1666d_row17_col27" class="data row17 col27" >99,509</td>
      <td id="T_1666d_row17_col28" class="data row17 col28" >0</td>
      <td id="T_1666d_row17_col29" class="data row17 col29" >0</td>
      <td id="T_1666d_row17_col30" class="data row17 col30" >3,203,191</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row18" class="row_heading level2 row18" >H</th>
      <td id="T_1666d_row18_col0" class="data row18 col0" >0</td>
      <td id="T_1666d_row18_col1" class="data row18 col1" >0</td>
      <td id="T_1666d_row18_col2" class="data row18 col2" >0</td>
      <td id="T_1666d_row18_col3" class="data row18 col3" >0</td>
      <td id="T_1666d_row18_col4" class="data row18 col4" >0</td>
      <td id="T_1666d_row18_col5" class="data row18 col5" >0</td>
      <td id="T_1666d_row18_col6" class="data row18 col6" >0</td>
      <td id="T_1666d_row18_col7" class="data row18 col7" >0</td>
      <td id="T_1666d_row18_col8" class="data row18 col8" >0</td>
      <td id="T_1666d_row18_col9" class="data row18 col9" >0</td>
      <td id="T_1666d_row18_col10" class="data row18 col10" >0</td>
      <td id="T_1666d_row18_col11" class="data row18 col11" >0</td>
      <td id="T_1666d_row18_col12" class="data row18 col12" >0</td>
      <td id="T_1666d_row18_col13" class="data row18 col13" >22,132</td>
      <td id="T_1666d_row18_col14" class="data row18 col14" >0</td>
      <td id="T_1666d_row18_col15" class="data row18 col15" >0</td>
      <td id="T_1666d_row18_col16" class="data row18 col16" >0</td>
      <td id="T_1666d_row18_col17" class="data row18 col17" >0</td>
      <td id="T_1666d_row18_col18" class="data row18 col18" >0</td>
      <td id="T_1666d_row18_col19" class="data row18 col19" >0</td>
      <td id="T_1666d_row18_col20" class="data row18 col20" >0</td>
      <td id="T_1666d_row18_col21" class="data row18 col21" >0</td>
      <td id="T_1666d_row18_col22" class="data row18 col22" >0</td>
      <td id="T_1666d_row18_col23" class="data row18 col23" >0</td>
      <td id="T_1666d_row18_col24" class="data row18 col24" >0</td>
      <td id="T_1666d_row18_col25" class="data row18 col25" >0</td>
      <td id="T_1666d_row18_col26" class="data row18 col26" >0</td>
      <td id="T_1666d_row18_col27" class="data row18 col27" >0</td>
      <td id="T_1666d_row18_col28" class="data row18 col28" >0</td>
      <td id="T_1666d_row18_col29" class="data row18 col29" >0</td>
      <td id="T_1666d_row18_col30" class="data row18 col30" >22,132</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row19" class="row_heading level2 row19" >D</th>
      <td id="T_1666d_row19_col0" class="data row19 col0" >54,475</td>
      <td id="T_1666d_row19_col1" class="data row19 col1" >0</td>
      <td id="T_1666d_row19_col2" class="data row19 col2" >0</td>
      <td id="T_1666d_row19_col3" class="data row19 col3" >4,566</td>
      <td id="T_1666d_row19_col4" class="data row19 col4" >0</td>
      <td id="T_1666d_row19_col5" class="data row19 col5" >0</td>
      <td id="T_1666d_row19_col6" class="data row19 col6" >3,665</td>
      <td id="T_1666d_row19_col7" class="data row19 col7" >13,819</td>
      <td id="T_1666d_row19_col8" class="data row19 col8" >0</td>
      <td id="T_1666d_row19_col9" class="data row19 col9" >0</td>
      <td id="T_1666d_row19_col10" class="data row19 col10" >11,248</td>
      <td id="T_1666d_row19_col11" class="data row19 col11" >28,165</td>
      <td id="T_1666d_row19_col12" class="data row19 col12" >0</td>
      <td id="T_1666d_row19_col13" class="data row19 col13" >0</td>
      <td id="T_1666d_row19_col14" class="data row19 col14" >3,166</td>
      <td id="T_1666d_row19_col15" class="data row19 col15" >6,062</td>
      <td id="T_1666d_row19_col16" class="data row19 col16" >0</td>
      <td id="T_1666d_row19_col17" class="data row19 col17" >0</td>
      <td id="T_1666d_row19_col18" class="data row19 col18" >0</td>
      <td id="T_1666d_row19_col19" class="data row19 col19" >0</td>
      <td id="T_1666d_row19_col20" class="data row19 col20" >0</td>
      <td id="T_1666d_row19_col21" class="data row19 col21" >96,901</td>
      <td id="T_1666d_row19_col22" class="data row19 col22" >0</td>
      <td id="T_1666d_row19_col23" class="data row19 col23" >0</td>
      <td id="T_1666d_row19_col24" class="data row19 col24" >0</td>
      <td id="T_1666d_row19_col25" class="data row19 col25" >0</td>
      <td id="T_1666d_row19_col26" class="data row19 col26" >0</td>
      <td id="T_1666d_row19_col27" class="data row19 col27" >0</td>
      <td id="T_1666d_row19_col28" class="data row19 col28" >73,984</td>
      <td id="T_1666d_row19_col29" class="data row19 col29" >0</td>
      <td id="T_1666d_row19_col30" class="data row19 col30" >296,051</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row20" class="row_heading level2 row20" >G</th>
      <td id="T_1666d_row20_col0" class="data row20 col0" >72,078</td>
      <td id="T_1666d_row20_col1" class="data row20 col1" >0</td>
      <td id="T_1666d_row20_col2" class="data row20 col2" >0</td>
      <td id="T_1666d_row20_col3" class="data row20 col3" >0</td>
      <td id="T_1666d_row20_col4" class="data row20 col4" >0</td>
      <td id="T_1666d_row20_col5" class="data row20 col5" >0</td>
      <td id="T_1666d_row20_col6" class="data row20 col6" >0</td>
      <td id="T_1666d_row20_col7" class="data row20 col7" >0</td>
      <td id="T_1666d_row20_col8" class="data row20 col8" >0</td>
      <td id="T_1666d_row20_col9" class="data row20 col9" >0</td>
      <td id="T_1666d_row20_col10" class="data row20 col10" >200,758</td>
      <td id="T_1666d_row20_col11" class="data row20 col11" >0</td>
      <td id="T_1666d_row20_col12" class="data row20 col12" >0</td>
      <td id="T_1666d_row20_col13" class="data row20 col13" >0</td>
      <td id="T_1666d_row20_col14" class="data row20 col14" >0</td>
      <td id="T_1666d_row20_col15" class="data row20 col15" >0</td>
      <td id="T_1666d_row20_col16" class="data row20 col16" >0</td>
      <td id="T_1666d_row20_col17" class="data row20 col17" >2,000</td>
      <td id="T_1666d_row20_col18" class="data row20 col18" >0</td>
      <td id="T_1666d_row20_col19" class="data row20 col19" >0</td>
      <td id="T_1666d_row20_col20" class="data row20 col20" >0</td>
      <td id="T_1666d_row20_col21" class="data row20 col21" >0</td>
      <td id="T_1666d_row20_col22" class="data row20 col22" >3,368</td>
      <td id="T_1666d_row20_col23" class="data row20 col23" >0</td>
      <td id="T_1666d_row20_col24" class="data row20 col24" >0</td>
      <td id="T_1666d_row20_col25" class="data row20 col25" >0</td>
      <td id="T_1666d_row20_col26" class="data row20 col26" >0</td>
      <td id="T_1666d_row20_col27" class="data row20 col27" >0</td>
      <td id="T_1666d_row20_col28" class="data row20 col28" >0</td>
      <td id="T_1666d_row20_col29" class="data row20 col29" >200,758</td>
      <td id="T_1666d_row20_col30" class="data row20 col30" >478,962</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row21" class="row_heading level2 row21" >A</th>
      <td id="T_1666d_row21_col0" class="data row21 col0" >0</td>
      <td id="T_1666d_row21_col1" class="data row21 col1" >0</td>
      <td id="T_1666d_row21_col2" class="data row21 col2" >153,077</td>
      <td id="T_1666d_row21_col3" class="data row21 col3" >101,203</td>
      <td id="T_1666d_row21_col4" class="data row21 col4" >0</td>
      <td id="T_1666d_row21_col5" class="data row21 col5" >0</td>
      <td id="T_1666d_row21_col6" class="data row21 col6" >0</td>
      <td id="T_1666d_row21_col7" class="data row21 col7" >52,547</td>
      <td id="T_1666d_row21_col8" class="data row21 col8" >0</td>
      <td id="T_1666d_row21_col9" class="data row21 col9" >0</td>
      <td id="T_1666d_row21_col10" class="data row21 col10" >0</td>
      <td id="T_1666d_row21_col11" class="data row21 col11" >584,011</td>
      <td id="T_1666d_row21_col12" class="data row21 col12" >0</td>
      <td id="T_1666d_row21_col13" class="data row21 col13" >0</td>
      <td id="T_1666d_row21_col14" class="data row21 col14" >141,596</td>
      <td id="T_1666d_row21_col15" class="data row21 col15" >0</td>
      <td id="T_1666d_row21_col16" class="data row21 col16" >0</td>
      <td id="T_1666d_row21_col17" class="data row21 col17" >176,025</td>
      <td id="T_1666d_row21_col18" class="data row21 col18" >0</td>
      <td id="T_1666d_row21_col19" class="data row21 col19" >0</td>
      <td id="T_1666d_row21_col20" class="data row21 col20" >2,000</td>
      <td id="T_1666d_row21_col21" class="data row21 col21" >50,070</td>
      <td id="T_1666d_row21_col22" class="data row21 col22" >0</td>
      <td id="T_1666d_row21_col23" class="data row21 col23" >166,598</td>
      <td id="T_1666d_row21_col24" class="data row21 col24" >0</td>
      <td id="T_1666d_row21_col25" class="data row21 col25" >0</td>
      <td id="T_1666d_row21_col26" class="data row21 col26" >0</td>
      <td id="T_1666d_row21_col27" class="data row21 col27" >0</td>
      <td id="T_1666d_row21_col28" class="data row21 col28" >0</td>
      <td id="T_1666d_row21_col29" class="data row21 col29" >0</td>
      <td id="T_1666d_row21_col30" class="data row21 col30" >1,427,127</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row22" class="row_heading level2 row22" >I</th>
      <td id="T_1666d_row22_col0" class="data row22 col0" >0</td>
      <td id="T_1666d_row22_col1" class="data row22 col1" >376,307</td>
      <td id="T_1666d_row22_col2" class="data row22 col2" >462,449</td>
      <td id="T_1666d_row22_col3" class="data row22 col3" >669,600</td>
      <td id="T_1666d_row22_col4" class="data row22 col4" >0</td>
      <td id="T_1666d_row22_col5" class="data row22 col5" >0</td>
      <td id="T_1666d_row22_col6" class="data row22 col6" >651,017</td>
      <td id="T_1666d_row22_col7" class="data row22 col7" >264,786</td>
      <td id="T_1666d_row22_col8" class="data row22 col8" >339,366</td>
      <td id="T_1666d_row22_col9" class="data row22 col9" >390,740</td>
      <td id="T_1666d_row22_col10" class="data row22 col10" >690,661</td>
      <td id="T_1666d_row22_col11" class="data row22 col11" >617,760</td>
      <td id="T_1666d_row22_col12" class="data row22 col12" >659,322</td>
      <td id="T_1666d_row22_col13" class="data row22 col13" >525,760</td>
      <td id="T_1666d_row22_col14" class="data row22 col14" >422,854</td>
      <td id="T_1666d_row22_col15" class="data row22 col15" >142,580</td>
      <td id="T_1666d_row22_col16" class="data row22 col16" >614,096</td>
      <td id="T_1666d_row22_col17" class="data row22 col17" >236,882</td>
      <td id="T_1666d_row22_col18" class="data row22 col18" >0</td>
      <td id="T_1666d_row22_col19" class="data row22 col19" >0</td>
      <td id="T_1666d_row22_col20" class="data row22 col20" >518,237</td>
      <td id="T_1666d_row22_col21" class="data row22 col21" >133,109</td>
      <td id="T_1666d_row22_col22" class="data row22 col22" >168,402</td>
      <td id="T_1666d_row22_col23" class="data row22 col23" >551,442</td>
      <td id="T_1666d_row22_col24" class="data row22 col24" >0</td>
      <td id="T_1666d_row22_col25" class="data row22 col25" >0</td>
      <td id="T_1666d_row22_col26" class="data row22 col26" >0</td>
      <td id="T_1666d_row22_col27" class="data row22 col27" >275,718</td>
      <td id="T_1666d_row22_col28" class="data row22 col28" >174,443</td>
      <td id="T_1666d_row22_col29" class="data row22 col29" >0</td>
      <td id="T_1666d_row22_col30" class="data row22 col30" >8,885,531</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row23" class="row_heading level2 row23" >E</th>
      <td id="T_1666d_row23_col0" class="data row23 col0" >0</td>
      <td id="T_1666d_row23_col1" class="data row23 col1" >0</td>
      <td id="T_1666d_row23_col2" class="data row23 col2" >0</td>
      <td id="T_1666d_row23_col3" class="data row23 col3" >0</td>
      <td id="T_1666d_row23_col4" class="data row23 col4" >0</td>
      <td id="T_1666d_row23_col5" class="data row23 col5" >0</td>
      <td id="T_1666d_row23_col6" class="data row23 col6" >0</td>
      <td id="T_1666d_row23_col7" class="data row23 col7" >0</td>
      <td id="T_1666d_row23_col8" class="data row23 col8" >0</td>
      <td id="T_1666d_row23_col9" class="data row23 col9" >0</td>
      <td id="T_1666d_row23_col10" class="data row23 col10" >0</td>
      <td id="T_1666d_row23_col11" class="data row23 col11" >0</td>
      <td id="T_1666d_row23_col12" class="data row23 col12" >0</td>
      <td id="T_1666d_row23_col13" class="data row23 col13" >0</td>
      <td id="T_1666d_row23_col14" class="data row23 col14" >0</td>
      <td id="T_1666d_row23_col15" class="data row23 col15" >0</td>
      <td id="T_1666d_row23_col16" class="data row23 col16" >0</td>
      <td id="T_1666d_row23_col17" class="data row23 col17" >0</td>
      <td id="T_1666d_row23_col18" class="data row23 col18" >0</td>
      <td id="T_1666d_row23_col19" class="data row23 col19" >0</td>
      <td id="T_1666d_row23_col20" class="data row23 col20" >0</td>
      <td id="T_1666d_row23_col21" class="data row23 col21" >0</td>
      <td id="T_1666d_row23_col22" class="data row23 col22" >0</td>
      <td id="T_1666d_row23_col23" class="data row23 col23" >0</td>
      <td id="T_1666d_row23_col24" class="data row23 col24" >0</td>
      <td id="T_1666d_row23_col25" class="data row23 col25" >0</td>
      <td id="T_1666d_row23_col26" class="data row23 col26" >0</td>
      <td id="T_1666d_row23_col27" class="data row23 col27" >0</td>
      <td id="T_1666d_row23_col28" class="data row23 col28" >8,777</td>
      <td id="T_1666d_row23_col29" class="data row23 col29" >0</td>
      <td id="T_1666d_row23_col30" class="data row23 col30" >8,777</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row24" class="row_heading level2 row24" >C</th>
      <td id="T_1666d_row24_col0" class="data row24 col0" >0</td>
      <td id="T_1666d_row24_col1" class="data row24 col1" >138,854</td>
      <td id="T_1666d_row24_col2" class="data row24 col2" >0</td>
      <td id="T_1666d_row24_col3" class="data row24 col3" >32,173</td>
      <td id="T_1666d_row24_col4" class="data row24 col4" >0</td>
      <td id="T_1666d_row24_col5" class="data row24 col5" >0</td>
      <td id="T_1666d_row24_col6" class="data row24 col6" >242,114</td>
      <td id="T_1666d_row24_col7" class="data row24 col7" >304,275</td>
      <td id="T_1666d_row24_col8" class="data row24 col8" >0</td>
      <td id="T_1666d_row24_col9" class="data row24 col9" >92,940</td>
      <td id="T_1666d_row24_col10" class="data row24 col10" >319,306</td>
      <td id="T_1666d_row24_col11" class="data row24 col11" >202,869</td>
      <td id="T_1666d_row24_col12" class="data row24 col12" >2,000</td>
      <td id="T_1666d_row24_col13" class="data row24 col13" >150,911</td>
      <td id="T_1666d_row24_col14" class="data row24 col14" >129,481</td>
      <td id="T_1666d_row24_col15" class="data row24 col15" >257,111</td>
      <td id="T_1666d_row24_col16" class="data row24 col16" >26,283</td>
      <td id="T_1666d_row24_col17" class="data row24 col17" >0</td>
      <td id="T_1666d_row24_col18" class="data row24 col18" >0</td>
      <td id="T_1666d_row24_col19" class="data row24 col19" >0</td>
      <td id="T_1666d_row24_col20" class="data row24 col20" >0</td>
      <td id="T_1666d_row24_col21" class="data row24 col21" >402,883</td>
      <td id="T_1666d_row24_col22" class="data row24 col22" >300,478</td>
      <td id="T_1666d_row24_col23" class="data row24 col23" >186,777</td>
      <td id="T_1666d_row24_col24" class="data row24 col24" >357,970</td>
      <td id="T_1666d_row24_col25" class="data row24 col25" >0</td>
      <td id="T_1666d_row24_col26" class="data row24 col26" >0</td>
      <td id="T_1666d_row24_col27" class="data row24 col27" >415,819</td>
      <td id="T_1666d_row24_col28" class="data row24 col28" >403,200</td>
      <td id="T_1666d_row24_col29" class="data row24 col29" >652,800</td>
      <td id="T_1666d_row24_col30" class="data row24 col30" >4,618,244</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row25" class="row_heading level2 row25" >F</th>
      <td id="T_1666d_row25_col0" class="data row25 col0" >0</td>
      <td id="T_1666d_row25_col1" class="data row25 col1" >77,781</td>
      <td id="T_1666d_row25_col2" class="data row25 col2" >94,519</td>
      <td id="T_1666d_row25_col3" class="data row25 col3" >0</td>
      <td id="T_1666d_row25_col4" class="data row25 col4" >0</td>
      <td id="T_1666d_row25_col5" class="data row25 col5" >0</td>
      <td id="T_1666d_row25_col6" class="data row25 col6" >0</td>
      <td id="T_1666d_row25_col7" class="data row25 col7" >0</td>
      <td id="T_1666d_row25_col8" class="data row25 col8" >0</td>
      <td id="T_1666d_row25_col9" class="data row25 col9" >0</td>
      <td id="T_1666d_row25_col10" class="data row25 col10" >0</td>
      <td id="T_1666d_row25_col11" class="data row25 col11" >0</td>
      <td id="T_1666d_row25_col12" class="data row25 col12" >9,852</td>
      <td id="T_1666d_row25_col13" class="data row25 col13" >28,874</td>
      <td id="T_1666d_row25_col14" class="data row25 col14" >0</td>
      <td id="T_1666d_row25_col15" class="data row25 col15" >90,448</td>
      <td id="T_1666d_row25_col16" class="data row25 col16" >2,000</td>
      <td id="T_1666d_row25_col17" class="data row25 col17" >132,606</td>
      <td id="T_1666d_row25_col18" class="data row25 col18" >0</td>
      <td id="T_1666d_row25_col19" class="data row25 col19" >0</td>
      <td id="T_1666d_row25_col20" class="data row25 col20" >101,844</td>
      <td id="T_1666d_row25_col21" class="data row25 col21" >0</td>
      <td id="T_1666d_row25_col22" class="data row25 col22" >120,694</td>
      <td id="T_1666d_row25_col23" class="data row25 col23" >0</td>
      <td id="T_1666d_row25_col24" class="data row25 col24" >517,251</td>
      <td id="T_1666d_row25_col25" class="data row25 col25" >0</td>
      <td id="T_1666d_row25_col26" class="data row25 col26" >0</td>
      <td id="T_1666d_row25_col27" class="data row25 col27" >0</td>
      <td id="T_1666d_row25_col28" class="data row25 col28" >269,637</td>
      <td id="T_1666d_row25_col29" class="data row25 col29" >0</td>
      <td id="T_1666d_row25_col30" class="data row25 col30" >1,445,506</td>
    </tr>
    <tr>
      <th id="T_1666d_level0_row26" class="row_heading level0 row26" rowspan="4">Man1</th>
      <th id="T_1666d_level1_row26" class="row_heading level1 row26" rowspan="4">Day</th>
      <th id="T_1666d_level2_row26" class="row_heading level2 row26" >B</th>
      <td id="T_1666d_row26_col0" class="data row26 col0" >79,177</td>
      <td id="T_1666d_row26_col1" class="data row26 col1" >79,059</td>
      <td id="T_1666d_row26_col2" class="data row26 col2" >0</td>
      <td id="T_1666d_row26_col3" class="data row26 col3" >0</td>
      <td id="T_1666d_row26_col4" class="data row26 col4" >0</td>
      <td id="T_1666d_row26_col5" class="data row26 col5" >0</td>
      <td id="T_1666d_row26_col6" class="data row26 col6" >0</td>
      <td id="T_1666d_row26_col7" class="data row26 col7" >0</td>
      <td id="T_1666d_row26_col8" class="data row26 col8" >134,400</td>
      <td id="T_1666d_row26_col9" class="data row26 col9" >250</td>
      <td id="T_1666d_row26_col10" class="data row26 col10" >0</td>
      <td id="T_1666d_row26_col11" class="data row26 col11" >0</td>
      <td id="T_1666d_row26_col12" class="data row26 col12" >0</td>
      <td id="T_1666d_row26_col13" class="data row26 col13" >0</td>
      <td id="T_1666d_row26_col14" class="data row26 col14" >79,059</td>
      <td id="T_1666d_row26_col15" class="data row26 col15" >0</td>
      <td id="T_1666d_row26_col16" class="data row26 col16" >52,776</td>
      <td id="T_1666d_row26_col17" class="data row26 col17" >91,411</td>
      <td id="T_1666d_row26_col18" class="data row26 col18" >0</td>
      <td id="T_1666d_row26_col19" class="data row26 col19" >0</td>
      <td id="T_1666d_row26_col20" class="data row26 col20" >79,059</td>
      <td id="T_1666d_row26_col21" class="data row26 col21" >0</td>
      <td id="T_1666d_row26_col22" class="data row26 col22" >0</td>
      <td id="T_1666d_row26_col23" class="data row26 col23" >0</td>
      <td id="T_1666d_row26_col24" class="data row26 col24" >0</td>
      <td id="T_1666d_row26_col25" class="data row26 col25" >0</td>
      <td id="T_1666d_row26_col26" class="data row26 col26" >0</td>
      <td id="T_1666d_row26_col27" class="data row26 col27" >0</td>
      <td id="T_1666d_row26_col28" class="data row26 col28" >0</td>
      <td id="T_1666d_row26_col29" class="data row26 col29" >0</td>
      <td id="T_1666d_row26_col30" class="data row26 col30" >595,191</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row27" class="row_heading level2 row27" >D</th>
      <td id="T_1666d_row27_col0" class="data row27 col0" >0</td>
      <td id="T_1666d_row27_col1" class="data row27 col1" >0</td>
      <td id="T_1666d_row27_col2" class="data row27 col2" >16,284</td>
      <td id="T_1666d_row27_col3" class="data row27 col3" >0</td>
      <td id="T_1666d_row27_col4" class="data row27 col4" >0</td>
      <td id="T_1666d_row27_col5" class="data row27 col5" >0</td>
      <td id="T_1666d_row27_col6" class="data row27 col6" >0</td>
      <td id="T_1666d_row27_col7" class="data row27 col7" >0</td>
      <td id="T_1666d_row27_col8" class="data row27 col8" >0</td>
      <td id="T_1666d_row27_col9" class="data row27 col9" >0</td>
      <td id="T_1666d_row27_col10" class="data row27 col10" >0</td>
      <td id="T_1666d_row27_col11" class="data row27 col11" >0</td>
      <td id="T_1666d_row27_col12" class="data row27 col12" >0</td>
      <td id="T_1666d_row27_col13" class="data row27 col13" >3,281</td>
      <td id="T_1666d_row27_col14" class="data row27 col14" >0</td>
      <td id="T_1666d_row27_col15" class="data row27 col15" >250</td>
      <td id="T_1666d_row27_col16" class="data row27 col16" >0</td>
      <td id="T_1666d_row27_col17" class="data row27 col17" >0</td>
      <td id="T_1666d_row27_col18" class="data row27 col18" >0</td>
      <td id="T_1666d_row27_col19" class="data row27 col19" >0</td>
      <td id="T_1666d_row27_col20" class="data row27 col20" >0</td>
      <td id="T_1666d_row27_col21" class="data row27 col21" >0</td>
      <td id="T_1666d_row27_col22" class="data row27 col22" >0</td>
      <td id="T_1666d_row27_col23" class="data row27 col23" >18,772</td>
      <td id="T_1666d_row27_col24" class="data row27 col24" >0</td>
      <td id="T_1666d_row27_col25" class="data row27 col25" >0</td>
      <td id="T_1666d_row27_col26" class="data row27 col26" >0</td>
      <td id="T_1666d_row27_col27" class="data row27 col27" >0</td>
      <td id="T_1666d_row27_col28" class="data row27 col28" >49,170</td>
      <td id="T_1666d_row27_col29" class="data row27 col29" >13,170</td>
      <td id="T_1666d_row27_col30" class="data row27 col30" >100,927</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row28" class="row_heading level2 row28" >A</th>
      <td id="T_1666d_row28_col0" class="data row28 col0" >0</td>
      <td id="T_1666d_row28_col1" class="data row28 col1" >0</td>
      <td id="T_1666d_row28_col2" class="data row28 col2" >59,851</td>
      <td id="T_1666d_row28_col3" class="data row28 col3" >91,924</td>
      <td id="T_1666d_row28_col4" class="data row28 col4" >0</td>
      <td id="T_1666d_row28_col5" class="data row28 col5" >0</td>
      <td id="T_1666d_row28_col6" class="data row28 col6" >0</td>
      <td id="T_1666d_row28_col7" class="data row28 col7" >0</td>
      <td id="T_1666d_row28_col8" class="data row28 col8" >0</td>
      <td id="T_1666d_row28_col9" class="data row28 col9" >0</td>
      <td id="T_1666d_row28_col10" class="data row28 col10" >0</td>
      <td id="T_1666d_row28_col11" class="data row28 col11" >0</td>
      <td id="T_1666d_row28_col12" class="data row28 col12" >79,059</td>
      <td id="T_1666d_row28_col13" class="data row28 col13" >0</td>
      <td id="T_1666d_row28_col14" class="data row28 col14" >0</td>
      <td id="T_1666d_row28_col15" class="data row28 col15" >0</td>
      <td id="T_1666d_row28_col16" class="data row28 col16" >0</td>
      <td id="T_1666d_row28_col17" class="data row28 col17" >0</td>
      <td id="T_1666d_row28_col18" class="data row28 col18" >0</td>
      <td id="T_1666d_row28_col19" class="data row28 col19" >0</td>
      <td id="T_1666d_row28_col20" class="data row28 col20" >0</td>
      <td id="T_1666d_row28_col21" class="data row28 col21" >0</td>
      <td id="T_1666d_row28_col22" class="data row28 col22" >0</td>
      <td id="T_1666d_row28_col23" class="data row28 col23" >27,747</td>
      <td id="T_1666d_row28_col24" class="data row28 col24" >0</td>
      <td id="T_1666d_row28_col25" class="data row28 col25" >0</td>
      <td id="T_1666d_row28_col26" class="data row28 col26" >0</td>
      <td id="T_1666d_row28_col27" class="data row28 col27" >0</td>
      <td id="T_1666d_row28_col28" class="data row28 col28" >0</td>
      <td id="T_1666d_row28_col29" class="data row28 col29" >0</td>
      <td id="T_1666d_row28_col30" class="data row28 col30" >258,581</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row29" class="row_heading level2 row29" >C</th>
      <td id="T_1666d_row29_col0" class="data row29 col0" >0</td>
      <td id="T_1666d_row29_col1" class="data row29 col1" >0</td>
      <td id="T_1666d_row29_col2" class="data row29 col2" >3,154</td>
      <td id="T_1666d_row29_col3" class="data row29 col3" >0</td>
      <td id="T_1666d_row29_col4" class="data row29 col4" >0</td>
      <td id="T_1666d_row29_col5" class="data row29 col5" >0</td>
      <td id="T_1666d_row29_col6" class="data row29 col6" >79,059</td>
      <td id="T_1666d_row29_col7" class="data row29 col7" >79,059</td>
      <td id="T_1666d_row29_col8" class="data row29 col8" >0</td>
      <td id="T_1666d_row29_col9" class="data row29 col9" >78,809</td>
      <td id="T_1666d_row29_col10" class="data row29 col10" >79,059</td>
      <td id="T_1666d_row29_col11" class="data row29 col11" >79,059</td>
      <td id="T_1666d_row29_col12" class="data row29 col12" >0</td>
      <td id="T_1666d_row29_col13" class="data row29 col13" >70,093</td>
      <td id="T_1666d_row29_col14" class="data row29 col14" >0</td>
      <td id="T_1666d_row29_col15" class="data row29 col15" >78,572</td>
      <td id="T_1666d_row29_col16" class="data row29 col16" >26,283</td>
      <td id="T_1666d_row29_col17" class="data row29 col17" >0</td>
      <td id="T_1666d_row29_col18" class="data row29 col18" >0</td>
      <td id="T_1666d_row29_col19" class="data row29 col19" >0</td>
      <td id="T_1666d_row29_col20" class="data row29 col20" >0</td>
      <td id="T_1666d_row29_col21" class="data row29 col21" >120,602</td>
      <td id="T_1666d_row29_col22" class="data row29 col22" >134,400</td>
      <td id="T_1666d_row29_col23" class="data row29 col23" >250</td>
      <td id="T_1666d_row29_col24" class="data row29 col24" >79,058</td>
      <td id="T_1666d_row29_col25" class="data row29 col25" >0</td>
      <td id="T_1666d_row29_col26" class="data row29 col26" >0</td>
      <td id="T_1666d_row29_col27" class="data row29 col27" >134,400</td>
      <td id="T_1666d_row29_col28" class="data row29 col28" >0</td>
      <td id="T_1666d_row29_col29" class="data row29 col29" >43,061</td>
      <td id="T_1666d_row29_col30" class="data row29 col30" >1,084,918</td>
    </tr>
    <tr>
      <th id="T_1666d_level0_row30" class="row_heading level0 row30" rowspan="9">Man2</th>
      <th id="T_1666d_level1_row30" class="row_heading level1 row30" rowspan="4">Sday_1</th>
      <th id="T_1666d_level2_row30" class="row_heading level2 row30" >B</th>
      <td id="T_1666d_row30_col0" class="data row30 col0" >69,177</td>
      <td id="T_1666d_row30_col1" class="data row30 col1" >0</td>
      <td id="T_1666d_row30_col2" class="data row30 col2" >0</td>
      <td id="T_1666d_row30_col3" class="data row30 col3" >34,001</td>
      <td id="T_1666d_row30_col4" class="data row30 col4" >0</td>
      <td id="T_1666d_row30_col5" class="data row30 col5" >0</td>
      <td id="T_1666d_row30_col6" class="data row30 col6" >0</td>
      <td id="T_1666d_row30_col7" class="data row30 col7" >0</td>
      <td id="T_1666d_row30_col8" class="data row30 col8" >69,177</td>
      <td id="T_1666d_row30_col9" class="data row30 col9" >17,079</td>
      <td id="T_1666d_row30_col10" class="data row30 col10" >0</td>
      <td id="T_1666d_row30_col11" class="data row30 col11" >0</td>
      <td id="T_1666d_row30_col12" class="data row30 col12" >0</td>
      <td id="T_1666d_row30_col13" class="data row30 col13" >0</td>
      <td id="T_1666d_row30_col14" class="data row30 col14" >0</td>
      <td id="T_1666d_row30_col15" class="data row30 col15" >0</td>
      <td id="T_1666d_row30_col16" class="data row30 col16" >68,689</td>
      <td id="T_1666d_row30_col17" class="data row30 col17" >52,466</td>
      <td id="T_1666d_row30_col18" class="data row30 col18" >0</td>
      <td id="T_1666d_row30_col19" class="data row30 col19" >0</td>
      <td id="T_1666d_row30_col20" class="data row30 col20" >69,176</td>
      <td id="T_1666d_row30_col21" class="data row30 col21" >0</td>
      <td id="T_1666d_row30_col22" class="data row30 col22" >0</td>
      <td id="T_1666d_row30_col23" class="data row30 col23" >0</td>
      <td id="T_1666d_row30_col24" class="data row30 col24" >0</td>
      <td id="T_1666d_row30_col25" class="data row30 col25" >0</td>
      <td id="T_1666d_row30_col26" class="data row30 col26" >0</td>
      <td id="T_1666d_row30_col27" class="data row30 col27" >0</td>
      <td id="T_1666d_row30_col28" class="data row30 col28" >69,177</td>
      <td id="T_1666d_row30_col29" class="data row30 col29" >0</td>
      <td id="T_1666d_row30_col30" class="data row30 col30" >448,942</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row31" class="row_heading level2 row31" >D</th>
      <td id="T_1666d_row31_col0" class="data row31 col0" >0</td>
      <td id="T_1666d_row31_col1" class="data row31 col1" >25,309</td>
      <td id="T_1666d_row31_col2" class="data row31 col2" >0</td>
      <td id="T_1666d_row31_col3" class="data row31 col3" >0</td>
      <td id="T_1666d_row31_col4" class="data row31 col4" >0</td>
      <td id="T_1666d_row31_col5" class="data row31 col5" >0</td>
      <td id="T_1666d_row31_col6" class="data row31 col6" >17,715</td>
      <td id="T_1666d_row31_col7" class="data row31 col7" >0</td>
      <td id="T_1666d_row31_col8" class="data row31 col8" >0</td>
      <td id="T_1666d_row31_col9" class="data row31 col9" >0</td>
      <td id="T_1666d_row31_col10" class="data row31 col10" >0</td>
      <td id="T_1666d_row31_col11" class="data row31 col11" >0</td>
      <td id="T_1666d_row31_col12" class="data row31 col12" >0</td>
      <td id="T_1666d_row31_col13" class="data row31 col13" >2,038</td>
      <td id="T_1666d_row31_col14" class="data row31 col14" >0</td>
      <td id="T_1666d_row31_col15" class="data row31 col15" >1</td>
      <td id="T_1666d_row31_col16" class="data row31 col16" >0</td>
      <td id="T_1666d_row31_col17" class="data row31 col17" >0</td>
      <td id="T_1666d_row31_col18" class="data row31 col18" >0</td>
      <td id="T_1666d_row31_col19" class="data row31 col19" >0</td>
      <td id="T_1666d_row31_col20" class="data row31 col20" >0</td>
      <td id="T_1666d_row31_col21" class="data row31 col21" >0</td>
      <td id="T_1666d_row31_col22" class="data row31 col22" >0</td>
      <td id="T_1666d_row31_col23" class="data row31 col23" >0</td>
      <td id="T_1666d_row31_col24" class="data row31 col24" >0</td>
      <td id="T_1666d_row31_col25" class="data row31 col25" >0</td>
      <td id="T_1666d_row31_col26" class="data row31 col26" >0</td>
      <td id="T_1666d_row31_col27" class="data row31 col27" >0</td>
      <td id="T_1666d_row31_col28" class="data row31 col28" >0</td>
      <td id="T_1666d_row31_col29" class="data row31 col29" >0</td>
      <td id="T_1666d_row31_col30" class="data row31 col30" >45,063</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row32" class="row_heading level2 row32" >A</th>
      <td id="T_1666d_row32_col0" class="data row32 col0" >0</td>
      <td id="T_1666d_row32_col1" class="data row32 col1" >0</td>
      <td id="T_1666d_row32_col2" class="data row32 col2" >0</td>
      <td id="T_1666d_row32_col3" class="data row32 col3" >0</td>
      <td id="T_1666d_row32_col4" class="data row32 col4" >0</td>
      <td id="T_1666d_row32_col5" class="data row32 col5" >0</td>
      <td id="T_1666d_row32_col6" class="data row32 col6" >0</td>
      <td id="T_1666d_row32_col7" class="data row32 col7" >0</td>
      <td id="T_1666d_row32_col8" class="data row32 col8" >0</td>
      <td id="T_1666d_row32_col9" class="data row32 col9" >52,099</td>
      <td id="T_1666d_row32_col10" class="data row32 col10" >0</td>
      <td id="T_1666d_row32_col11" class="data row32 col11" >0</td>
      <td id="T_1666d_row32_col12" class="data row32 col12" >69,177</td>
      <td id="T_1666d_row32_col13" class="data row32 col13" >69,176</td>
      <td id="T_1666d_row32_col14" class="data row32 col14" >0</td>
      <td id="T_1666d_row32_col15" class="data row32 col15" >0</td>
      <td id="T_1666d_row32_col16" class="data row32 col16" >0</td>
      <td id="T_1666d_row32_col17" class="data row32 col17" >16,711</td>
      <td id="T_1666d_row32_col18" class="data row32 col18" >0</td>
      <td id="T_1666d_row32_col19" class="data row32 col19" >0</td>
      <td id="T_1666d_row32_col20" class="data row32 col20" >0</td>
      <td id="T_1666d_row32_col21" class="data row32 col21" >0</td>
      <td id="T_1666d_row32_col22" class="data row32 col22" >0</td>
      <td id="T_1666d_row32_col23" class="data row32 col23" >0</td>
      <td id="T_1666d_row32_col24" class="data row32 col24" >0</td>
      <td id="T_1666d_row32_col25" class="data row32 col25" >0</td>
      <td id="T_1666d_row32_col26" class="data row32 col26" >0</td>
      <td id="T_1666d_row32_col27" class="data row32 col27" >0</td>
      <td id="T_1666d_row32_col28" class="data row32 col28" >0</td>
      <td id="T_1666d_row32_col29" class="data row32 col29" >0</td>
      <td id="T_1666d_row32_col30" class="data row32 col30" >207,163</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row33" class="row_heading level2 row33" >C</th>
      <td id="T_1666d_row33_col0" class="data row33 col0" >0</td>
      <td id="T_1666d_row33_col1" class="data row33 col1" >0</td>
      <td id="T_1666d_row33_col2" class="data row33 col2" >69,177</td>
      <td id="T_1666d_row33_col3" class="data row33 col3" >35,176</td>
      <td id="T_1666d_row33_col4" class="data row33 col4" >0</td>
      <td id="T_1666d_row33_col5" class="data row33 col5" >0</td>
      <td id="T_1666d_row33_col6" class="data row33 col6" >69,176</td>
      <td id="T_1666d_row33_col7" class="data row33 col7" >69,177</td>
      <td id="T_1666d_row33_col8" class="data row33 col8" >0</td>
      <td id="T_1666d_row33_col9" class="data row33 col9" >0</td>
      <td id="T_1666d_row33_col10" class="data row33 col10" >69,177</td>
      <td id="T_1666d_row33_col11" class="data row33 col11" >69,177</td>
      <td id="T_1666d_row33_col12" class="data row33 col12" >0</td>
      <td id="T_1666d_row33_col13" class="data row33 col13" >0</td>
      <td id="T_1666d_row33_col14" class="data row33 col14" >69,177</td>
      <td id="T_1666d_row33_col15" class="data row33 col15" >69,176</td>
      <td id="T_1666d_row33_col16" class="data row33 col16" >487</td>
      <td id="T_1666d_row33_col17" class="data row33 col17" >0</td>
      <td id="T_1666d_row33_col18" class="data row33 col18" >0</td>
      <td id="T_1666d_row33_col19" class="data row33 col19" >0</td>
      <td id="T_1666d_row33_col20" class="data row33 col20" >0</td>
      <td id="T_1666d_row33_col21" class="data row33 col21" >117,600</td>
      <td id="T_1666d_row33_col22" class="data row33 col22" >117,600</td>
      <td id="T_1666d_row33_col23" class="data row33 col23" >69,177</td>
      <td id="T_1666d_row33_col24" class="data row33 col24" >117,600</td>
      <td id="T_1666d_row33_col25" class="data row33 col25" >0</td>
      <td id="T_1666d_row33_col26" class="data row33 col26" >0</td>
      <td id="T_1666d_row33_col27" class="data row33 col27" >69,177</td>
      <td id="T_1666d_row33_col28" class="data row33 col28" >0</td>
      <td id="T_1666d_row33_col29" class="data row33 col29" >168,000</td>
      <td id="T_1666d_row33_col30" class="data row33 col30" >1,179,054</td>
    </tr>
    <tr>
      <th id="T_1666d_level1_row34" class="row_heading level1 row34" rowspan="5">Sday_2</th>
      <th id="T_1666d_level2_row34" class="row_heading level2 row34" >B</th>
      <td id="T_1666d_row34_col0" class="data row34 col0" >69,177</td>
      <td id="T_1666d_row34_col1" class="data row34 col1" >67,756</td>
      <td id="T_1666d_row34_col2" class="data row34 col2" >0</td>
      <td id="T_1666d_row34_col3" class="data row34 col3" >69,177</td>
      <td id="T_1666d_row34_col4" class="data row34 col4" >0</td>
      <td id="T_1666d_row34_col5" class="data row34 col5" >0</td>
      <td id="T_1666d_row34_col6" class="data row34 col6" >0</td>
      <td id="T_1666d_row34_col7" class="data row34 col7" >0</td>
      <td id="T_1666d_row34_col8" class="data row34 col8" >69,176</td>
      <td id="T_1666d_row34_col9" class="data row34 col9" >17,079</td>
      <td id="T_1666d_row34_col10" class="data row34 col10" >0</td>
      <td id="T_1666d_row34_col11" class="data row34 col11" >0</td>
      <td id="T_1666d_row34_col12" class="data row34 col12" >0</td>
      <td id="T_1666d_row34_col13" class="data row34 col13" >0</td>
      <td id="T_1666d_row34_col14" class="data row34 col14" >69,666</td>
      <td id="T_1666d_row34_col15" class="data row34 col15" >0</td>
      <td id="T_1666d_row34_col16" class="data row34 col16" >69,176</td>
      <td id="T_1666d_row34_col17" class="data row34 col17" >0</td>
      <td id="T_1666d_row34_col18" class="data row34 col18" >0</td>
      <td id="T_1666d_row34_col19" class="data row34 col19" >0</td>
      <td id="T_1666d_row34_col20" class="data row34 col20" >69,176</td>
      <td id="T_1666d_row34_col21" class="data row34 col21" >0</td>
      <td id="T_1666d_row34_col22" class="data row34 col22" >0</td>
      <td id="T_1666d_row34_col23" class="data row34 col23" >0</td>
      <td id="T_1666d_row34_col24" class="data row34 col24" >0</td>
      <td id="T_1666d_row34_col25" class="data row34 col25" >0</td>
      <td id="T_1666d_row34_col26" class="data row34 col26" >0</td>
      <td id="T_1666d_row34_col27" class="data row34 col27" >0</td>
      <td id="T_1666d_row34_col28" class="data row34 col28" >30,332</td>
      <td id="T_1666d_row34_col29" class="data row34 col29" >0</td>
      <td id="T_1666d_row34_col30" class="data row34 col30" >530,715</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row35" class="row_heading level2 row35" >D</th>
      <td id="T_1666d_row35_col0" class="data row35 col0" >0</td>
      <td id="T_1666d_row35_col1" class="data row35 col1" >520</td>
      <td id="T_1666d_row35_col2" class="data row35 col2" >0</td>
      <td id="T_1666d_row35_col3" class="data row35 col3" >0</td>
      <td id="T_1666d_row35_col4" class="data row35 col4" >0</td>
      <td id="T_1666d_row35_col5" class="data row35 col5" >0</td>
      <td id="T_1666d_row35_col6" class="data row35 col6" >0</td>
      <td id="T_1666d_row35_col7" class="data row35 col7" >0</td>
      <td id="T_1666d_row35_col8" class="data row35 col8" >0</td>
      <td id="T_1666d_row35_col9" class="data row35 col9" >13,819</td>
      <td id="T_1666d_row35_col10" class="data row35 col10" >0</td>
      <td id="T_1666d_row35_col11" class="data row35 col11" >0</td>
      <td id="T_1666d_row35_col12" class="data row35 col12" >0</td>
      <td id="T_1666d_row35_col13" class="data row35 col13" >0</td>
      <td id="T_1666d_row35_col14" class="data row35 col14" >0</td>
      <td id="T_1666d_row35_col15" class="data row35 col15" >3,192</td>
      <td id="T_1666d_row35_col16" class="data row35 col16" >0</td>
      <td id="T_1666d_row35_col17" class="data row35 col17" >0</td>
      <td id="T_1666d_row35_col18" class="data row35 col18" >0</td>
      <td id="T_1666d_row35_col19" class="data row35 col19" >0</td>
      <td id="T_1666d_row35_col20" class="data row35 col20" >0</td>
      <td id="T_1666d_row35_col21" class="data row35 col21" >0</td>
      <td id="T_1666d_row35_col22" class="data row35 col22" >20,190</td>
      <td id="T_1666d_row35_col23" class="data row35 col23" >0</td>
      <td id="T_1666d_row35_col24" class="data row35 col24" >0</td>
      <td id="T_1666d_row35_col25" class="data row35 col25" >0</td>
      <td id="T_1666d_row35_col26" class="data row35 col26" >0</td>
      <td id="T_1666d_row35_col27" class="data row35 col27" >0</td>
      <td id="T_1666d_row35_col28" class="data row35 col28" >16,690</td>
      <td id="T_1666d_row35_col29" class="data row35 col29" >0</td>
      <td id="T_1666d_row35_col30" class="data row35 col30" >54,411</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row36" class="row_heading level2 row36" >A</th>
      <td id="T_1666d_row36_col0" class="data row36 col0" >0</td>
      <td id="T_1666d_row36_col1" class="data row36 col1" >0</td>
      <td id="T_1666d_row36_col2" class="data row36 col2" >19,497</td>
      <td id="T_1666d_row36_col3" class="data row36 col3" >0</td>
      <td id="T_1666d_row36_col4" class="data row36 col4" >0</td>
      <td id="T_1666d_row36_col5" class="data row36 col5" >0</td>
      <td id="T_1666d_row36_col6" class="data row36 col6" >0</td>
      <td id="T_1666d_row36_col7" class="data row36 col7" >0</td>
      <td id="T_1666d_row36_col8" class="data row36 col8" >0</td>
      <td id="T_1666d_row36_col9" class="data row36 col9" >195</td>
      <td id="T_1666d_row36_col10" class="data row36 col10" >0</td>
      <td id="T_1666d_row36_col11" class="data row36 col11" >0</td>
      <td id="T_1666d_row36_col12" class="data row36 col12" >69,177</td>
      <td id="T_1666d_row36_col13" class="data row36 col13" >67,652</td>
      <td id="T_1666d_row36_col14" class="data row36 col14" >0</td>
      <td id="T_1666d_row36_col15" class="data row36 col15" >0</td>
      <td id="T_1666d_row36_col16" class="data row36 col16" >0</td>
      <td id="T_1666d_row36_col17" class="data row36 col17" >71,816</td>
      <td id="T_1666d_row36_col18" class="data row36 col18" >0</td>
      <td id="T_1666d_row36_col19" class="data row36 col19" >0</td>
      <td id="T_1666d_row36_col20" class="data row36 col20" >0</td>
      <td id="T_1666d_row36_col21" class="data row36 col21" >0</td>
      <td id="T_1666d_row36_col22" class="data row36 col22" >439</td>
      <td id="T_1666d_row36_col23" class="data row36 col23" >69,177</td>
      <td id="T_1666d_row36_col24" class="data row36 col24" >0</td>
      <td id="T_1666d_row36_col25" class="data row36 col25" >0</td>
      <td id="T_1666d_row36_col26" class="data row36 col26" >0</td>
      <td id="T_1666d_row36_col27" class="data row36 col27" >0</td>
      <td id="T_1666d_row36_col28" class="data row36 col28" >0</td>
      <td id="T_1666d_row36_col29" class="data row36 col29" >0</td>
      <td id="T_1666d_row36_col30" class="data row36 col30" >297,953</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row37" class="row_heading level2 row37" >E</th>
      <td id="T_1666d_row37_col0" class="data row37 col0" >0</td>
      <td id="T_1666d_row37_col1" class="data row37 col1" >0</td>
      <td id="T_1666d_row37_col2" class="data row37 col2" >0</td>
      <td id="T_1666d_row37_col3" class="data row37 col3" >0</td>
      <td id="T_1666d_row37_col4" class="data row37 col4" >0</td>
      <td id="T_1666d_row37_col5" class="data row37 col5" >0</td>
      <td id="T_1666d_row37_col6" class="data row37 col6" >0</td>
      <td id="T_1666d_row37_col7" class="data row37 col7" >0</td>
      <td id="T_1666d_row37_col8" class="data row37 col8" >0</td>
      <td id="T_1666d_row37_col9" class="data row37 col9" >0</td>
      <td id="T_1666d_row37_col10" class="data row37 col10" >0</td>
      <td id="T_1666d_row37_col11" class="data row37 col11" >0</td>
      <td id="T_1666d_row37_col12" class="data row37 col12" >0</td>
      <td id="T_1666d_row37_col13" class="data row37 col13" >0</td>
      <td id="T_1666d_row37_col14" class="data row37 col14" >0</td>
      <td id="T_1666d_row37_col15" class="data row37 col15" >0</td>
      <td id="T_1666d_row37_col16" class="data row37 col16" >0</td>
      <td id="T_1666d_row37_col17" class="data row37 col17" >0</td>
      <td id="T_1666d_row37_col18" class="data row37 col18" >0</td>
      <td id="T_1666d_row37_col19" class="data row37 col19" >0</td>
      <td id="T_1666d_row37_col20" class="data row37 col20" >0</td>
      <td id="T_1666d_row37_col21" class="data row37 col21" >0</td>
      <td id="T_1666d_row37_col22" class="data row37 col22" >0</td>
      <td id="T_1666d_row37_col23" class="data row37 col23" >0</td>
      <td id="T_1666d_row37_col24" class="data row37 col24" >0</td>
      <td id="T_1666d_row37_col25" class="data row37 col25" >0</td>
      <td id="T_1666d_row37_col26" class="data row37 col26" >0</td>
      <td id="T_1666d_row37_col27" class="data row37 col27" >0</td>
      <td id="T_1666d_row37_col28" class="data row37 col28" >8,777</td>
      <td id="T_1666d_row37_col29" class="data row37 col29" >0</td>
      <td id="T_1666d_row37_col30" class="data row37 col30" >8,777</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row38" class="row_heading level2 row38" >C</th>
      <td id="T_1666d_row38_col0" class="data row38 col0" >0</td>
      <td id="T_1666d_row38_col1" class="data row38 col1" >0</td>
      <td id="T_1666d_row38_col2" class="data row38 col2" >49,680</td>
      <td id="T_1666d_row38_col3" class="data row38 col3" >0</td>
      <td id="T_1666d_row38_col4" class="data row38 col4" >0</td>
      <td id="T_1666d_row38_col5" class="data row38 col5" >0</td>
      <td id="T_1666d_row38_col6" class="data row38 col6" >69,177</td>
      <td id="T_1666d_row38_col7" class="data row38 col7" >69,177</td>
      <td id="T_1666d_row38_col8" class="data row38 col8" >0</td>
      <td id="T_1666d_row38_col9" class="data row38 col9" >14,131</td>
      <td id="T_1666d_row38_col10" class="data row38 col10" >69,177</td>
      <td id="T_1666d_row38_col11" class="data row38 col11" >69,177</td>
      <td id="T_1666d_row38_col12" class="data row38 col12" >0</td>
      <td id="T_1666d_row38_col13" class="data row38 col13" >1,525</td>
      <td id="T_1666d_row38_col14" class="data row38 col14" >0</td>
      <td id="T_1666d_row38_col15" class="data row38 col15" >108,875</td>
      <td id="T_1666d_row38_col16" class="data row38 col16" >1</td>
      <td id="T_1666d_row38_col17" class="data row38 col17" >0</td>
      <td id="T_1666d_row38_col18" class="data row38 col18" >0</td>
      <td id="T_1666d_row38_col19" class="data row38 col19" >0</td>
      <td id="T_1666d_row38_col20" class="data row38 col20" >0</td>
      <td id="T_1666d_row38_col21" class="data row38 col21" >117,600</td>
      <td id="T_1666d_row38_col22" class="data row38 col22" >13,552</td>
      <td id="T_1666d_row38_col23" class="data row38 col23" >0</td>
      <td id="T_1666d_row38_col24" class="data row38 col24" >117,600</td>
      <td id="T_1666d_row38_col25" class="data row38 col25" >0</td>
      <td id="T_1666d_row38_col26" class="data row38 col26" >0</td>
      <td id="T_1666d_row38_col27" class="data row38 col27" >117,600</td>
      <td id="T_1666d_row38_col28" class="data row38 col28" >0</td>
      <td id="T_1666d_row38_col29" class="data row38 col29" >168,000</td>
      <td id="T_1666d_row38_col30" class="data row38 col30" >985,272</td>
    </tr>
    <tr>
      <th id="T_1666d_level0_row39" class="row_heading level0 row39" rowspan="8">Man3</th>
      <th id="T_1666d_level1_row39" class="row_heading level1 row39" rowspan="4">Sday_1</th>
      <th id="T_1666d_level2_row39" class="row_heading level2 row39" >B</th>
      <td id="T_1666d_row39_col0" class="data row39 col0" >71,186</td>
      <td id="T_1666d_row39_col1" class="data row39 col1" >16,895</td>
      <td id="T_1666d_row39_col2" class="data row39 col2" >19,497</td>
      <td id="T_1666d_row39_col3" class="data row39 col3" >112,287</td>
      <td id="T_1666d_row39_col4" class="data row39 col4" >0</td>
      <td id="T_1666d_row39_col5" class="data row39 col5" >0</td>
      <td id="T_1666d_row39_col6" class="data row39 col6" >58,346</td>
      <td id="T_1666d_row39_col7" class="data row39 col7" >25,395</td>
      <td id="T_1666d_row39_col8" class="data row39 col8" >69,177</td>
      <td id="T_1666d_row39_col9" class="data row39 col9" >69,175</td>
      <td id="T_1666d_row39_col10" class="data row39 col10" >19,758</td>
      <td id="T_1666d_row39_col11" class="data row39 col11" >22,547</td>
      <td id="T_1666d_row39_col12" class="data row39 col12" >0</td>
      <td id="T_1666d_row39_col13" class="data row39 col13" >31,123</td>
      <td id="T_1666d_row39_col14" class="data row39 col14" >66,184</td>
      <td id="T_1666d_row39_col15" class="data row39 col15" >0</td>
      <td id="T_1666d_row39_col16" class="data row39 col16" >69,177</td>
      <td id="T_1666d_row39_col17" class="data row39 col17" >0</td>
      <td id="T_1666d_row39_col18" class="data row39 col18" >0</td>
      <td id="T_1666d_row39_col19" class="data row39 col19" >0</td>
      <td id="T_1666d_row39_col20" class="data row39 col20" >69,177</td>
      <td id="T_1666d_row39_col21" class="data row39 col21" >117,600</td>
      <td id="T_1666d_row39_col22" class="data row39 col22" >0</td>
      <td id="T_1666d_row39_col23" class="data row39 col23" >0</td>
      <td id="T_1666d_row39_col24" class="data row39 col24" >0</td>
      <td id="T_1666d_row39_col25" class="data row39 col25" >0</td>
      <td id="T_1666d_row39_col26" class="data row39 col26" >0</td>
      <td id="T_1666d_row39_col27" class="data row39 col27" >0</td>
      <td id="T_1666d_row39_col28" class="data row39 col28" >0</td>
      <td id="T_1666d_row39_col29" class="data row39 col29" >0</td>
      <td id="T_1666d_row39_col30" class="data row39 col30" >837,524</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row40" class="row_heading level2 row40" >D</th>
      <td id="T_1666d_row40_col0" class="data row40 col0" >0</td>
      <td id="T_1666d_row40_col1" class="data row40 col1" >0</td>
      <td id="T_1666d_row40_col2" class="data row40 col2" >15,142</td>
      <td id="T_1666d_row40_col3" class="data row40 col3" >0</td>
      <td id="T_1666d_row40_col4" class="data row40 col4" >0</td>
      <td id="T_1666d_row40_col5" class="data row40 col5" >0</td>
      <td id="T_1666d_row40_col6" class="data row40 col6" >1,842</td>
      <td id="T_1666d_row40_col7" class="data row40 col7" >0</td>
      <td id="T_1666d_row40_col8" class="data row40 col8" >0</td>
      <td id="T_1666d_row40_col9" class="data row40 col9" >0</td>
      <td id="T_1666d_row40_col10" class="data row40 col10" >0</td>
      <td id="T_1666d_row40_col11" class="data row40 col11" >11,423</td>
      <td id="T_1666d_row40_col12" class="data row40 col12" >0</td>
      <td id="T_1666d_row40_col13" class="data row40 col13" >0</td>
      <td id="T_1666d_row40_col14" class="data row40 col14" >0</td>
      <td id="T_1666d_row40_col15" class="data row40 col15" >5,602</td>
      <td id="T_1666d_row40_col16" class="data row40 col16" >0</td>
      <td id="T_1666d_row40_col17" class="data row40 col17" >0</td>
      <td id="T_1666d_row40_col18" class="data row40 col18" >0</td>
      <td id="T_1666d_row40_col19" class="data row40 col19" >0</td>
      <td id="T_1666d_row40_col20" class="data row40 col20" >0</td>
      <td id="T_1666d_row40_col21" class="data row40 col21" >0</td>
      <td id="T_1666d_row40_col22" class="data row40 col22" >2,387</td>
      <td id="T_1666d_row40_col23" class="data row40 col23" >0</td>
      <td id="T_1666d_row40_col24" class="data row40 col24" >25,308</td>
      <td id="T_1666d_row40_col25" class="data row40 col25" >0</td>
      <td id="T_1666d_row40_col26" class="data row40 col26" >0</td>
      <td id="T_1666d_row40_col27" class="data row40 col27" >0</td>
      <td id="T_1666d_row40_col28" class="data row40 col28" >0</td>
      <td id="T_1666d_row40_col29" class="data row40 col29" >0</td>
      <td id="T_1666d_row40_col30" class="data row40 col30" >61,704</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row41" class="row_heading level2 row41" >A</th>
      <td id="T_1666d_row41_col0" class="data row41 col0" >0</td>
      <td id="T_1666d_row41_col1" class="data row41 col1" >85,119</td>
      <td id="T_1666d_row41_col2" class="data row41 col2" >0</td>
      <td id="T_1666d_row41_col3" class="data row41 col3" >5,313</td>
      <td id="T_1666d_row41_col4" class="data row41 col4" >0</td>
      <td id="T_1666d_row41_col5" class="data row41 col5" >0</td>
      <td id="T_1666d_row41_col6" class="data row41 col6" >0</td>
      <td id="T_1666d_row41_col7" class="data row41 col7" >0</td>
      <td id="T_1666d_row41_col8" class="data row41 col8" >0</td>
      <td id="T_1666d_row41_col9" class="data row41 col9" >2</td>
      <td id="T_1666d_row41_col10" class="data row41 col10" >0</td>
      <td id="T_1666d_row41_col11" class="data row41 col11" >15,408</td>
      <td id="T_1666d_row41_col12" class="data row41 col12" >69,176</td>
      <td id="T_1666d_row41_col13" class="data row41 col13" >38,054</td>
      <td id="T_1666d_row41_col14" class="data row41 col14" >0</td>
      <td id="T_1666d_row41_col15" class="data row41 col15" >53,867</td>
      <td id="T_1666d_row41_col16" class="data row41 col16" >0</td>
      <td id="T_1666d_row41_col17" class="data row41 col17" >106,245</td>
      <td id="T_1666d_row41_col18" class="data row41 col18" >0</td>
      <td id="T_1666d_row41_col19" class="data row41 col19" >0</td>
      <td id="T_1666d_row41_col20" class="data row41 col20" >0</td>
      <td id="T_1666d_row41_col21" class="data row41 col21" >0</td>
      <td id="T_1666d_row41_col22" class="data row41 col22" >0</td>
      <td id="T_1666d_row41_col23" class="data row41 col23" >251</td>
      <td id="T_1666d_row41_col24" class="data row41 col24" >1</td>
      <td id="T_1666d_row41_col25" class="data row41 col25" >0</td>
      <td id="T_1666d_row41_col26" class="data row41 col26" >0</td>
      <td id="T_1666d_row41_col27" class="data row41 col27" >0</td>
      <td id="T_1666d_row41_col28" class="data row41 col28" >0</td>
      <td id="T_1666d_row41_col29" class="data row41 col29" >0</td>
      <td id="T_1666d_row41_col30" class="data row41 col30" >373,436</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row42" class="row_heading level2 row42" >C</th>
      <td id="T_1666d_row42_col0" class="data row42 col0" >0</td>
      <td id="T_1666d_row42_col1" class="data row42 col1" >416</td>
      <td id="T_1666d_row42_col2" class="data row42 col2" >8,290</td>
      <td id="T_1666d_row42_col3" class="data row42 col3" >0</td>
      <td id="T_1666d_row42_col4" class="data row42 col4" >0</td>
      <td id="T_1666d_row42_col5" class="data row42 col5" >0</td>
      <td id="T_1666d_row42_col6" class="data row42 col6" >11,213</td>
      <td id="T_1666d_row42_col7" class="data row42 col7" >43,782</td>
      <td id="T_1666d_row42_col8" class="data row42 col8" >0</td>
      <td id="T_1666d_row42_col9" class="data row42 col9" >0</td>
      <td id="T_1666d_row42_col10" class="data row42 col10" >49,419</td>
      <td id="T_1666d_row42_col11" class="data row42 col11" >0</td>
      <td id="T_1666d_row42_col12" class="data row42 col12" >1</td>
      <td id="T_1666d_row42_col13" class="data row42 col13" >0</td>
      <td id="T_1666d_row42_col14" class="data row42 col14" >2,993</td>
      <td id="T_1666d_row42_col15" class="data row42 col15" >0</td>
      <td id="T_1666d_row42_col16" class="data row42 col16" >0</td>
      <td id="T_1666d_row42_col17" class="data row42 col17" >0</td>
      <td id="T_1666d_row42_col18" class="data row42 col18" >0</td>
      <td id="T_1666d_row42_col19" class="data row42 col19" >0</td>
      <td id="T_1666d_row42_col20" class="data row42 col20" >0</td>
      <td id="T_1666d_row42_col21" class="data row42 col21" >0</td>
      <td id="T_1666d_row42_col22" class="data row42 col22" >62,653</td>
      <td id="T_1666d_row42_col23" class="data row42 col23" >117,348</td>
      <td id="T_1666d_row42_col24" class="data row42 col24" >1</td>
      <td id="T_1666d_row42_col25" class="data row42 col25" >0</td>
      <td id="T_1666d_row42_col26" class="data row42 col26" >0</td>
      <td id="T_1666d_row42_col27" class="data row42 col27" >69,177</td>
      <td id="T_1666d_row42_col28" class="data row42 col28" >168,000</td>
      <td id="T_1666d_row42_col29" class="data row42 col29" >120,695</td>
      <td id="T_1666d_row42_col30" class="data row42 col30" >653,988</td>
    </tr>
    <tr>
      <th id="T_1666d_level1_row43" class="row_heading level1 row43" rowspan="4">Sday_2</th>
      <th id="T_1666d_level2_row43" class="row_heading level2 row43" >B</th>
      <td id="T_1666d_row43_col0" class="data row43 col0" >70,160</td>
      <td id="T_1666d_row43_col1" class="data row43 col1" >0</td>
      <td id="T_1666d_row43_col2" class="data row43 col2" >69,176</td>
      <td id="T_1666d_row43_col3" class="data row43 col3" >7,927</td>
      <td id="T_1666d_row43_col4" class="data row43 col4" >0</td>
      <td id="T_1666d_row43_col5" class="data row43 col5" >0</td>
      <td id="T_1666d_row43_col6" class="data row43 col6" >57,581</td>
      <td id="T_1666d_row43_col7" class="data row43 col7" >25,395</td>
      <td id="T_1666d_row43_col8" class="data row43 col8" >117,600</td>
      <td id="T_1666d_row43_col9" class="data row43 col9" >68,926</td>
      <td id="T_1666d_row43_col10" class="data row43 col10" >502</td>
      <td id="T_1666d_row43_col11" class="data row43 col11" >0</td>
      <td id="T_1666d_row43_col12" class="data row43 col12" >0</td>
      <td id="T_1666d_row43_col13" class="data row43 col13" >0</td>
      <td id="T_1666d_row43_col14" class="data row43 col14" >0</td>
      <td id="T_1666d_row43_col15" class="data row43 col15" >0</td>
      <td id="T_1666d_row43_col16" class="data row43 col16" >69,176</td>
      <td id="T_1666d_row43_col17" class="data row43 col17" >69,176</td>
      <td id="T_1666d_row43_col18" class="data row43 col18" >0</td>
      <td id="T_1666d_row43_col19" class="data row43 col19" >0</td>
      <td id="T_1666d_row43_col20" class="data row43 col20" >117,600</td>
      <td id="T_1666d_row43_col21" class="data row43 col21" >117,600</td>
      <td id="T_1666d_row43_col22" class="data row43 col22" >0</td>
      <td id="T_1666d_row43_col23" class="data row43 col23" >0</td>
      <td id="T_1666d_row43_col24" class="data row43 col24" >0</td>
      <td id="T_1666d_row43_col25" class="data row43 col25" >0</td>
      <td id="T_1666d_row43_col26" class="data row43 col26" >0</td>
      <td id="T_1666d_row43_col27" class="data row43 col27" >0</td>
      <td id="T_1666d_row43_col28" class="data row43 col28" >0</td>
      <td id="T_1666d_row43_col29" class="data row43 col29" >0</td>
      <td id="T_1666d_row43_col30" class="data row43 col30" >790,819</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row44" class="row_heading level2 row44" >D</th>
      <td id="T_1666d_row44_col0" class="data row44 col0" >0</td>
      <td id="T_1666d_row44_col1" class="data row44 col1" >19,475</td>
      <td id="T_1666d_row44_col2" class="data row44 col2" >0</td>
      <td id="T_1666d_row44_col3" class="data row44 col3" >4,566</td>
      <td id="T_1666d_row44_col4" class="data row44 col4" >0</td>
      <td id="T_1666d_row44_col5" class="data row44 col5" >0</td>
      <td id="T_1666d_row44_col6" class="data row44 col6" >1,824</td>
      <td id="T_1666d_row44_col7" class="data row44 col7" >0</td>
      <td id="T_1666d_row44_col8" class="data row44 col8" >0</td>
      <td id="T_1666d_row44_col9" class="data row44 col9" >0</td>
      <td id="T_1666d_row44_col10" class="data row44 col10" >11,248</td>
      <td id="T_1666d_row44_col11" class="data row44 col11" >11,423</td>
      <td id="T_1666d_row44_col12" class="data row44 col12" >0</td>
      <td id="T_1666d_row44_col13" class="data row44 col13" >0</td>
      <td id="T_1666d_row44_col14" class="data row44 col14" >0</td>
      <td id="T_1666d_row44_col15" class="data row44 col15" >72</td>
      <td id="T_1666d_row44_col16" class="data row44 col16" >0</td>
      <td id="T_1666d_row44_col17" class="data row44 col17" >0</td>
      <td id="T_1666d_row44_col18" class="data row44 col18" >0</td>
      <td id="T_1666d_row44_col19" class="data row44 col19" >0</td>
      <td id="T_1666d_row44_col20" class="data row44 col20" >0</td>
      <td id="T_1666d_row44_col21" class="data row44 col21" >0</td>
      <td id="T_1666d_row44_col22" class="data row44 col22" >0</td>
      <td id="T_1666d_row44_col23" class="data row44 col23" >0</td>
      <td id="T_1666d_row44_col24" class="data row44 col24" >25,309</td>
      <td id="T_1666d_row44_col25" class="data row44 col25" >0</td>
      <td id="T_1666d_row44_col26" class="data row44 col26" >0</td>
      <td id="T_1666d_row44_col27" class="data row44 col27" >0</td>
      <td id="T_1666d_row44_col28" class="data row44 col28" >0</td>
      <td id="T_1666d_row44_col29" class="data row44 col29" >0</td>
      <td id="T_1666d_row44_col30" class="data row44 col30" >73,917</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row45" class="row_heading level2 row45" >A</th>
      <td id="T_1666d_row45_col0" class="data row45 col0" >47,440</td>
      <td id="T_1666d_row45_col1" class="data row45 col1" >15,944</td>
      <td id="T_1666d_row45_col2" class="data row45 col2" >0</td>
      <td id="T_1666d_row45_col3" class="data row45 col3" >97,192</td>
      <td id="T_1666d_row45_col4" class="data row45 col4" >0</td>
      <td id="T_1666d_row45_col5" class="data row45 col5" >0</td>
      <td id="T_1666d_row45_col6" class="data row45 col6" >0</td>
      <td id="T_1666d_row45_col7" class="data row45 col7" >0</td>
      <td id="T_1666d_row45_col8" class="data row45 col8" >0</td>
      <td id="T_1666d_row45_col9" class="data row45 col9" >250</td>
      <td id="T_1666d_row45_col10" class="data row45 col10" >0</td>
      <td id="T_1666d_row45_col11" class="data row45 col11" >37,956</td>
      <td id="T_1666d_row45_col12" class="data row45 col12" >69,177</td>
      <td id="T_1666d_row45_col13" class="data row45 col13" >0</td>
      <td id="T_1666d_row45_col14" class="data row45 col14" >0</td>
      <td id="T_1666d_row45_col15" class="data row45 col15" >68,981</td>
      <td id="T_1666d_row45_col16" class="data row45 col16" >0</td>
      <td id="T_1666d_row45_col17" class="data row45 col17" >1</td>
      <td id="T_1666d_row45_col18" class="data row45 col18" >0</td>
      <td id="T_1666d_row45_col19" class="data row45 col19" >0</td>
      <td id="T_1666d_row45_col20" class="data row45 col20" >0</td>
      <td id="T_1666d_row45_col21" class="data row45 col21" >0</td>
      <td id="T_1666d_row45_col22" class="data row45 col22" >50,070</td>
      <td id="T_1666d_row45_col23" class="data row45 col23" >69,177</td>
      <td id="T_1666d_row45_col24" class="data row45 col24" >249</td>
      <td id="T_1666d_row45_col25" class="data row45 col25" >0</td>
      <td id="T_1666d_row45_col26" class="data row45 col26" >0</td>
      <td id="T_1666d_row45_col27" class="data row45 col27" >0</td>
      <td id="T_1666d_row45_col28" class="data row45 col28" >0</td>
      <td id="T_1666d_row45_col29" class="data row45 col29" >0</td>
      <td id="T_1666d_row45_col30" class="data row45 col30" >456,437</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row46" class="row_heading level2 row46" >C</th>
      <td id="T_1666d_row46_col0" class="data row46 col0" >0</td>
      <td id="T_1666d_row46_col1" class="data row46 col1" >3</td>
      <td id="T_1666d_row46_col2" class="data row46 col2" >5,129</td>
      <td id="T_1666d_row46_col3" class="data row46 col3" >0</td>
      <td id="T_1666d_row46_col4" class="data row46 col4" >0</td>
      <td id="T_1666d_row46_col5" class="data row46 col5" >0</td>
      <td id="T_1666d_row46_col6" class="data row46 col6" >12,789</td>
      <td id="T_1666d_row46_col7" class="data row46 col7" >43,782</td>
      <td id="T_1666d_row46_col8" class="data row46 col8" >0</td>
      <td id="T_1666d_row46_col9" class="data row46 col9" >0</td>
      <td id="T_1666d_row46_col10" class="data row46 col10" >37,930</td>
      <td id="T_1666d_row46_col11" class="data row46 col11" >0</td>
      <td id="T_1666d_row46_col12" class="data row46 col12" >249</td>
      <td id="T_1666d_row46_col13" class="data row46 col13" >69,177</td>
      <td id="T_1666d_row46_col14" class="data row46 col14" >69,177</td>
      <td id="T_1666d_row46_col15" class="data row46 col15" >0</td>
      <td id="T_1666d_row46_col16" class="data row46 col16" >0</td>
      <td id="T_1666d_row46_col17" class="data row46 col17" >0</td>
      <td id="T_1666d_row46_col18" class="data row46 col18" >0</td>
      <td id="T_1666d_row46_col19" class="data row46 col19" >0</td>
      <td id="T_1666d_row46_col20" class="data row46 col20" >0</td>
      <td id="T_1666d_row46_col21" class="data row46 col21" >0</td>
      <td id="T_1666d_row46_col22" class="data row46 col22" >19,107</td>
      <td id="T_1666d_row46_col23" class="data row46 col23" >0</td>
      <td id="T_1666d_row46_col24" class="data row46 col24" >249</td>
      <td id="T_1666d_row46_col25" class="data row46 col25" >0</td>
      <td id="T_1666d_row46_col26" class="data row46 col26" >0</td>
      <td id="T_1666d_row46_col27" class="data row46 col27" >69,176</td>
      <td id="T_1666d_row46_col28" class="data row46 col28" >168,000</td>
      <td id="T_1666d_row46_col29" class="data row46 col29" >168,000</td>
      <td id="T_1666d_row46_col30" class="data row46 col30" >662,768</td>
    </tr>
    <tr>
      <th id="T_1666d_level0_row47" class="row_heading level0 row47" rowspan="5">Sew1</th>
      <th id="T_1666d_level1_row47" class="row_heading level1 row47" rowspan="5">Day</th>
      <th id="T_1666d_level2_row47" class="row_heading level2 row47" >H</th>
      <td id="T_1666d_row47_col0" class="data row47 col0" >0</td>
      <td id="T_1666d_row47_col1" class="data row47 col1" >0</td>
      <td id="T_1666d_row47_col2" class="data row47 col2" >0</td>
      <td id="T_1666d_row47_col3" class="data row47 col3" >0</td>
      <td id="T_1666d_row47_col4" class="data row47 col4" >0</td>
      <td id="T_1666d_row47_col5" class="data row47 col5" >0</td>
      <td id="T_1666d_row47_col6" class="data row47 col6" >0</td>
      <td id="T_1666d_row47_col7" class="data row47 col7" >0</td>
      <td id="T_1666d_row47_col8" class="data row47 col8" >0</td>
      <td id="T_1666d_row47_col9" class="data row47 col9" >0</td>
      <td id="T_1666d_row47_col10" class="data row47 col10" >0</td>
      <td id="T_1666d_row47_col11" class="data row47 col11" >0</td>
      <td id="T_1666d_row47_col12" class="data row47 col12" >0</td>
      <td id="T_1666d_row47_col13" class="data row47 col13" >0</td>
      <td id="T_1666d_row47_col14" class="data row47 col14" >2,208</td>
      <td id="T_1666d_row47_col15" class="data row47 col15" >0</td>
      <td id="T_1666d_row47_col16" class="data row47 col16" >0</td>
      <td id="T_1666d_row47_col17" class="data row47 col17" >0</td>
      <td id="T_1666d_row47_col18" class="data row47 col18" >0</td>
      <td id="T_1666d_row47_col19" class="data row47 col19" >0</td>
      <td id="T_1666d_row47_col20" class="data row47 col20" >0</td>
      <td id="T_1666d_row47_col21" class="data row47 col21" >0</td>
      <td id="T_1666d_row47_col22" class="data row47 col22" >0</td>
      <td id="T_1666d_row47_col23" class="data row47 col23" >0</td>
      <td id="T_1666d_row47_col24" class="data row47 col24" >0</td>
      <td id="T_1666d_row47_col25" class="data row47 col25" >0</td>
      <td id="T_1666d_row47_col26" class="data row47 col26" >0</td>
      <td id="T_1666d_row47_col27" class="data row47 col27" >0</td>
      <td id="T_1666d_row47_col28" class="data row47 col28" >0</td>
      <td id="T_1666d_row47_col29" class="data row47 col29" >15,811</td>
      <td id="T_1666d_row47_col30" class="data row47 col30" >18,019</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row48" class="row_heading level2 row48" >D</th>
      <td id="T_1666d_row48_col0" class="data row48 col0" >12,575</td>
      <td id="T_1666d_row48_col1" class="data row48 col1" >21,589</td>
      <td id="T_1666d_row48_col2" class="data row48 col2" >0</td>
      <td id="T_1666d_row48_col3" class="data row48 col3" >0</td>
      <td id="T_1666d_row48_col4" class="data row48 col4" >0</td>
      <td id="T_1666d_row48_col5" class="data row48 col5" >0</td>
      <td id="T_1666d_row48_col6" class="data row48 col6" >0</td>
      <td id="T_1666d_row48_col7" class="data row48 col7" >2,026</td>
      <td id="T_1666d_row48_col8" class="data row48 col8" >20,888</td>
      <td id="T_1666d_row48_col9" class="data row48 col9" >0</td>
      <td id="T_1666d_row48_col10" class="data row48 col10" >0</td>
      <td id="T_1666d_row48_col11" class="data row48 col11" >0</td>
      <td id="T_1666d_row48_col12" class="data row48 col12" >0</td>
      <td id="T_1666d_row48_col13" class="data row48 col13" >0</td>
      <td id="T_1666d_row48_col14" class="data row48 col14" >0</td>
      <td id="T_1666d_row48_col15" class="data row48 col15" >0</td>
      <td id="T_1666d_row48_col16" class="data row48 col16" >0</td>
      <td id="T_1666d_row48_col17" class="data row48 col17" >0</td>
      <td id="T_1666d_row48_col18" class="data row48 col18" >0</td>
      <td id="T_1666d_row48_col19" class="data row48 col19" >0</td>
      <td id="T_1666d_row48_col20" class="data row48 col20" >0</td>
      <td id="T_1666d_row48_col21" class="data row48 col21" >500</td>
      <td id="T_1666d_row48_col22" class="data row48 col22" >0</td>
      <td id="T_1666d_row48_col23" class="data row48 col23" >0</td>
      <td id="T_1666d_row48_col24" class="data row48 col24" >0</td>
      <td id="T_1666d_row48_col25" class="data row48 col25" >0</td>
      <td id="T_1666d_row48_col26" class="data row48 col26" >0</td>
      <td id="T_1666d_row48_col27" class="data row48 col27" >0</td>
      <td id="T_1666d_row48_col28" class="data row48 col28" >40,154</td>
      <td id="T_1666d_row48_col29" class="data row48 col29" >0</td>
      <td id="T_1666d_row48_col30" class="data row48 col30" >97,732</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row49" class="row_heading level2 row49" >G</th>
      <td id="T_1666d_row49_col0" class="data row49 col0" >41,117</td>
      <td id="T_1666d_row49_col1" class="data row49 col1" >0</td>
      <td id="T_1666d_row49_col2" class="data row49 col2" >0</td>
      <td id="T_1666d_row49_col3" class="data row49 col3" >0</td>
      <td id="T_1666d_row49_col4" class="data row49 col4" >0</td>
      <td id="T_1666d_row49_col5" class="data row49 col5" >0</td>
      <td id="T_1666d_row49_col6" class="data row49 col6" >0</td>
      <td id="T_1666d_row49_col7" class="data row49 col7" >0</td>
      <td id="T_1666d_row49_col8" class="data row49 col8" >4,245</td>
      <td id="T_1666d_row49_col9" class="data row49 col9" >0</td>
      <td id="T_1666d_row49_col10" class="data row49 col10" >0</td>
      <td id="T_1666d_row49_col11" class="data row49 col11" >0</td>
      <td id="T_1666d_row49_col12" class="data row49 col12" >0</td>
      <td id="T_1666d_row49_col13" class="data row49 col13" >0</td>
      <td id="T_1666d_row49_col14" class="data row49 col14" >0</td>
      <td id="T_1666d_row49_col15" class="data row49 col15" >0</td>
      <td id="T_1666d_row49_col16" class="data row49 col16" >0</td>
      <td id="T_1666d_row49_col17" class="data row49 col17" >0</td>
      <td id="T_1666d_row49_col18" class="data row49 col18" >0</td>
      <td id="T_1666d_row49_col19" class="data row49 col19" >0</td>
      <td id="T_1666d_row49_col20" class="data row49 col20" >0</td>
      <td id="T_1666d_row49_col21" class="data row49 col21" >0</td>
      <td id="T_1666d_row49_col22" class="data row49 col22" >0</td>
      <td id="T_1666d_row49_col23" class="data row49 col23" >0</td>
      <td id="T_1666d_row49_col24" class="data row49 col24" >0</td>
      <td id="T_1666d_row49_col25" class="data row49 col25" >0</td>
      <td id="T_1666d_row49_col26" class="data row49 col26" >0</td>
      <td id="T_1666d_row49_col27" class="data row49 col27" >0</td>
      <td id="T_1666d_row49_col28" class="data row49 col28" >0</td>
      <td id="T_1666d_row49_col29" class="data row49 col29" >0</td>
      <td id="T_1666d_row49_col30" class="data row49 col30" >45,362</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row50" class="row_heading level2 row50" >I</th>
      <td id="T_1666d_row50_col0" class="data row50 col0" >1,001</td>
      <td id="T_1666d_row50_col1" class="data row50 col1" >108,033</td>
      <td id="T_1666d_row50_col2" class="data row50 col2" >172,300</td>
      <td id="T_1666d_row50_col3" class="data row50 col3" >172,800</td>
      <td id="T_1666d_row50_col4" class="data row50 col4" >0</td>
      <td id="T_1666d_row50_col5" class="data row50 col5" >0</td>
      <td id="T_1666d_row50_col6" class="data row50 col6" >172,800</td>
      <td id="T_1666d_row50_col7" class="data row50 col7" >65,076</td>
      <td id="T_1666d_row50_col8" class="data row50 col8" >0</td>
      <td id="T_1666d_row50_col9" class="data row50 col9" >71,153</td>
      <td id="T_1666d_row50_col10" class="data row50 col10" >172,800</td>
      <td id="T_1666d_row50_col11" class="data row50 col11" >172,800</td>
      <td id="T_1666d_row50_col12" class="data row50 col12" >172,800</td>
      <td id="T_1666d_row50_col13" class="data row50 col13" >80,943</td>
      <td id="T_1666d_row50_col14" class="data row50 col14" >116,044</td>
      <td id="T_1666d_row50_col15" class="data row50 col15" >166,121</td>
      <td id="T_1666d_row50_col16" class="data row50 col16" >172,799</td>
      <td id="T_1666d_row50_col17" class="data row50 col17" >171,300</td>
      <td id="T_1666d_row50_col18" class="data row50 col18" >0</td>
      <td id="T_1666d_row50_col19" class="data row50 col19" >0</td>
      <td id="T_1666d_row50_col20" class="data row50 col20" >172,800</td>
      <td id="T_1666d_row50_col21" class="data row50 col21" >171,300</td>
      <td id="T_1666d_row50_col22" class="data row50 col22" >71,153</td>
      <td id="T_1666d_row50_col23" class="data row50 col23" >172,800</td>
      <td id="T_1666d_row50_col24" class="data row50 col24" >0</td>
      <td id="T_1666d_row50_col25" class="data row50 col25" >0</td>
      <td id="T_1666d_row50_col26" class="data row50 col26" >0</td>
      <td id="T_1666d_row50_col27" class="data row50 col27" >71,153</td>
      <td id="T_1666d_row50_col28" class="data row50 col28" >500</td>
      <td id="T_1666d_row50_col29" class="data row50 col29" >39,532</td>
      <td id="T_1666d_row50_col30" class="data row50 col30" >2,688,008</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row51" class="row_heading level2 row51" >F</th>
      <td id="T_1666d_row51_col0" class="data row51 col0" >0</td>
      <td id="T_1666d_row51_col1" class="data row51 col1" >0</td>
      <td id="T_1666d_row51_col2" class="data row51 col2" >500</td>
      <td id="T_1666d_row51_col3" class="data row51 col3" >0</td>
      <td id="T_1666d_row51_col4" class="data row51 col4" >0</td>
      <td id="T_1666d_row51_col5" class="data row51 col5" >0</td>
      <td id="T_1666d_row51_col6" class="data row51 col6" >0</td>
      <td id="T_1666d_row51_col7" class="data row51 col7" >0</td>
      <td id="T_1666d_row51_col8" class="data row51 col8" >0</td>
      <td id="T_1666d_row51_col9" class="data row51 col9" >0</td>
      <td id="T_1666d_row51_col10" class="data row51 col10" >0</td>
      <td id="T_1666d_row51_col11" class="data row51 col11" >0</td>
      <td id="T_1666d_row51_col12" class="data row51 col12" >0</td>
      <td id="T_1666d_row51_col13" class="data row51 col13" >27,874</td>
      <td id="T_1666d_row51_col14" class="data row51 col14" >500</td>
      <td id="T_1666d_row51_col15" class="data row51 col15" >500</td>
      <td id="T_1666d_row51_col16" class="data row51 col16" >0</td>
      <td id="T_1666d_row51_col17" class="data row51 col17" >0</td>
      <td id="T_1666d_row51_col18" class="data row51 col18" >0</td>
      <td id="T_1666d_row51_col19" class="data row51 col19" >0</td>
      <td id="T_1666d_row51_col20" class="data row51 col20" >0</td>
      <td id="T_1666d_row51_col21" class="data row51 col21" >0</td>
      <td id="T_1666d_row51_col22" class="data row51 col22" >0</td>
      <td id="T_1666d_row51_col23" class="data row51 col23" >0</td>
      <td id="T_1666d_row51_col24" class="data row51 col24" >172,800</td>
      <td id="T_1666d_row51_col25" class="data row51 col25" >0</td>
      <td id="T_1666d_row51_col26" class="data row51 col26" >0</td>
      <td id="T_1666d_row51_col27" class="data row51 col27" >0</td>
      <td id="T_1666d_row51_col28" class="data row51 col28" >0</td>
      <td id="T_1666d_row51_col29" class="data row51 col29" >0</td>
      <td id="T_1666d_row51_col30" class="data row51 col30" >202,174</td>
    </tr>
    <tr>
      <th id="T_1666d_level0_row52" class="row_heading level0 row52" rowspan="6">Sew2</th>
      <th id="T_1666d_level1_row52" class="row_heading level1 row52" rowspan="6">Day</th>
      <th id="T_1666d_level2_row52" class="row_heading level2 row52" >H</th>
      <td id="T_1666d_row52_col0" class="data row52 col0" >0</td>
      <td id="T_1666d_row52_col1" class="data row52 col1" >0</td>
      <td id="T_1666d_row52_col2" class="data row52 col2" >0</td>
      <td id="T_1666d_row52_col3" class="data row52 col3" >0</td>
      <td id="T_1666d_row52_col4" class="data row52 col4" >0</td>
      <td id="T_1666d_row52_col5" class="data row52 col5" >0</td>
      <td id="T_1666d_row52_col6" class="data row52 col6" >0</td>
      <td id="T_1666d_row52_col7" class="data row52 col7" >0</td>
      <td id="T_1666d_row52_col8" class="data row52 col8" >0</td>
      <td id="T_1666d_row52_col9" class="data row52 col9" >0</td>
      <td id="T_1666d_row52_col10" class="data row52 col10" >0</td>
      <td id="T_1666d_row52_col11" class="data row52 col11" >0</td>
      <td id="T_1666d_row52_col12" class="data row52 col12" >0</td>
      <td id="T_1666d_row52_col13" class="data row52 col13" >0</td>
      <td id="T_1666d_row52_col14" class="data row52 col14" >0</td>
      <td id="T_1666d_row52_col15" class="data row52 col15" >0</td>
      <td id="T_1666d_row52_col16" class="data row52 col16" >0</td>
      <td id="T_1666d_row52_col17" class="data row52 col17" >0</td>
      <td id="T_1666d_row52_col18" class="data row52 col18" >0</td>
      <td id="T_1666d_row52_col19" class="data row52 col19" >0</td>
      <td id="T_1666d_row52_col20" class="data row52 col20" >0</td>
      <td id="T_1666d_row52_col21" class="data row52 col21" >0</td>
      <td id="T_1666d_row52_col22" class="data row52 col22" >0</td>
      <td id="T_1666d_row52_col23" class="data row52 col23" >789</td>
      <td id="T_1666d_row52_col24" class="data row52 col24" >0</td>
      <td id="T_1666d_row52_col25" class="data row52 col25" >0</td>
      <td id="T_1666d_row52_col26" class="data row52 col26" >0</td>
      <td id="T_1666d_row52_col27" class="data row52 col27" >0</td>
      <td id="T_1666d_row52_col28" class="data row52 col28" >0</td>
      <td id="T_1666d_row52_col29" class="data row52 col29" >0</td>
      <td id="T_1666d_row52_col30" class="data row52 col30" >789</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row53" class="row_heading level2 row53" >D</th>
      <td id="T_1666d_row53_col0" class="data row53 col0" >0</td>
      <td id="T_1666d_row53_col1" class="data row53 col1" >0</td>
      <td id="T_1666d_row53_col2" class="data row53 col2" >0</td>
      <td id="T_1666d_row53_col3" class="data row53 col3" >0</td>
      <td id="T_1666d_row53_col4" class="data row53 col4" >0</td>
      <td id="T_1666d_row53_col5" class="data row53 col5" >0</td>
      <td id="T_1666d_row53_col6" class="data row53 col6" >0</td>
      <td id="T_1666d_row53_col7" class="data row53 col7" >0</td>
      <td id="T_1666d_row53_col8" class="data row53 col8" >0</td>
      <td id="T_1666d_row53_col9" class="data row53 col9" >0</td>
      <td id="T_1666d_row53_col10" class="data row53 col10" >0</td>
      <td id="T_1666d_row53_col11" class="data row53 col11" >0</td>
      <td id="T_1666d_row53_col12" class="data row53 col12" >0</td>
      <td id="T_1666d_row53_col13" class="data row53 col13" >0</td>
      <td id="T_1666d_row53_col14" class="data row53 col14" >0</td>
      <td id="T_1666d_row53_col15" class="data row53 col15" >13,945</td>
      <td id="T_1666d_row53_col16" class="data row53 col16" >0</td>
      <td id="T_1666d_row53_col17" class="data row53 col17" >0</td>
      <td id="T_1666d_row53_col18" class="data row53 col18" >0</td>
      <td id="T_1666d_row53_col19" class="data row53 col19" >0</td>
      <td id="T_1666d_row53_col20" class="data row53 col20" >0</td>
      <td id="T_1666d_row53_col21" class="data row53 col21" >6,372</td>
      <td id="T_1666d_row53_col22" class="data row53 col22" >23,718</td>
      <td id="T_1666d_row53_col23" class="data row53 col23" >0</td>
      <td id="T_1666d_row53_col24" class="data row53 col24" >0</td>
      <td id="T_1666d_row53_col25" class="data row53 col25" >0</td>
      <td id="T_1666d_row53_col26" class="data row53 col26" >0</td>
      <td id="T_1666d_row53_col27" class="data row53 col27" >0</td>
      <td id="T_1666d_row53_col28" class="data row53 col28" >23,222</td>
      <td id="T_1666d_row53_col29" class="data row53 col29" >0</td>
      <td id="T_1666d_row53_col30" class="data row53 col30" >67,257</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row54" class="row_heading level2 row54" >G</th>
      <td id="T_1666d_row54_col0" class="data row54 col0" >0</td>
      <td id="T_1666d_row54_col1" class="data row54 col1" >0</td>
      <td id="T_1666d_row54_col2" class="data row54 col2" >0</td>
      <td id="T_1666d_row54_col3" class="data row54 col3" >0</td>
      <td id="T_1666d_row54_col4" class="data row54 col4" >0</td>
      <td id="T_1666d_row54_col5" class="data row54 col5" >0</td>
      <td id="T_1666d_row54_col6" class="data row54 col6" >0</td>
      <td id="T_1666d_row54_col7" class="data row54 col7" >0</td>
      <td id="T_1666d_row54_col8" class="data row54 col8" >0</td>
      <td id="T_1666d_row54_col9" class="data row54 col9" >0</td>
      <td id="T_1666d_row54_col10" class="data row54 col10" >0</td>
      <td id="T_1666d_row54_col11" class="data row54 col11" >0</td>
      <td id="T_1666d_row54_col12" class="data row54 col12" >0</td>
      <td id="T_1666d_row54_col13" class="data row54 col13" >0</td>
      <td id="T_1666d_row54_col14" class="data row54 col14" >830</td>
      <td id="T_1666d_row54_col15" class="data row54 col15" >0</td>
      <td id="T_1666d_row54_col16" class="data row54 col16" >0</td>
      <td id="T_1666d_row54_col17" class="data row54 col17" >0</td>
      <td id="T_1666d_row54_col18" class="data row54 col18" >0</td>
      <td id="T_1666d_row54_col19" class="data row54 col19" >0</td>
      <td id="T_1666d_row54_col20" class="data row54 col20" >0</td>
      <td id="T_1666d_row54_col21" class="data row54 col21" >0</td>
      <td id="T_1666d_row54_col22" class="data row54 col22" >0</td>
      <td id="T_1666d_row54_col23" class="data row54 col23" >0</td>
      <td id="T_1666d_row54_col24" class="data row54 col24" >0</td>
      <td id="T_1666d_row54_col25" class="data row54 col25" >0</td>
      <td id="T_1666d_row54_col26" class="data row54 col26" >0</td>
      <td id="T_1666d_row54_col27" class="data row54 col27" >0</td>
      <td id="T_1666d_row54_col28" class="data row54 col28" >0</td>
      <td id="T_1666d_row54_col29" class="data row54 col29" >0</td>
      <td id="T_1666d_row54_col30" class="data row54 col30" >830</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row55" class="row_heading level2 row55" >I</th>
      <td id="T_1666d_row55_col0" class="data row55 col0" >122,797</td>
      <td id="T_1666d_row55_col1" class="data row55 col1" >172,800</td>
      <td id="T_1666d_row55_col2" class="data row55 col2" >155,080</td>
      <td id="T_1666d_row55_col3" class="data row55 col3" >172,800</td>
      <td id="T_1666d_row55_col4" class="data row55 col4" >0</td>
      <td id="T_1666d_row55_col5" class="data row55 col5" >0</td>
      <td id="T_1666d_row55_col6" class="data row55 col6" >172,798</td>
      <td id="T_1666d_row55_col7" class="data row55 col7" >71,153</td>
      <td id="T_1666d_row55_col8" class="data row55 col8" >71,153</td>
      <td id="T_1666d_row55_col9" class="data row55 col9" >149,983</td>
      <td id="T_1666d_row55_col10" class="data row55 col10" >172,800</td>
      <td id="T_1666d_row55_col11" class="data row55 col11" >172,800</td>
      <td id="T_1666d_row55_col12" class="data row55 col12" >172,800</td>
      <td id="T_1666d_row55_col13" class="data row55 col13" >172,800</td>
      <td id="T_1666d_row55_col14" class="data row55 col14" >120,960</td>
      <td id="T_1666d_row55_col15" class="data row55 col15" >0</td>
      <td id="T_1666d_row55_col16" class="data row55 col16" >172,800</td>
      <td id="T_1666d_row55_col17" class="data row55 col17" >8,900</td>
      <td id="T_1666d_row55_col18" class="data row55 col18" >0</td>
      <td id="T_1666d_row55_col19" class="data row55 col19" >0</td>
      <td id="T_1666d_row55_col20" class="data row55 col20" >120,960</td>
      <td id="T_1666d_row55_col21" class="data row55 col21" >0</td>
      <td id="T_1666d_row55_col22" class="data row55 col22" >0</td>
      <td id="T_1666d_row55_col23" class="data row55 col23" >171,222</td>
      <td id="T_1666d_row55_col24" class="data row55 col24" >2,034</td>
      <td id="T_1666d_row55_col25" class="data row55 col25" >0</td>
      <td id="T_1666d_row55_col26" class="data row55 col26" >0</td>
      <td id="T_1666d_row55_col27" class="data row55 col27" >71,152</td>
      <td id="T_1666d_row55_col28" class="data row55 col28" >0</td>
      <td id="T_1666d_row55_col29" class="data row55 col29" >71,153</td>
      <td id="T_1666d_row55_col30" class="data row55 col30" >2,518,945</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row56" class="row_heading level2 row56" >E</th>
      <td id="T_1666d_row56_col0" class="data row56 col0" >0</td>
      <td id="T_1666d_row56_col1" class="data row56 col1" >0</td>
      <td id="T_1666d_row56_col2" class="data row56 col2" >0</td>
      <td id="T_1666d_row56_col3" class="data row56 col3" >0</td>
      <td id="T_1666d_row56_col4" class="data row56 col4" >0</td>
      <td id="T_1666d_row56_col5" class="data row56 col5" >0</td>
      <td id="T_1666d_row56_col6" class="data row56 col6" >0</td>
      <td id="T_1666d_row56_col7" class="data row56 col7" >0</td>
      <td id="T_1666d_row56_col8" class="data row56 col8" >0</td>
      <td id="T_1666d_row56_col9" class="data row56 col9" >0</td>
      <td id="T_1666d_row56_col10" class="data row56 col10" >0</td>
      <td id="T_1666d_row56_col11" class="data row56 col11" >0</td>
      <td id="T_1666d_row56_col12" class="data row56 col12" >0</td>
      <td id="T_1666d_row56_col13" class="data row56 col13" >0</td>
      <td id="T_1666d_row56_col14" class="data row56 col14" >0</td>
      <td id="T_1666d_row56_col15" class="data row56 col15" >1,018</td>
      <td id="T_1666d_row56_col16" class="data row56 col16" >0</td>
      <td id="T_1666d_row56_col17" class="data row56 col17" >0</td>
      <td id="T_1666d_row56_col18" class="data row56 col18" >0</td>
      <td id="T_1666d_row56_col19" class="data row56 col19" >0</td>
      <td id="T_1666d_row56_col20" class="data row56 col20" >0</td>
      <td id="T_1666d_row56_col21" class="data row56 col21" >0</td>
      <td id="T_1666d_row56_col22" class="data row56 col22" >0</td>
      <td id="T_1666d_row56_col23" class="data row56 col23" >0</td>
      <td id="T_1666d_row56_col24" class="data row56 col24" >0</td>
      <td id="T_1666d_row56_col25" class="data row56 col25" >0</td>
      <td id="T_1666d_row56_col26" class="data row56 col26" >0</td>
      <td id="T_1666d_row56_col27" class="data row56 col27" >0</td>
      <td id="T_1666d_row56_col28" class="data row56 col28" >0</td>
      <td id="T_1666d_row56_col29" class="data row56 col29" >0</td>
      <td id="T_1666d_row56_col30" class="data row56 col30" >1,018</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row57" class="row_heading level2 row57" >F</th>
      <td id="T_1666d_row57_col0" class="data row57 col0" >0</td>
      <td id="T_1666d_row57_col1" class="data row57 col1" >0</td>
      <td id="T_1666d_row57_col2" class="data row57 col2" >0</td>
      <td id="T_1666d_row57_col3" class="data row57 col3" >0</td>
      <td id="T_1666d_row57_col4" class="data row57 col4" >0</td>
      <td id="T_1666d_row57_col5" class="data row57 col5" >0</td>
      <td id="T_1666d_row57_col6" class="data row57 col6" >0</td>
      <td id="T_1666d_row57_col7" class="data row57 col7" >0</td>
      <td id="T_1666d_row57_col8" class="data row57 col8" >0</td>
      <td id="T_1666d_row57_col9" class="data row57 col9" >0</td>
      <td id="T_1666d_row57_col10" class="data row57 col10" >0</td>
      <td id="T_1666d_row57_col11" class="data row57 col11" >0</td>
      <td id="T_1666d_row57_col12" class="data row57 col12" >0</td>
      <td id="T_1666d_row57_col13" class="data row57 col13" >0</td>
      <td id="T_1666d_row57_col14" class="data row57 col14" >0</td>
      <td id="T_1666d_row57_col15" class="data row57 col15" >27,689</td>
      <td id="T_1666d_row57_col16" class="data row57 col16" >0</td>
      <td id="T_1666d_row57_col17" class="data row57 col17" >62,253</td>
      <td id="T_1666d_row57_col18" class="data row57 col18" >0</td>
      <td id="T_1666d_row57_col19" class="data row57 col19" >0</td>
      <td id="T_1666d_row57_col20" class="data row57 col20" >0</td>
      <td id="T_1666d_row57_col21" class="data row57 col21" >101,844</td>
      <td id="T_1666d_row57_col22" class="data row57 col22" >0</td>
      <td id="T_1666d_row57_col23" class="data row57 col23" >0</td>
      <td id="T_1666d_row57_col24" class="data row57 col24" >170,766</td>
      <td id="T_1666d_row57_col25" class="data row57 col25" >0</td>
      <td id="T_1666d_row57_col26" class="data row57 col26" >0</td>
      <td id="T_1666d_row57_col27" class="data row57 col27" >0</td>
      <td id="T_1666d_row57_col28" class="data row57 col28" >51,294</td>
      <td id="T_1666d_row57_col29" class="data row57 col29" >0</td>
      <td id="T_1666d_row57_col30" class="data row57 col30" >413,846</td>
    </tr>
    <tr>
      <th id="T_1666d_level0_row58" class="row_heading level0 row58" rowspan="6">Sew3</th>
      <th id="T_1666d_level1_row58" class="row_heading level1 row58" rowspan="6">Day</th>
      <th id="T_1666d_level2_row58" class="row_heading level2 row58" >H</th>
      <td id="T_1666d_row58_col0" class="data row58 col0" >0</td>
      <td id="T_1666d_row58_col1" class="data row58 col1" >0</td>
      <td id="T_1666d_row58_col2" class="data row58 col2" >0</td>
      <td id="T_1666d_row58_col3" class="data row58 col3" >0</td>
      <td id="T_1666d_row58_col4" class="data row58 col4" >0</td>
      <td id="T_1666d_row58_col5" class="data row58 col5" >0</td>
      <td id="T_1666d_row58_col6" class="data row58 col6" >0</td>
      <td id="T_1666d_row58_col7" class="data row58 col7" >0</td>
      <td id="T_1666d_row58_col8" class="data row58 col8" >0</td>
      <td id="T_1666d_row58_col9" class="data row58 col9" >0</td>
      <td id="T_1666d_row58_col10" class="data row58 col10" >0</td>
      <td id="T_1666d_row58_col11" class="data row58 col11" >0</td>
      <td id="T_1666d_row58_col12" class="data row58 col12" >0</td>
      <td id="T_1666d_row58_col13" class="data row58 col13" >0</td>
      <td id="T_1666d_row58_col14" class="data row58 col14" >0</td>
      <td id="T_1666d_row58_col15" class="data row58 col15" >0</td>
      <td id="T_1666d_row58_col16" class="data row58 col16" >0</td>
      <td id="T_1666d_row58_col17" class="data row58 col17" >0</td>
      <td id="T_1666d_row58_col18" class="data row58 col18" >0</td>
      <td id="T_1666d_row58_col19" class="data row58 col19" >0</td>
      <td id="T_1666d_row58_col20" class="data row58 col20" >0</td>
      <td id="T_1666d_row58_col21" class="data row58 col21" >0</td>
      <td id="T_1666d_row58_col22" class="data row58 col22" >3,324</td>
      <td id="T_1666d_row58_col23" class="data row58 col23" >0</td>
      <td id="T_1666d_row58_col24" class="data row58 col24" >0</td>
      <td id="T_1666d_row58_col25" class="data row58 col25" >0</td>
      <td id="T_1666d_row58_col26" class="data row58 col26" >0</td>
      <td id="T_1666d_row58_col27" class="data row58 col27" >0</td>
      <td id="T_1666d_row58_col28" class="data row58 col28" >0</td>
      <td id="T_1666d_row58_col29" class="data row58 col29" >0</td>
      <td id="T_1666d_row58_col30" class="data row58 col30" >3,324</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row59" class="row_heading level2 row59" >D</th>
      <td id="T_1666d_row59_col0" class="data row59 col0" >40,054</td>
      <td id="T_1666d_row59_col1" class="data row59 col1" >0</td>
      <td id="T_1666d_row59_col2" class="data row59 col2" >0</td>
      <td id="T_1666d_row59_col3" class="data row59 col3" >0</td>
      <td id="T_1666d_row59_col4" class="data row59 col4" >0</td>
      <td id="T_1666d_row59_col5" class="data row59 col5" >0</td>
      <td id="T_1666d_row59_col6" class="data row59 col6" >0</td>
      <td id="T_1666d_row59_col7" class="data row59 col7" >0</td>
      <td id="T_1666d_row59_col8" class="data row59 col8" >0</td>
      <td id="T_1666d_row59_col9" class="data row59 col9" >0</td>
      <td id="T_1666d_row59_col10" class="data row59 col10" >0</td>
      <td id="T_1666d_row59_col11" class="data row59 col11" >0</td>
      <td id="T_1666d_row59_col12" class="data row59 col12" >0</td>
      <td id="T_1666d_row59_col13" class="data row59 col13" >0</td>
      <td id="T_1666d_row59_col14" class="data row59 col14" >0</td>
      <td id="T_1666d_row59_col15" class="data row59 col15" >16,345</td>
      <td id="T_1666d_row59_col16" class="data row59 col16" >0</td>
      <td id="T_1666d_row59_col17" class="data row59 col17" >0</td>
      <td id="T_1666d_row59_col18" class="data row59 col18" >0</td>
      <td id="T_1666d_row59_col19" class="data row59 col19" >0</td>
      <td id="T_1666d_row59_col20" class="data row59 col20" >0</td>
      <td id="T_1666d_row59_col21" class="data row59 col21" >7,422</td>
      <td id="T_1666d_row59_col22" class="data row59 col22" >10,748</td>
      <td id="T_1666d_row59_col23" class="data row59 col23" >22,272</td>
      <td id="T_1666d_row59_col24" class="data row59 col24" >50,616</td>
      <td id="T_1666d_row59_col25" class="data row59 col25" >0</td>
      <td id="T_1666d_row59_col26" class="data row59 col26" >0</td>
      <td id="T_1666d_row59_col27" class="data row59 col27" >500</td>
      <td id="T_1666d_row59_col28" class="data row59 col28" >0</td>
      <td id="T_1666d_row59_col29" class="data row59 col29" >13,170</td>
      <td id="T_1666d_row59_col30" class="data row59 col30" >161,127</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row60" class="row_heading level2 row60" >G</th>
      <td id="T_1666d_row60_col0" class="data row60 col0" >0</td>
      <td id="T_1666d_row60_col1" class="data row60 col1" >24,171</td>
      <td id="T_1666d_row60_col2" class="data row60 col2" >0</td>
      <td id="T_1666d_row60_col3" class="data row60 col3" >0</td>
      <td id="T_1666d_row60_col4" class="data row60 col4" >0</td>
      <td id="T_1666d_row60_col5" class="data row60 col5" >0</td>
      <td id="T_1666d_row60_col6" class="data row60 col6" >0</td>
      <td id="T_1666d_row60_col7" class="data row60 col7" >0</td>
      <td id="T_1666d_row60_col8" class="data row60 col8" >0</td>
      <td id="T_1666d_row60_col9" class="data row60 col9" >0</td>
      <td id="T_1666d_row60_col10" class="data row60 col10" >0</td>
      <td id="T_1666d_row60_col11" class="data row60 col11" >15,389</td>
      <td id="T_1666d_row60_col12" class="data row60 col12" >0</td>
      <td id="T_1666d_row60_col13" class="data row60 col13" >0</td>
      <td id="T_1666d_row60_col14" class="data row60 col14" >0</td>
      <td id="T_1666d_row60_col15" class="data row60 col15" >0</td>
      <td id="T_1666d_row60_col16" class="data row60 col16" >0</td>
      <td id="T_1666d_row60_col17" class="data row60 col17" >0</td>
      <td id="T_1666d_row60_col18" class="data row60 col18" >0</td>
      <td id="T_1666d_row60_col19" class="data row60 col19" >0</td>
      <td id="T_1666d_row60_col20" class="data row60 col20" >0</td>
      <td id="T_1666d_row60_col21" class="data row60 col21" >70,432</td>
      <td id="T_1666d_row60_col22" class="data row60 col22" >0</td>
      <td id="T_1666d_row60_col23" class="data row60 col23" >0</td>
      <td id="T_1666d_row60_col24" class="data row60 col24" >0</td>
      <td id="T_1666d_row60_col25" class="data row60 col25" >0</td>
      <td id="T_1666d_row60_col26" class="data row60 col26" >0</td>
      <td id="T_1666d_row60_col27" class="data row60 col27" >0</td>
      <td id="T_1666d_row60_col28" class="data row60 col28" >0</td>
      <td id="T_1666d_row60_col29" class="data row60 col29" >15,822</td>
      <td id="T_1666d_row60_col30" class="data row60 col30" >125,814</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row61" class="row_heading level2 row61" >I</th>
      <td id="T_1666d_row61_col0" class="data row61 col0" >0</td>
      <td id="T_1666d_row61_col1" class="data row61 col1" >46,677</td>
      <td id="T_1666d_row61_col2" class="data row61 col2" >29,684</td>
      <td id="T_1666d_row61_col3" class="data row61 col3" >172,800</td>
      <td id="T_1666d_row61_col4" class="data row61 col4" >0</td>
      <td id="T_1666d_row61_col5" class="data row61 col5" >0</td>
      <td id="T_1666d_row61_col6" class="data row61 col6" >167,946</td>
      <td id="T_1666d_row61_col7" class="data row61 col7" >71,153</td>
      <td id="T_1666d_row61_col8" class="data row61 col8" >120,960</td>
      <td id="T_1666d_row61_col9" class="data row61 col9" >118,883</td>
      <td id="T_1666d_row61_col10" class="data row61 col10" >172,800</td>
      <td id="T_1666d_row61_col11" class="data row61 col11" >142,021</td>
      <td id="T_1666d_row61_col12" class="data row61 col12" >111,108</td>
      <td id="T_1666d_row61_col13" class="data row61 col13" >172,800</td>
      <td id="T_1666d_row61_col14" class="data row61 col14" >139,689</td>
      <td id="T_1666d_row61_col15" class="data row61 col15" >22,120</td>
      <td id="T_1666d_row61_col16" class="data row61 col16" >120,960</td>
      <td id="T_1666d_row61_col17" class="data row61 col17" >0</td>
      <td id="T_1666d_row61_col18" class="data row61 col18" >0</td>
      <td id="T_1666d_row61_col19" class="data row61 col19" >0</td>
      <td id="T_1666d_row61_col20" class="data row61 col20" >172,800</td>
      <td id="T_1666d_row61_col21" class="data row61 col21" >500</td>
      <td id="T_1666d_row61_col22" class="data row61 col22" >0</td>
      <td id="T_1666d_row61_col23" class="data row61 col23" >54,142</td>
      <td id="T_1666d_row61_col24" class="data row61 col24" >0</td>
      <td id="T_1666d_row61_col25" class="data row61 col25" >0</td>
      <td id="T_1666d_row61_col26" class="data row61 col26" >0</td>
      <td id="T_1666d_row61_col27" class="data row61 col27" >71,152</td>
      <td id="T_1666d_row61_col28" class="data row61 col28" >500</td>
      <td id="T_1666d_row61_col29" class="data row61 col29" >0</td>
      <td id="T_1666d_row61_col30" class="data row61 col30" >1,908,695</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row62" class="row_heading level2 row62" >E</th>
      <td id="T_1666d_row62_col0" class="data row62 col0" >500</td>
      <td id="T_1666d_row62_col1" class="data row62 col1" >0</td>
      <td id="T_1666d_row62_col2" class="data row62 col2" >0</td>
      <td id="T_1666d_row62_col3" class="data row62 col3" >0</td>
      <td id="T_1666d_row62_col4" class="data row62 col4" >0</td>
      <td id="T_1666d_row62_col5" class="data row62 col5" >0</td>
      <td id="T_1666d_row62_col6" class="data row62 col6" >0</td>
      <td id="T_1666d_row62_col7" class="data row62 col7" >0</td>
      <td id="T_1666d_row62_col8" class="data row62 col8" >0</td>
      <td id="T_1666d_row62_col9" class="data row62 col9" >0</td>
      <td id="T_1666d_row62_col10" class="data row62 col10" >0</td>
      <td id="T_1666d_row62_col11" class="data row62 col11" >0</td>
      <td id="T_1666d_row62_col12" class="data row62 col12" >0</td>
      <td id="T_1666d_row62_col13" class="data row62 col13" >0</td>
      <td id="T_1666d_row62_col14" class="data row62 col14" >0</td>
      <td id="T_1666d_row62_col15" class="data row62 col15" >0</td>
      <td id="T_1666d_row62_col16" class="data row62 col16" >0</td>
      <td id="T_1666d_row62_col17" class="data row62 col17" >500</td>
      <td id="T_1666d_row62_col18" class="data row62 col18" >0</td>
      <td id="T_1666d_row62_col19" class="data row62 col19" >0</td>
      <td id="T_1666d_row62_col20" class="data row62 col20" >0</td>
      <td id="T_1666d_row62_col21" class="data row62 col21" >0</td>
      <td id="T_1666d_row62_col22" class="data row62 col22" >0</td>
      <td id="T_1666d_row62_col23" class="data row62 col23" >0</td>
      <td id="T_1666d_row62_col24" class="data row62 col24" >0</td>
      <td id="T_1666d_row62_col25" class="data row62 col25" >0</td>
      <td id="T_1666d_row62_col26" class="data row62 col26" >0</td>
      <td id="T_1666d_row62_col27" class="data row62 col27" >0</td>
      <td id="T_1666d_row62_col28" class="data row62 col28" >0</td>
      <td id="T_1666d_row62_col29" class="data row62 col29" >0</td>
      <td id="T_1666d_row62_col30" class="data row62 col30" >1,000</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row63" class="row_heading level2 row63" >F</th>
      <td id="T_1666d_row63_col0" class="data row63 col0" >0</td>
      <td id="T_1666d_row63_col1" class="data row63 col1" >77,781</td>
      <td id="T_1666d_row63_col2" class="data row63 col2" >91,276</td>
      <td id="T_1666d_row63_col3" class="data row63 col3" >0</td>
      <td id="T_1666d_row63_col4" class="data row63 col4" >0</td>
      <td id="T_1666d_row63_col5" class="data row63 col5" >0</td>
      <td id="T_1666d_row63_col6" class="data row63 col6" >0</td>
      <td id="T_1666d_row63_col7" class="data row63 col7" >0</td>
      <td id="T_1666d_row63_col8" class="data row63 col8" >0</td>
      <td id="T_1666d_row63_col9" class="data row63 col9" >2,743</td>
      <td id="T_1666d_row63_col10" class="data row63 col10" >0</td>
      <td id="T_1666d_row63_col11" class="data row63 col11" >0</td>
      <td id="T_1666d_row63_col12" class="data row63 col12" >9,852</td>
      <td id="T_1666d_row63_col13" class="data row63 col13" >0</td>
      <td id="T_1666d_row63_col14" class="data row63 col14" >500</td>
      <td id="T_1666d_row63_col15" class="data row63 col15" >0</td>
      <td id="T_1666d_row63_col16" class="data row63 col16" >0</td>
      <td id="T_1666d_row63_col17" class="data row63 col17" >70,353</td>
      <td id="T_1666d_row63_col18" class="data row63 col18" >0</td>
      <td id="T_1666d_row63_col19" class="data row63 col19" >0</td>
      <td id="T_1666d_row63_col20" class="data row63 col20" >0</td>
      <td id="T_1666d_row63_col21" class="data row63 col21" >0</td>
      <td id="T_1666d_row63_col22" class="data row63 col22" >112,100</td>
      <td id="T_1666d_row63_col23" class="data row63 col23" >0</td>
      <td id="T_1666d_row63_col24" class="data row63 col24" >20,951</td>
      <td id="T_1666d_row63_col25" class="data row63 col25" >0</td>
      <td id="T_1666d_row63_col26" class="data row63 col26" >0</td>
      <td id="T_1666d_row63_col27" class="data row63 col27" >0</td>
      <td id="T_1666d_row63_col28" class="data row63 col28" >120,460</td>
      <td id="T_1666d_row63_col29" class="data row63 col29" >0</td>
      <td id="T_1666d_row63_col30" class="data row63 col30" >506,016</td>
    </tr>
    <tr>
      <th id="T_1666d_level0_row64" class="row_heading level0 row64" rowspan="4">Sew4</th>
      <th id="T_1666d_level1_row64" class="row_heading level1 row64" rowspan="4">Sday_1</th>
      <th id="T_1666d_level2_row64" class="row_heading level2 row64" >D</th>
      <td id="T_1666d_row64_col0" class="data row64 col0" >13,356</td>
      <td id="T_1666d_row64_col1" class="data row64 col1" >34,286</td>
      <td id="T_1666d_row64_col2" class="data row64 col2" >0</td>
      <td id="T_1666d_row64_col3" class="data row64 col3" >0</td>
      <td id="T_1666d_row64_col4" class="data row64 col4" >0</td>
      <td id="T_1666d_row64_col5" class="data row64 col5" >0</td>
      <td id="T_1666d_row64_col6" class="data row64 col6" >4,191</td>
      <td id="T_1666d_row64_col7" class="data row64 col7" >0</td>
      <td id="T_1666d_row64_col8" class="data row64 col8" >0</td>
      <td id="T_1666d_row64_col9" class="data row64 col9" >3,576</td>
      <td id="T_1666d_row64_col10" class="data row64 col10" >0</td>
      <td id="T_1666d_row64_col11" class="data row64 col11" >0</td>
      <td id="T_1666d_row64_col12" class="data row64 col12" >0</td>
      <td id="T_1666d_row64_col13" class="data row64 col13" >0</td>
      <td id="T_1666d_row64_col14" class="data row64 col14" >8,887</td>
      <td id="T_1666d_row64_col15" class="data row64 col15" >0</td>
      <td id="T_1666d_row64_col16" class="data row64 col16" >0</td>
      <td id="T_1666d_row64_col17" class="data row64 col17" >0</td>
      <td id="T_1666d_row64_col18" class="data row64 col18" >0</td>
      <td id="T_1666d_row64_col19" class="data row64 col19" >0</td>
      <td id="T_1666d_row64_col20" class="data row64 col20" >0</td>
      <td id="T_1666d_row64_col21" class="data row64 col21" >20,753</td>
      <td id="T_1666d_row64_col22" class="data row64 col22" >2,865</td>
      <td id="T_1666d_row64_col23" class="data row64 col23" >0</td>
      <td id="T_1666d_row64_col24" class="data row64 col24" >0</td>
      <td id="T_1666d_row64_col25" class="data row64 col25" >0</td>
      <td id="T_1666d_row64_col26" class="data row64 col26" >0</td>
      <td id="T_1666d_row64_col27" class="data row64 col27" >0</td>
      <td id="T_1666d_row64_col28" class="data row64 col28" >1,654</td>
      <td id="T_1666d_row64_col29" class="data row64 col29" >0</td>
      <td id="T_1666d_row64_col30" class="data row64 col30" >89,568</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row65" class="row_heading level2 row65" >G</th>
      <td id="T_1666d_row65_col0" class="data row65 col0" >500</td>
      <td id="T_1666d_row65_col1" class="data row65 col1" >0</td>
      <td id="T_1666d_row65_col2" class="data row65 col2" >0</td>
      <td id="T_1666d_row65_col3" class="data row65 col3" >0</td>
      <td id="T_1666d_row65_col4" class="data row65 col4" >0</td>
      <td id="T_1666d_row65_col5" class="data row65 col5" >0</td>
      <td id="T_1666d_row65_col6" class="data row65 col6" >0</td>
      <td id="T_1666d_row65_col7" class="data row65 col7" >0</td>
      <td id="T_1666d_row65_col8" class="data row65 col8" >0</td>
      <td id="T_1666d_row65_col9" class="data row65 col9" >2,045</td>
      <td id="T_1666d_row65_col10" class="data row65 col10" >0</td>
      <td id="T_1666d_row65_col11" class="data row65 col11" >0</td>
      <td id="T_1666d_row65_col12" class="data row65 col12" >0</td>
      <td id="T_1666d_row65_col13" class="data row65 col13" >0</td>
      <td id="T_1666d_row65_col14" class="data row65 col14" >39,589</td>
      <td id="T_1666d_row65_col15" class="data row65 col15" >0</td>
      <td id="T_1666d_row65_col16" class="data row65 col16" >0</td>
      <td id="T_1666d_row65_col17" class="data row65 col17" >23,849</td>
      <td id="T_1666d_row65_col18" class="data row65 col18" >0</td>
      <td id="T_1666d_row65_col19" class="data row65 col19" >0</td>
      <td id="T_1666d_row65_col20" class="data row65 col20" >52,669</td>
      <td id="T_1666d_row65_col21" class="data row65 col21" >0</td>
      <td id="T_1666d_row65_col22" class="data row65 col22" >3,368</td>
      <td id="T_1666d_row65_col23" class="data row65 col23" >0</td>
      <td id="T_1666d_row65_col24" class="data row65 col24" >0</td>
      <td id="T_1666d_row65_col25" class="data row65 col25" >0</td>
      <td id="T_1666d_row65_col26" class="data row65 col26" >0</td>
      <td id="T_1666d_row65_col27" class="data row65 col27" >0</td>
      <td id="T_1666d_row65_col28" class="data row65 col28" >0</td>
      <td id="T_1666d_row65_col29" class="data row65 col29" >0</td>
      <td id="T_1666d_row65_col30" class="data row65 col30" >122,020</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row66" class="row_heading level2 row66" >I</th>
      <td id="T_1666d_row66_col0" class="data row66 col0" >64,772</td>
      <td id="T_1666d_row66_col1" class="data row66 col1" >48,342</td>
      <td id="T_1666d_row66_col2" class="data row66 col2" >105,840</td>
      <td id="T_1666d_row66_col3" class="data row66 col3" >151,200</td>
      <td id="T_1666d_row66_col4" class="data row66 col4" >0</td>
      <td id="T_1666d_row66_col5" class="data row66 col5" >0</td>
      <td id="T_1666d_row66_col6" class="data row66 col6" >132,617</td>
      <td id="T_1666d_row66_col7" class="data row66 col7" >62,259</td>
      <td id="T_1666d_row66_col8" class="data row66 col8" >62,259</td>
      <td id="T_1666d_row66_col9" class="data row66 col9" >135,716</td>
      <td id="T_1666d_row66_col10" class="data row66 col10" >151,200</td>
      <td id="T_1666d_row66_col11" class="data row66 col11" >151,200</td>
      <td id="T_1666d_row66_col12" class="data row66 col12" >151,200</td>
      <td id="T_1666d_row66_col13" class="data row66 col13" >150,631</td>
      <td id="T_1666d_row66_col14" class="data row66 col14" >500</td>
      <td id="T_1666d_row66_col15" class="data row66 col15" >0</td>
      <td id="T_1666d_row66_col16" class="data row66 col16" >146,076</td>
      <td id="T_1666d_row66_col17" class="data row66 col17" >58,142</td>
      <td id="T_1666d_row66_col18" class="data row66 col18" >0</td>
      <td id="T_1666d_row66_col19" class="data row66 col19" >0</td>
      <td id="T_1666d_row66_col20" class="data row66 col20" >12,987</td>
      <td id="T_1666d_row66_col21" class="data row66 col21" >0</td>
      <td id="T_1666d_row66_col22" class="data row66 col22" >97,249</td>
      <td id="T_1666d_row66_col23" class="data row66 col23" >150,281</td>
      <td id="T_1666d_row66_col24" class="data row66 col24" >0</td>
      <td id="T_1666d_row66_col25" class="data row66 col25" >0</td>
      <td id="T_1666d_row66_col26" class="data row66 col26" >0</td>
      <td id="T_1666d_row66_col27" class="data row66 col27" >62,259</td>
      <td id="T_1666d_row66_col28" class="data row66 col28" >1,462</td>
      <td id="T_1666d_row66_col29" class="data row66 col29" >62,259</td>
      <td id="T_1666d_row66_col30" class="data row66 col30" >1,958,451</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row67" class="row_heading level2 row67" >F</th>
      <td id="T_1666d_row67_col0" class="data row67 col0" >0</td>
      <td id="T_1666d_row67_col1" class="data row67 col1" >0</td>
      <td id="T_1666d_row67_col2" class="data row67 col2" >0</td>
      <td id="T_1666d_row67_col3" class="data row67 col3" >0</td>
      <td id="T_1666d_row67_col4" class="data row67 col4" >0</td>
      <td id="T_1666d_row67_col5" class="data row67 col5" >0</td>
      <td id="T_1666d_row67_col6" class="data row67 col6" >0</td>
      <td id="T_1666d_row67_col7" class="data row67 col7" >0</td>
      <td id="T_1666d_row67_col8" class="data row67 col8" >0</td>
      <td id="T_1666d_row67_col9" class="data row67 col9" >0</td>
      <td id="T_1666d_row67_col10" class="data row67 col10" >0</td>
      <td id="T_1666d_row67_col11" class="data row67 col11" >0</td>
      <td id="T_1666d_row67_col12" class="data row67 col12" >0</td>
      <td id="T_1666d_row67_col13" class="data row67 col13" >0</td>
      <td id="T_1666d_row67_col14" class="data row67 col14" >0</td>
      <td id="T_1666d_row67_col15" class="data row67 col15" >62,259</td>
      <td id="T_1666d_row67_col16" class="data row67 col16" >500</td>
      <td id="T_1666d_row67_col17" class="data row67 col17" >1,500</td>
      <td id="T_1666d_row67_col18" class="data row67 col18" >0</td>
      <td id="T_1666d_row67_col19" class="data row67 col19" >0</td>
      <td id="T_1666d_row67_col20" class="data row67 col20" >0</td>
      <td id="T_1666d_row67_col21" class="data row67 col21" >0</td>
      <td id="T_1666d_row67_col22" class="data row67 col22" >8,594</td>
      <td id="T_1666d_row67_col23" class="data row67 col23" >0</td>
      <td id="T_1666d_row67_col24" class="data row67 col24" >151,200</td>
      <td id="T_1666d_row67_col25" class="data row67 col25" >0</td>
      <td id="T_1666d_row67_col26" class="data row67 col26" >0</td>
      <td id="T_1666d_row67_col27" class="data row67 col27" >0</td>
      <td id="T_1666d_row67_col28" class="data row67 col28" >99,416</td>
      <td id="T_1666d_row67_col29" class="data row67 col29" >0</td>
      <td id="T_1666d_row67_col30" class="data row67 col30" >323,469</td>
    </tr>
    <tr>
      <th id="T_1666d_level0_row68" class="row_heading level0 row68" rowspan="8">QaC1</th>
      <th id="T_1666d_level1_row68" class="row_heading level1 row68" rowspan="8">Day</th>
      <th id="T_1666d_level2_row68" class="row_heading level2 row68" >B</th>
      <td id="T_1666d_row68_col0" class="data row68 col0" >288,000</td>
      <td id="T_1666d_row68_col1" class="data row68 col1" >211,677</td>
      <td id="T_1666d_row68_col2" class="data row68 col2" >88,673</td>
      <td id="T_1666d_row68_col3" class="data row68 col3" >32,142</td>
      <td id="T_1666d_row68_col4" class="data row68 col4" >0</td>
      <td id="T_1666d_row68_col5" class="data row68 col5" >0</td>
      <td id="T_1666d_row68_col6" class="data row68 col6" >133,545</td>
      <td id="T_1666d_row68_col7" class="data row68 col7" >0</td>
      <td id="T_1666d_row68_col8" class="data row68 col8" >44,470</td>
      <td id="T_1666d_row68_col9" class="data row68 col9" >172,509</td>
      <td id="T_1666d_row68_col10" class="data row68 col10" >0</td>
      <td id="T_1666d_row68_col11" class="data row68 col11" >0</td>
      <td id="T_1666d_row68_col12" class="data row68 col12" >0</td>
      <td id="T_1666d_row68_col13" class="data row68 col13" >0</td>
      <td id="T_1666d_row68_col14" class="data row68 col14" >0</td>
      <td id="T_1666d_row68_col15" class="data row68 col15" >0</td>
      <td id="T_1666d_row68_col16" class="data row68 col16" >118,589</td>
      <td id="T_1666d_row68_col17" class="data row68 col17" >0</td>
      <td id="T_1666d_row68_col18" class="data row68 col18" >0</td>
      <td id="T_1666d_row68_col19" class="data row68 col19" >0</td>
      <td id="T_1666d_row68_col20" class="data row68 col20" >201,600</td>
      <td id="T_1666d_row68_col21" class="data row68 col21" >0</td>
      <td id="T_1666d_row68_col22" class="data row68 col22" >0</td>
      <td id="T_1666d_row68_col23" class="data row68 col23" >0</td>
      <td id="T_1666d_row68_col24" class="data row68 col24" >0</td>
      <td id="T_1666d_row68_col25" class="data row68 col25" >0</td>
      <td id="T_1666d_row68_col26" class="data row68 col26" >0</td>
      <td id="T_1666d_row68_col27" class="data row68 col27" >0</td>
      <td id="T_1666d_row68_col28" class="data row68 col28" >0</td>
      <td id="T_1666d_row68_col29" class="data row68 col29" >0</td>
      <td id="T_1666d_row68_col30" class="data row68 col30" >1,291,205</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row69" class="row_heading level2 row69" >D</th>
      <td id="T_1666d_row69_col0" class="data row69 col0" >0</td>
      <td id="T_1666d_row69_col1" class="data row69 col1" >1,055</td>
      <td id="T_1666d_row69_col2" class="data row69 col2" >0</td>
      <td id="T_1666d_row69_col3" class="data row69 col3" >0</td>
      <td id="T_1666d_row69_col4" class="data row69 col4" >0</td>
      <td id="T_1666d_row69_col5" class="data row69 col5" >0</td>
      <td id="T_1666d_row69_col6" class="data row69 col6" >0</td>
      <td id="T_1666d_row69_col7" class="data row69 col7" >0</td>
      <td id="T_1666d_row69_col8" class="data row69 col8" >32,169</td>
      <td id="T_1666d_row69_col9" class="data row69 col9" >0</td>
      <td id="T_1666d_row69_col10" class="data row69 col10" >0</td>
      <td id="T_1666d_row69_col11" class="data row69 col11" >0</td>
      <td id="T_1666d_row69_col12" class="data row69 col12" >0</td>
      <td id="T_1666d_row69_col13" class="data row69 col13" >0</td>
      <td id="T_1666d_row69_col14" class="data row69 col14" >0</td>
      <td id="T_1666d_row69_col15" class="data row69 col15" >0</td>
      <td id="T_1666d_row69_col16" class="data row69 col16" >0</td>
      <td id="T_1666d_row69_col17" class="data row69 col17" >44,800</td>
      <td id="T_1666d_row69_col18" class="data row69 col18" >0</td>
      <td id="T_1666d_row69_col19" class="data row69 col19" >0</td>
      <td id="T_1666d_row69_col20" class="data row69 col20" >0</td>
      <td id="T_1666d_row69_col21" class="data row69 col21" >26,353</td>
      <td id="T_1666d_row69_col22" class="data row69 col22" >0</td>
      <td id="T_1666d_row69_col23" class="data row69 col23" >0</td>
      <td id="T_1666d_row69_col24" class="data row69 col24" >0</td>
      <td id="T_1666d_row69_col25" class="data row69 col25" >0</td>
      <td id="T_1666d_row69_col26" class="data row69 col26" >0</td>
      <td id="T_1666d_row69_col27" class="data row69 col27" >44,800</td>
      <td id="T_1666d_row69_col28" class="data row69 col28" >0</td>
      <td id="T_1666d_row69_col29" class="data row69 col29" >0</td>
      <td id="T_1666d_row69_col30" class="data row69 col30" >149,177</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row70" class="row_heading level2 row70" >G</th>
      <td id="T_1666d_row70_col0" class="data row70 col0" >0</td>
      <td id="T_1666d_row70_col1" class="data row70 col1" >0</td>
      <td id="T_1666d_row70_col2" class="data row70 col2" >0</td>
      <td id="T_1666d_row70_col3" class="data row70 col3" >0</td>
      <td id="T_1666d_row70_col4" class="data row70 col4" >0</td>
      <td id="T_1666d_row70_col5" class="data row70 col5" >0</td>
      <td id="T_1666d_row70_col6" class="data row70 col6" >0</td>
      <td id="T_1666d_row70_col7" class="data row70 col7" >0</td>
      <td id="T_1666d_row70_col8" class="data row70 col8" >0</td>
      <td id="T_1666d_row70_col9" class="data row70 col9" >0</td>
      <td id="T_1666d_row70_col10" class="data row70 col10" >0</td>
      <td id="T_1666d_row70_col11" class="data row70 col11" >0</td>
      <td id="T_1666d_row70_col12" class="data row70 col12" >0</td>
      <td id="T_1666d_row70_col13" class="data row70 col13" >0</td>
      <td id="T_1666d_row70_col14" class="data row70 col14" >0</td>
      <td id="T_1666d_row70_col15" class="data row70 col15" >0</td>
      <td id="T_1666d_row70_col16" class="data row70 col16" >0</td>
      <td id="T_1666d_row70_col17" class="data row70 col17" >0</td>
      <td id="T_1666d_row70_col18" class="data row70 col18" >0</td>
      <td id="T_1666d_row70_col19" class="data row70 col19" >0</td>
      <td id="T_1666d_row70_col20" class="data row70 col20" >0</td>
      <td id="T_1666d_row70_col21" class="data row70 col21" >0</td>
      <td id="T_1666d_row70_col22" class="data row70 col22" >0</td>
      <td id="T_1666d_row70_col23" class="data row70 col23" >0</td>
      <td id="T_1666d_row70_col24" class="data row70 col24" >571</td>
      <td id="T_1666d_row70_col25" class="data row70 col25" >0</td>
      <td id="T_1666d_row70_col26" class="data row70 col26" >0</td>
      <td id="T_1666d_row70_col27" class="data row70 col27" >0</td>
      <td id="T_1666d_row70_col28" class="data row70 col28" >0</td>
      <td id="T_1666d_row70_col29" class="data row70 col29" >0</td>
      <td id="T_1666d_row70_col30" class="data row70 col30" >571</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row71" class="row_heading level2 row71" >A</th>
      <td id="T_1666d_row71_col0" class="data row71 col0" >0</td>
      <td id="T_1666d_row71_col1" class="data row71 col1" >71,574</td>
      <td id="T_1666d_row71_col2" class="data row71 col2" >136,175</td>
      <td id="T_1666d_row71_col3" class="data row71 col3" >198,696</td>
      <td id="T_1666d_row71_col4" class="data row71 col4" >0</td>
      <td id="T_1666d_row71_col5" class="data row71 col5" >0</td>
      <td id="T_1666d_row71_col6" class="data row71 col6" >0</td>
      <td id="T_1666d_row71_col7" class="data row71 col7" >0</td>
      <td id="T_1666d_row71_col8" class="data row71 col8" >0</td>
      <td id="T_1666d_row71_col9" class="data row71 col9" >0</td>
      <td id="T_1666d_row71_col10" class="data row71 col10" >0</td>
      <td id="T_1666d_row71_col11" class="data row71 col11" >0</td>
      <td id="T_1666d_row71_col12" class="data row71 col12" >100,799</td>
      <td id="T_1666d_row71_col13" class="data row71 col13" >174,882</td>
      <td id="T_1666d_row71_col14" class="data row71 col14" >0</td>
      <td id="T_1666d_row71_col15" class="data row71 col15" >0</td>
      <td id="T_1666d_row71_col16" class="data row71 col16" >0</td>
      <td id="T_1666d_row71_col17" class="data row71 col17" >0</td>
      <td id="T_1666d_row71_col18" class="data row71 col18" >0</td>
      <td id="T_1666d_row71_col19" class="data row71 col19" >0</td>
      <td id="T_1666d_row71_col20" class="data row71 col20" >0</td>
      <td id="T_1666d_row71_col21" class="data row71 col21" >0</td>
      <td id="T_1666d_row71_col22" class="data row71 col22" >0</td>
      <td id="T_1666d_row71_col23" class="data row71 col23" >0</td>
      <td id="T_1666d_row71_col24" class="data row71 col24" >0</td>
      <td id="T_1666d_row71_col25" class="data row71 col25" >0</td>
      <td id="T_1666d_row71_col26" class="data row71 col26" >0</td>
      <td id="T_1666d_row71_col27" class="data row71 col27" >0</td>
      <td id="T_1666d_row71_col28" class="data row71 col28" >0</td>
      <td id="T_1666d_row71_col29" class="data row71 col29" >0</td>
      <td id="T_1666d_row71_col30" class="data row71 col30" >682,126</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row72" class="row_heading level2 row72" >I</th>
      <td id="T_1666d_row72_col0" class="data row72 col0" >0</td>
      <td id="T_1666d_row72_col1" class="data row72 col1" >0</td>
      <td id="T_1666d_row72_col2" class="data row72 col2" >501</td>
      <td id="T_1666d_row72_col3" class="data row72 col3" >0</td>
      <td id="T_1666d_row72_col4" class="data row72 col4" >0</td>
      <td id="T_1666d_row72_col5" class="data row72 col5" >0</td>
      <td id="T_1666d_row72_col6" class="data row72 col6" >501</td>
      <td id="T_1666d_row72_col7" class="data row72 col7" >0</td>
      <td id="T_1666d_row72_col8" class="data row72 col8" >16,494</td>
      <td id="T_1666d_row72_col9" class="data row72 col9" >0</td>
      <td id="T_1666d_row72_col10" class="data row72 col10" >243,724</td>
      <td id="T_1666d_row72_col11" class="data row72 col11" >157,773</td>
      <td id="T_1666d_row72_col12" class="data row72 col12" >0</td>
      <td id="T_1666d_row72_col13" class="data row72 col13" >0</td>
      <td id="T_1666d_row72_col14" class="data row72 col14" >268,800</td>
      <td id="T_1666d_row72_col15" class="data row72 col15" >105,012</td>
      <td id="T_1666d_row72_col16" class="data row72 col16" >0</td>
      <td id="T_1666d_row72_col17" class="data row72 col17" >0</td>
      <td id="T_1666d_row72_col18" class="data row72 col18" >0</td>
      <td id="T_1666d_row72_col19" class="data row72 col19" >0</td>
      <td id="T_1666d_row72_col20" class="data row72 col20" >0</td>
      <td id="T_1666d_row72_col21" class="data row72 col21" >0</td>
      <td id="T_1666d_row72_col22" class="data row72 col22" >0</td>
      <td id="T_1666d_row72_col23" class="data row72 col23" >177,389</td>
      <td id="T_1666d_row72_col24" class="data row72 col24" >0</td>
      <td id="T_1666d_row72_col25" class="data row72 col25" >0</td>
      <td id="T_1666d_row72_col26" class="data row72 col26" >0</td>
      <td id="T_1666d_row72_col27" class="data row72 col27" >0</td>
      <td id="T_1666d_row72_col28" class="data row72 col28" >500</td>
      <td id="T_1666d_row72_col29" class="data row72 col29" >0</td>
      <td id="T_1666d_row72_col30" class="data row72 col30" >970,694</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row73" class="row_heading level2 row73" >E</th>
      <td id="T_1666d_row73_col0" class="data row73 col0" >0</td>
      <td id="T_1666d_row73_col1" class="data row73 col1" >0</td>
      <td id="T_1666d_row73_col2" class="data row73 col2" >0</td>
      <td id="T_1666d_row73_col3" class="data row73 col3" >0</td>
      <td id="T_1666d_row73_col4" class="data row73 col4" >0</td>
      <td id="T_1666d_row73_col5" class="data row73 col5" >0</td>
      <td id="T_1666d_row73_col6" class="data row73 col6" >0</td>
      <td id="T_1666d_row73_col7" class="data row73 col7" >500</td>
      <td id="T_1666d_row73_col8" class="data row73 col8" >0</td>
      <td id="T_1666d_row73_col9" class="data row73 col9" >0</td>
      <td id="T_1666d_row73_col10" class="data row73 col10" >0</td>
      <td id="T_1666d_row73_col11" class="data row73 col11" >0</td>
      <td id="T_1666d_row73_col12" class="data row73 col12" >0</td>
      <td id="T_1666d_row73_col13" class="data row73 col13" >0</td>
      <td id="T_1666d_row73_col14" class="data row73 col14" >0</td>
      <td id="T_1666d_row73_col15" class="data row73 col15" >0</td>
      <td id="T_1666d_row73_col16" class="data row73 col16" >0</td>
      <td id="T_1666d_row73_col17" class="data row73 col17" >0</td>
      <td id="T_1666d_row73_col18" class="data row73 col18" >0</td>
      <td id="T_1666d_row73_col19" class="data row73 col19" >0</td>
      <td id="T_1666d_row73_col20" class="data row73 col20" >0</td>
      <td id="T_1666d_row73_col21" class="data row73 col21" >0</td>
      <td id="T_1666d_row73_col22" class="data row73 col22" >0</td>
      <td id="T_1666d_row73_col23" class="data row73 col23" >0</td>
      <td id="T_1666d_row73_col24" class="data row73 col24" >0</td>
      <td id="T_1666d_row73_col25" class="data row73 col25" >0</td>
      <td id="T_1666d_row73_col26" class="data row73 col26" >0</td>
      <td id="T_1666d_row73_col27" class="data row73 col27" >0</td>
      <td id="T_1666d_row73_col28" class="data row73 col28" >0</td>
      <td id="T_1666d_row73_col29" class="data row73 col29" >0</td>
      <td id="T_1666d_row73_col30" class="data row73 col30" >500</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row74" class="row_heading level2 row74" >C</th>
      <td id="T_1666d_row74_col0" class="data row74 col0" >0</td>
      <td id="T_1666d_row74_col1" class="data row74 col1" >0</td>
      <td id="T_1666d_row74_col2" class="data row74 col2" >40,885</td>
      <td id="T_1666d_row74_col3" class="data row74 col3" >0</td>
      <td id="T_1666d_row74_col4" class="data row74 col4" >0</td>
      <td id="T_1666d_row74_col5" class="data row74 col5" >0</td>
      <td id="T_1666d_row74_col6" class="data row74 col6" >123,263</td>
      <td id="T_1666d_row74_col7" class="data row74 col7" >94,871</td>
      <td id="T_1666d_row74_col8" class="data row74 col8" >0</td>
      <td id="T_1666d_row74_col9" class="data row74 col9" >23,272</td>
      <td id="T_1666d_row74_col10" class="data row74 col10" >64,463</td>
      <td id="T_1666d_row74_col11" class="data row74 col11" >0</td>
      <td id="T_1666d_row74_col12" class="data row74 col12" >0</td>
      <td id="T_1666d_row74_col13" class="data row74 col13" >21,374</td>
      <td id="T_1666d_row74_col14" class="data row74 col14" >0</td>
      <td id="T_1666d_row74_col15" class="data row74 col15" >81,435</td>
      <td id="T_1666d_row74_col16" class="data row74 col16" >0</td>
      <td id="T_1666d_row74_col17" class="data row74 col17" >0</td>
      <td id="T_1666d_row74_col18" class="data row74 col18" >0</td>
      <td id="T_1666d_row74_col19" class="data row74 col19" >0</td>
      <td id="T_1666d_row74_col20" class="data row74 col20" >0</td>
      <td id="T_1666d_row74_col21" class="data row74 col21" >0</td>
      <td id="T_1666d_row74_col22" class="data row74 col22" >161,280</td>
      <td id="T_1666d_row74_col23" class="data row74 col23" >0</td>
      <td id="T_1666d_row74_col24" class="data row74 col24" >131,719</td>
      <td id="T_1666d_row74_col25" class="data row74 col25" >0</td>
      <td id="T_1666d_row74_col26" class="data row74 col26" >0</td>
      <td id="T_1666d_row74_col27" class="data row74 col27" >0</td>
      <td id="T_1666d_row74_col28" class="data row74 col28" >230,100</td>
      <td id="T_1666d_row74_col29" class="data row74 col29" >230,400</td>
      <td id="T_1666d_row74_col30" class="data row74 col30" >1,203,062</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row75" class="row_heading level2 row75" >F</th>
      <td id="T_1666d_row75_col0" class="data row75 col0" >0</td>
      <td id="T_1666d_row75_col1" class="data row75 col1" >0</td>
      <td id="T_1666d_row75_col2" class="data row75 col2" >15,560</td>
      <td id="T_1666d_row75_col3" class="data row75 col3" >76,216</td>
      <td id="T_1666d_row75_col4" class="data row75 col4" >0</td>
      <td id="T_1666d_row75_col5" class="data row75 col5" >0</td>
      <td id="T_1666d_row75_col6" class="data row75 col6" >0</td>
      <td id="T_1666d_row75_col7" class="data row75 col7" >0</td>
      <td id="T_1666d_row75_col8" class="data row75 col8" >0</td>
      <td id="T_1666d_row75_col9" class="data row75 col9" >0</td>
      <td id="T_1666d_row75_col10" class="data row75 col10" >0</td>
      <td id="T_1666d_row75_col11" class="data row75 col11" >0</td>
      <td id="T_1666d_row75_col12" class="data row75 col12" >0</td>
      <td id="T_1666d_row75_col13" class="data row75 col13" >0</td>
      <td id="T_1666d_row75_col14" class="data row75 col14" >0</td>
      <td id="T_1666d_row75_col15" class="data row75 col15" >28,063</td>
      <td id="T_1666d_row75_col16" class="data row75 col16" >0</td>
      <td id="T_1666d_row75_col17" class="data row75 col17" >0</td>
      <td id="T_1666d_row75_col18" class="data row75 col18" >0</td>
      <td id="T_1666d_row75_col19" class="data row75 col19" >0</td>
      <td id="T_1666d_row75_col20" class="data row75 col20" >0</td>
      <td id="T_1666d_row75_col21" class="data row75 col21" >0</td>
      <td id="T_1666d_row75_col22" class="data row75 col22" >0</td>
      <td id="T_1666d_row75_col23" class="data row75 col23" >0</td>
      <td id="T_1666d_row75_col24" class="data row75 col24" >48,317</td>
      <td id="T_1666d_row75_col25" class="data row75 col25" >0</td>
      <td id="T_1666d_row75_col26" class="data row75 col26" >0</td>
      <td id="T_1666d_row75_col27" class="data row75 col27" >0</td>
      <td id="T_1666d_row75_col28" class="data row75 col28" >0</td>
      <td id="T_1666d_row75_col29" class="data row75 col29" >0</td>
      <td id="T_1666d_row75_col30" class="data row75 col30" >168,156</td>
    </tr>
    <tr>
      <th id="T_1666d_level0_row76" class="row_heading level0 row76" rowspan="17">QaC2</th>
      <th id="T_1666d_level1_row76" class="row_heading level1 row76" rowspan="9">Sday_1</th>
      <th id="T_1666d_level2_row76" class="row_heading level2 row76" >B</th>
      <td id="T_1666d_row76_col0" class="data row76 col0" >0</td>
      <td id="T_1666d_row76_col1" class="data row76 col1" >0</td>
      <td id="T_1666d_row76_col2" class="data row76 col2" >0</td>
      <td id="T_1666d_row76_col3" class="data row76 col3" >0</td>
      <td id="T_1666d_row76_col4" class="data row76 col4" >0</td>
      <td id="T_1666d_row76_col5" class="data row76 col5" >0</td>
      <td id="T_1666d_row76_col6" class="data row76 col6" >0</td>
      <td id="T_1666d_row76_col7" class="data row76 col7" >50,790</td>
      <td id="T_1666d_row76_col8" class="data row76 col8" >103,765</td>
      <td id="T_1666d_row76_col9" class="data row76 col9" >0</td>
      <td id="T_1666d_row76_col10" class="data row76 col10" >1</td>
      <td id="T_1666d_row76_col11" class="data row76 col11" >0</td>
      <td id="T_1666d_row76_col12" class="data row76 col12" >0</td>
      <td id="T_1666d_row76_col13" class="data row76 col13" >0</td>
      <td id="T_1666d_row76_col14" class="data row76 col14" >0</td>
      <td id="T_1666d_row76_col15" class="data row76 col15" >0</td>
      <td id="T_1666d_row76_col16" class="data row76 col16" >43,910</td>
      <td id="T_1666d_row76_col17" class="data row76 col17" >0</td>
      <td id="T_1666d_row76_col18" class="data row76 col18" >0</td>
      <td id="T_1666d_row76_col19" class="data row76 col19" >0</td>
      <td id="T_1666d_row76_col20" class="data row76 col20" >0</td>
      <td id="T_1666d_row76_col21" class="data row76 col21" >95,883</td>
      <td id="T_1666d_row76_col22" class="data row76 col22" >0</td>
      <td id="T_1666d_row76_col23" class="data row76 col23" >0</td>
      <td id="T_1666d_row76_col24" class="data row76 col24" >0</td>
      <td id="T_1666d_row76_col25" class="data row76 col25" >0</td>
      <td id="T_1666d_row76_col26" class="data row76 col26" >0</td>
      <td id="T_1666d_row76_col27" class="data row76 col27" >0</td>
      <td id="T_1666d_row76_col28" class="data row76 col28" >0</td>
      <td id="T_1666d_row76_col29" class="data row76 col29" >0</td>
      <td id="T_1666d_row76_col30" class="data row76 col30" >294,349</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row77" class="row_heading level2 row77" >H</th>
      <td id="T_1666d_row77_col0" class="data row77 col0" >0</td>
      <td id="T_1666d_row77_col1" class="data row77 col1" >0</td>
      <td id="T_1666d_row77_col2" class="data row77 col2" >0</td>
      <td id="T_1666d_row77_col3" class="data row77 col3" >0</td>
      <td id="T_1666d_row77_col4" class="data row77 col4" >0</td>
      <td id="T_1666d_row77_col5" class="data row77 col5" >0</td>
      <td id="T_1666d_row77_col6" class="data row77 col6" >0</td>
      <td id="T_1666d_row77_col7" class="data row77 col7" >0</td>
      <td id="T_1666d_row77_col8" class="data row77 col8" >0</td>
      <td id="T_1666d_row77_col9" class="data row77 col9" >0</td>
      <td id="T_1666d_row77_col10" class="data row77 col10" >0</td>
      <td id="T_1666d_row77_col11" class="data row77 col11" >0</td>
      <td id="T_1666d_row77_col12" class="data row77 col12" >0</td>
      <td id="T_1666d_row77_col13" class="data row77 col13" >0</td>
      <td id="T_1666d_row77_col14" class="data row77 col14" >0</td>
      <td id="T_1666d_row77_col15" class="data row77 col15" >0</td>
      <td id="T_1666d_row77_col16" class="data row77 col16" >0</td>
      <td id="T_1666d_row77_col17" class="data row77 col17" >0</td>
      <td id="T_1666d_row77_col18" class="data row77 col18" >0</td>
      <td id="T_1666d_row77_col19" class="data row77 col19" >0</td>
      <td id="T_1666d_row77_col20" class="data row77 col20" >0</td>
      <td id="T_1666d_row77_col21" class="data row77 col21" >0</td>
      <td id="T_1666d_row77_col22" class="data row77 col22" >0</td>
      <td id="T_1666d_row77_col23" class="data row77 col23" >0</td>
      <td id="T_1666d_row77_col24" class="data row77 col24" >0</td>
      <td id="T_1666d_row77_col25" class="data row77 col25" >0</td>
      <td id="T_1666d_row77_col26" class="data row77 col26" >0</td>
      <td id="T_1666d_row77_col27" class="data row77 col27" >0</td>
      <td id="T_1666d_row77_col28" class="data row77 col28" >0</td>
      <td id="T_1666d_row77_col29" class="data row77 col29" >1,437</td>
      <td id="T_1666d_row77_col30" class="data row77 col30" >1,437</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row78" class="row_heading level2 row78" >D</th>
      <td id="T_1666d_row78_col0" class="data row78 col0" >0</td>
      <td id="T_1666d_row78_col1" class="data row78 col1" >3,865</td>
      <td id="T_1666d_row78_col2" class="data row78 col2" >0</td>
      <td id="T_1666d_row78_col3" class="data row78 col3" >0</td>
      <td id="T_1666d_row78_col4" class="data row78 col4" >0</td>
      <td id="T_1666d_row78_col5" class="data row78 col5" >0</td>
      <td id="T_1666d_row78_col6" class="data row78 col6" >0</td>
      <td id="T_1666d_row78_col7" class="data row78 col7" >0</td>
      <td id="T_1666d_row78_col8" class="data row78 col8" >0</td>
      <td id="T_1666d_row78_col9" class="data row78 col9" >0</td>
      <td id="T_1666d_row78_col10" class="data row78 col10" >0</td>
      <td id="T_1666d_row78_col11" class="data row78 col11" >0</td>
      <td id="T_1666d_row78_col12" class="data row78 col12" >0</td>
      <td id="T_1666d_row78_col13" class="data row78 col13" >0</td>
      <td id="T_1666d_row78_col14" class="data row78 col14" >14,027</td>
      <td id="T_1666d_row78_col15" class="data row78 col15" >0</td>
      <td id="T_1666d_row78_col16" class="data row78 col16" >0</td>
      <td id="T_1666d_row78_col17" class="data row78 col17" >0</td>
      <td id="T_1666d_row78_col18" class="data row78 col18" >0</td>
      <td id="T_1666d_row78_col19" class="data row78 col19" >0</td>
      <td id="T_1666d_row78_col20" class="data row78 col20" >0</td>
      <td id="T_1666d_row78_col21" class="data row78 col21" >0</td>
      <td id="T_1666d_row78_col22" class="data row78 col22" >0</td>
      <td id="T_1666d_row78_col23" class="data row78 col23" >4,575</td>
      <td id="T_1666d_row78_col24" class="data row78 col24" >0</td>
      <td id="T_1666d_row78_col25" class="data row78 col25" >0</td>
      <td id="T_1666d_row78_col26" class="data row78 col26" >0</td>
      <td id="T_1666d_row78_col27" class="data row78 col27" >0</td>
      <td id="T_1666d_row78_col28" class="data row78 col28" >166</td>
      <td id="T_1666d_row78_col29" class="data row78 col29" >0</td>
      <td id="T_1666d_row78_col30" class="data row78 col30" >22,633</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row79" class="row_heading level2 row79" >G</th>
      <td id="T_1666d_row79_col0" class="data row79 col0" >0</td>
      <td id="T_1666d_row79_col1" class="data row79 col1" >0</td>
      <td id="T_1666d_row79_col2" class="data row79 col2" >0</td>
      <td id="T_1666d_row79_col3" class="data row79 col3" >0</td>
      <td id="T_1666d_row79_col4" class="data row79 col4" >0</td>
      <td id="T_1666d_row79_col5" class="data row79 col5" >0</td>
      <td id="T_1666d_row79_col6" class="data row79 col6" >0</td>
      <td id="T_1666d_row79_col7" class="data row79 col7" >0</td>
      <td id="T_1666d_row79_col8" class="data row79 col8" >0</td>
      <td id="T_1666d_row79_col9" class="data row79 col9" >0</td>
      <td id="T_1666d_row79_col10" class="data row79 col10" >8,937</td>
      <td id="T_1666d_row79_col11" class="data row79 col11" >0</td>
      <td id="T_1666d_row79_col12" class="data row79 col12" >0</td>
      <td id="T_1666d_row79_col13" class="data row79 col13" >0</td>
      <td id="T_1666d_row79_col14" class="data row79 col14" >0</td>
      <td id="T_1666d_row79_col15" class="data row79 col15" >0</td>
      <td id="T_1666d_row79_col16" class="data row79 col16" >1</td>
      <td id="T_1666d_row79_col17" class="data row79 col17" >0</td>
      <td id="T_1666d_row79_col18" class="data row79 col18" >0</td>
      <td id="T_1666d_row79_col19" class="data row79 col19" >0</td>
      <td id="T_1666d_row79_col20" class="data row79 col20" >0</td>
      <td id="T_1666d_row79_col21" class="data row79 col21" >298</td>
      <td id="T_1666d_row79_col22" class="data row79 col22" >0</td>
      <td id="T_1666d_row79_col23" class="data row79 col23" >0</td>
      <td id="T_1666d_row79_col24" class="data row79 col24" >0</td>
      <td id="T_1666d_row79_col25" class="data row79 col25" >0</td>
      <td id="T_1666d_row79_col26" class="data row79 col26" >0</td>
      <td id="T_1666d_row79_col27" class="data row79 col27" >0</td>
      <td id="T_1666d_row79_col28" class="data row79 col28" >0</td>
      <td id="T_1666d_row79_col29" class="data row79 col29" >0</td>
      <td id="T_1666d_row79_col30" class="data row79 col30" >9,236</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row80" class="row_heading level2 row80" >A</th>
      <td id="T_1666d_row80_col0" class="data row80 col0" >94,041</td>
      <td id="T_1666d_row80_col1" class="data row80 col1" >0</td>
      <td id="T_1666d_row80_col2" class="data row80 col2" >194,922</td>
      <td id="T_1666d_row80_col3" class="data row80 col3" >200,974</td>
      <td id="T_1666d_row80_col4" class="data row80 col4" >0</td>
      <td id="T_1666d_row80_col5" class="data row80 col5" >0</td>
      <td id="T_1666d_row80_col6" class="data row80 col6" >0</td>
      <td id="T_1666d_row80_col7" class="data row80 col7" >0</td>
      <td id="T_1666d_row80_col8" class="data row80 col8" >0</td>
      <td id="T_1666d_row80_col9" class="data row80 col9" >0</td>
      <td id="T_1666d_row80_col10" class="data row80 col10" >0</td>
      <td id="T_1666d_row80_col11" class="data row80 col11" >0</td>
      <td id="T_1666d_row80_col12" class="data row80 col12" >176,400</td>
      <td id="T_1666d_row80_col13" class="data row80 col13" >0</td>
      <td id="T_1666d_row80_col14" class="data row80 col14" >0</td>
      <td id="T_1666d_row80_col15" class="data row80 col15" >0</td>
      <td id="T_1666d_row80_col16" class="data row80 col16" >0</td>
      <td id="T_1666d_row80_col17" class="data row80 col17" >0</td>
      <td id="T_1666d_row80_col18" class="data row80 col18" >0</td>
      <td id="T_1666d_row80_col19" class="data row80 col19" >0</td>
      <td id="T_1666d_row80_col20" class="data row80 col20" >0</td>
      <td id="T_1666d_row80_col21" class="data row80 col21" >0</td>
      <td id="T_1666d_row80_col22" class="data row80 col22" >0</td>
      <td id="T_1666d_row80_col23" class="data row80 col23" >155,813</td>
      <td id="T_1666d_row80_col24" class="data row80 col24" >0</td>
      <td id="T_1666d_row80_col25" class="data row80 col25" >0</td>
      <td id="T_1666d_row80_col26" class="data row80 col26" >0</td>
      <td id="T_1666d_row80_col27" class="data row80 col27" >0</td>
      <td id="T_1666d_row80_col28" class="data row80 col28" >0</td>
      <td id="T_1666d_row80_col29" class="data row80 col29" >0</td>
      <td id="T_1666d_row80_col30" class="data row80 col30" >822,150</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row81" class="row_heading level2 row81" >I</th>
      <td id="T_1666d_row81_col0" class="data row81 col0" >109,813</td>
      <td id="T_1666d_row81_col1" class="data row81 col1" >312,810</td>
      <td id="T_1666d_row81_col2" class="data row81 col2" >76,104</td>
      <td id="T_1666d_row81_col3" class="data row81 col3" >1,584</td>
      <td id="T_1666d_row81_col4" class="data row81 col4" >0</td>
      <td id="T_1666d_row81_col5" class="data row81 col5" >0</td>
      <td id="T_1666d_row81_col6" class="data row81 col6" >336,000</td>
      <td id="T_1666d_row81_col7" class="data row81 col7" >0</td>
      <td id="T_1666d_row81_col8" class="data row81 col8" >2,656</td>
      <td id="T_1666d_row81_col9" class="data row81 col9" >336,000</td>
      <td id="T_1666d_row81_col10" class="data row81 col10" >95,342</td>
      <td id="T_1666d_row81_col11" class="data row81 col11" >128,542</td>
      <td id="T_1666d_row81_col12" class="data row81 col12" >0</td>
      <td id="T_1666d_row81_col13" class="data row81 col13" >146,021</td>
      <td id="T_1666d_row81_col14" class="data row81 col14" >0</td>
      <td id="T_1666d_row81_col15" class="data row81 col15" >83,229</td>
      <td id="T_1666d_row81_col16" class="data row81 col16" >128,190</td>
      <td id="T_1666d_row81_col17" class="data row81 col17" >236,382</td>
      <td id="T_1666d_row81_col18" class="data row81 col18" >0</td>
      <td id="T_1666d_row81_col19" class="data row81 col19" >0</td>
      <td id="T_1666d_row81_col20" class="data row81 col20" >235,200</td>
      <td id="T_1666d_row81_col21" class="data row81 col21" >55,938</td>
      <td id="T_1666d_row81_col22" class="data row81 col22" >0</td>
      <td id="T_1666d_row81_col23" class="data row81 col23" >0</td>
      <td id="T_1666d_row81_col24" class="data row81 col24" >2,501</td>
      <td id="T_1666d_row81_col25" class="data row81 col25" >0</td>
      <td id="T_1666d_row81_col26" class="data row81 col26" >0</td>
      <td id="T_1666d_row81_col27" class="data row81 col27" >138,077</td>
      <td id="T_1666d_row81_col28" class="data row81 col28" >500</td>
      <td id="T_1666d_row81_col29" class="data row81 col29" >0</td>
      <td id="T_1666d_row81_col30" class="data row81 col30" >2,424,889</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row82" class="row_heading level2 row82" >E</th>
      <td id="T_1666d_row82_col0" class="data row82 col0" >0</td>
      <td id="T_1666d_row82_col1" class="data row82 col1" >0</td>
      <td id="T_1666d_row82_col2" class="data row82 col2" >0</td>
      <td id="T_1666d_row82_col3" class="data row82 col3" >0</td>
      <td id="T_1666d_row82_col4" class="data row82 col4" >0</td>
      <td id="T_1666d_row82_col5" class="data row82 col5" >0</td>
      <td id="T_1666d_row82_col6" class="data row82 col6" >0</td>
      <td id="T_1666d_row82_col7" class="data row82 col7" >0</td>
      <td id="T_1666d_row82_col8" class="data row82 col8" >0</td>
      <td id="T_1666d_row82_col9" class="data row82 col9" >0</td>
      <td id="T_1666d_row82_col10" class="data row82 col10" >0</td>
      <td id="T_1666d_row82_col11" class="data row82 col11" >0</td>
      <td id="T_1666d_row82_col12" class="data row82 col12" >0</td>
      <td id="T_1666d_row82_col13" class="data row82 col13" >0</td>
      <td id="T_1666d_row82_col14" class="data row82 col14" >0</td>
      <td id="T_1666d_row82_col15" class="data row82 col15" >0</td>
      <td id="T_1666d_row82_col16" class="data row82 col16" >0</td>
      <td id="T_1666d_row82_col17" class="data row82 col17" >0</td>
      <td id="T_1666d_row82_col18" class="data row82 col18" >0</td>
      <td id="T_1666d_row82_col19" class="data row82 col19" >0</td>
      <td id="T_1666d_row82_col20" class="data row82 col20" >0</td>
      <td id="T_1666d_row82_col21" class="data row82 col21" >0</td>
      <td id="T_1666d_row82_col22" class="data row82 col22" >0</td>
      <td id="T_1666d_row82_col23" class="data row82 col23" >0</td>
      <td id="T_1666d_row82_col24" class="data row82 col24" >0</td>
      <td id="T_1666d_row82_col25" class="data row82 col25" >0</td>
      <td id="T_1666d_row82_col26" class="data row82 col26" >0</td>
      <td id="T_1666d_row82_col27" class="data row82 col27" >499</td>
      <td id="T_1666d_row82_col28" class="data row82 col28" >0</td>
      <td id="T_1666d_row82_col29" class="data row82 col29" >0</td>
      <td id="T_1666d_row82_col30" class="data row82 col30" >499</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row83" class="row_heading level2 row83" >C</th>
      <td id="T_1666d_row83_col0" class="data row83 col0" >0</td>
      <td id="T_1666d_row83_col1" class="data row83 col1" >0</td>
      <td id="T_1666d_row83_col2" class="data row83 col2" >0</td>
      <td id="T_1666d_row83_col3" class="data row83 col3" >0</td>
      <td id="T_1666d_row83_col4" class="data row83 col4" >0</td>
      <td id="T_1666d_row83_col5" class="data row83 col5" >0</td>
      <td id="T_1666d_row83_col6" class="data row83 col6" >0</td>
      <td id="T_1666d_row83_col7" class="data row83 col7" >83,012</td>
      <td id="T_1666d_row83_col8" class="data row83 col8" >0</td>
      <td id="T_1666d_row83_col9" class="data row83 col9" >0</td>
      <td id="T_1666d_row83_col10" class="data row83 col10" >74,977</td>
      <td id="T_1666d_row83_col11" class="data row83 col11" >63,995</td>
      <td id="T_1666d_row83_col12" class="data row83 col12" >0</td>
      <td id="T_1666d_row83_col13" class="data row83 col13" >53,507</td>
      <td id="T_1666d_row83_col14" class="data row83 col14" >90,623</td>
      <td id="T_1666d_row83_col15" class="data row83 col15" >91,183</td>
      <td id="T_1666d_row83_col16" class="data row83 col16" >29,077</td>
      <td id="T_1666d_row83_col17" class="data row83 col17" >0</td>
      <td id="T_1666d_row83_col18" class="data row83 col18" >0</td>
      <td id="T_1666d_row83_col19" class="data row83 col19" >0</td>
      <td id="T_1666d_row83_col20" class="data row83 col20" >0</td>
      <td id="T_1666d_row83_col21" class="data row83 col21" >0</td>
      <td id="T_1666d_row83_col22" class="data row83 col22" >141,120</td>
      <td id="T_1666d_row83_col23" class="data row83 col23" >0</td>
      <td id="T_1666d_row83_col24" class="data row83 col24" >0</td>
      <td id="T_1666d_row83_col25" class="data row83 col25" >0</td>
      <td id="T_1666d_row83_col26" class="data row83 col26" >0</td>
      <td id="T_1666d_row83_col27" class="data row83 col27" >118,554</td>
      <td id="T_1666d_row83_col28" class="data row83 col28" >140,522</td>
      <td id="T_1666d_row83_col29" class="data row83 col29" >136,524</td>
      <td id="T_1666d_row83_col30" class="data row83 col30" >1,023,094</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row84" class="row_heading level2 row84" >F</th>
      <td id="T_1666d_row84_col0" class="data row84 col0" >0</td>
      <td id="T_1666d_row84_col1" class="data row84 col1" >0</td>
      <td id="T_1666d_row84_col2" class="data row84 col2" >0</td>
      <td id="T_1666d_row84_col3" class="data row84 col3" >0</td>
      <td id="T_1666d_row84_col4" class="data row84 col4" >0</td>
      <td id="T_1666d_row84_col5" class="data row84 col5" >0</td>
      <td id="T_1666d_row84_col6" class="data row84 col6" >0</td>
      <td id="T_1666d_row84_col7" class="data row84 col7" >0</td>
      <td id="T_1666d_row84_col8" class="data row84 col8" >0</td>
      <td id="T_1666d_row84_col9" class="data row84 col9" >0</td>
      <td id="T_1666d_row84_col10" class="data row84 col10" >0</td>
      <td id="T_1666d_row84_col11" class="data row84 col11" >0</td>
      <td id="T_1666d_row84_col12" class="data row84 col12" >0</td>
      <td id="T_1666d_row84_col13" class="data row84 col13" >0</td>
      <td id="T_1666d_row84_col14" class="data row84 col14" >0</td>
      <td id="T_1666d_row84_col15" class="data row84 col15" >0</td>
      <td id="T_1666d_row84_col16" class="data row84 col16" >0</td>
      <td id="T_1666d_row84_col17" class="data row84 col17" >0</td>
      <td id="T_1666d_row84_col18" class="data row84 col18" >0</td>
      <td id="T_1666d_row84_col19" class="data row84 col19" >0</td>
      <td id="T_1666d_row84_col20" class="data row84 col20" >0</td>
      <td id="T_1666d_row84_col21" class="data row84 col21" >50,922</td>
      <td id="T_1666d_row84_col22" class="data row84 col22" >57,545</td>
      <td id="T_1666d_row84_col23" class="data row84 col23" >0</td>
      <td id="T_1666d_row84_col24" class="data row84 col24" >232,699</td>
      <td id="T_1666d_row84_col25" class="data row84 col25" >0</td>
      <td id="T_1666d_row84_col26" class="data row84 col26" >0</td>
      <td id="T_1666d_row84_col27" class="data row84 col27" >0</td>
      <td id="T_1666d_row84_col28" class="data row84 col28" >0</td>
      <td id="T_1666d_row84_col29" class="data row84 col29" >0</td>
      <td id="T_1666d_row84_col30" class="data row84 col30" >341,166</td>
    </tr>
    <tr>
      <th id="T_1666d_level1_row85" class="row_heading level1 row85" rowspan="8">Sday_2</th>
      <th id="T_1666d_level2_row85" class="row_heading level2 row85" >B</th>
      <td id="T_1666d_row85_col0" class="data row85 col0" >0</td>
      <td id="T_1666d_row85_col1" class="data row85 col1" >0</td>
      <td id="T_1666d_row85_col2" class="data row85 col2" >0</td>
      <td id="T_1666d_row85_col3" class="data row85 col3" >0</td>
      <td id="T_1666d_row85_col4" class="data row85 col4" >0</td>
      <td id="T_1666d_row85_col5" class="data row85 col5" >0</td>
      <td id="T_1666d_row85_col6" class="data row85 col6" >0</td>
      <td id="T_1666d_row85_col7" class="data row85 col7" >0</td>
      <td id="T_1666d_row85_col8" class="data row85 col8" >103,765</td>
      <td id="T_1666d_row85_col9" class="data row85 col9" >0</td>
      <td id="T_1666d_row85_col10" class="data row85 col10" >11,692</td>
      <td id="T_1666d_row85_col11" class="data row85 col11" >22,547</td>
      <td id="T_1666d_row85_col12" class="data row85 col12" >0</td>
      <td id="T_1666d_row85_col13" class="data row85 col13" >0</td>
      <td id="T_1666d_row85_col14" class="data row85 col14" >0</td>
      <td id="T_1666d_row85_col15" class="data row85 col15" >0</td>
      <td id="T_1666d_row85_col16" class="data row85 col16" >166,494</td>
      <td id="T_1666d_row85_col17" class="data row85 col17" >165,998</td>
      <td id="T_1666d_row85_col18" class="data row85 col18" >0</td>
      <td id="T_1666d_row85_col19" class="data row85 col19" >0</td>
      <td id="T_1666d_row85_col20" class="data row85 col20" >0</td>
      <td id="T_1666d_row85_col21" class="data row85 col21" >42,875</td>
      <td id="T_1666d_row85_col22" class="data row85 col22" >0</td>
      <td id="T_1666d_row85_col23" class="data row85 col23" >0</td>
      <td id="T_1666d_row85_col24" class="data row85 col24" >0</td>
      <td id="T_1666d_row85_col25" class="data row85 col25" >0</td>
      <td id="T_1666d_row85_col26" class="data row85 col26" >0</td>
      <td id="T_1666d_row85_col27" class="data row85 col27" >0</td>
      <td id="T_1666d_row85_col28" class="data row85 col28" >0</td>
      <td id="T_1666d_row85_col29" class="data row85 col29" >0</td>
      <td id="T_1666d_row85_col30" class="data row85 col30" >513,371</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row86" class="row_heading level2 row86" >D</th>
      <td id="T_1666d_row86_col0" class="data row86 col0" >2,261</td>
      <td id="T_1666d_row86_col1" class="data row86 col1" >3,865</td>
      <td id="T_1666d_row86_col2" class="data row86 col2" >0</td>
      <td id="T_1666d_row86_col3" class="data row86 col3" >0</td>
      <td id="T_1666d_row86_col4" class="data row86 col4" >0</td>
      <td id="T_1666d_row86_col5" class="data row86 col5" >0</td>
      <td id="T_1666d_row86_col6" class="data row86 col6" >4,390</td>
      <td id="T_1666d_row86_col7" class="data row86 col7" >0</td>
      <td id="T_1666d_row86_col8" class="data row86 col8" >0</td>
      <td id="T_1666d_row86_col9" class="data row86 col9" >0</td>
      <td id="T_1666d_row86_col10" class="data row86 col10" >0</td>
      <td id="T_1666d_row86_col11" class="data row86 col11" >0</td>
      <td id="T_1666d_row86_col12" class="data row86 col12" >0</td>
      <td id="T_1666d_row86_col13" class="data row86 col13" >0</td>
      <td id="T_1666d_row86_col14" class="data row86 col14" >23,059</td>
      <td id="T_1666d_row86_col15" class="data row86 col15" >0</td>
      <td id="T_1666d_row86_col16" class="data row86 col16" >0</td>
      <td id="T_1666d_row86_col17" class="data row86 col17" >500</td>
      <td id="T_1666d_row86_col18" class="data row86 col18" >0</td>
      <td id="T_1666d_row86_col19" class="data row86 col19" >0</td>
      <td id="T_1666d_row86_col20" class="data row86 col20" >0</td>
      <td id="T_1666d_row86_col21" class="data row86 col21" >0</td>
      <td id="T_1666d_row86_col22" class="data row86 col22" >0</td>
      <td id="T_1666d_row86_col23" class="data row86 col23" >4,211</td>
      <td id="T_1666d_row86_col24" class="data row86 col24" >0</td>
      <td id="T_1666d_row86_col25" class="data row86 col25" >0</td>
      <td id="T_1666d_row86_col26" class="data row86 col26" >0</td>
      <td id="T_1666d_row86_col27" class="data row86 col27" >44,530</td>
      <td id="T_1666d_row86_col28" class="data row86 col28" >335</td>
      <td id="T_1666d_row86_col29" class="data row86 col29" >0</td>
      <td id="T_1666d_row86_col30" class="data row86 col30" >83,151</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row87" class="row_heading level2 row87" >G</th>
      <td id="T_1666d_row87_col0" class="data row87 col0" >0</td>
      <td id="T_1666d_row87_col1" class="data row87 col1" >0</td>
      <td id="T_1666d_row87_col2" class="data row87 col2" >0</td>
      <td id="T_1666d_row87_col3" class="data row87 col3" >0</td>
      <td id="T_1666d_row87_col4" class="data row87 col4" >0</td>
      <td id="T_1666d_row87_col5" class="data row87 col5" >0</td>
      <td id="T_1666d_row87_col6" class="data row87 col6" >0</td>
      <td id="T_1666d_row87_col7" class="data row87 col7" >0</td>
      <td id="T_1666d_row87_col8" class="data row87 col8" >0</td>
      <td id="T_1666d_row87_col9" class="data row87 col9" >0</td>
      <td id="T_1666d_row87_col10" class="data row87 col10" >0</td>
      <td id="T_1666d_row87_col11" class="data row87 col11" >0</td>
      <td id="T_1666d_row87_col12" class="data row87 col12" >0</td>
      <td id="T_1666d_row87_col13" class="data row87 col13" >0</td>
      <td id="T_1666d_row87_col14" class="data row87 col14" >0</td>
      <td id="T_1666d_row87_col15" class="data row87 col15" >14,731</td>
      <td id="T_1666d_row87_col16" class="data row87 col16" >2,138</td>
      <td id="T_1666d_row87_col17" class="data row87 col17" >0</td>
      <td id="T_1666d_row87_col18" class="data row87 col18" >0</td>
      <td id="T_1666d_row87_col19" class="data row87 col19" >0</td>
      <td id="T_1666d_row87_col20" class="data row87 col20" >0</td>
      <td id="T_1666d_row87_col21" class="data row87 col21" >42,704</td>
      <td id="T_1666d_row87_col22" class="data row87 col22" >0</td>
      <td id="T_1666d_row87_col23" class="data row87 col23" >0</td>
      <td id="T_1666d_row87_col24" class="data row87 col24" >0</td>
      <td id="T_1666d_row87_col25" class="data row87 col25" >0</td>
      <td id="T_1666d_row87_col26" class="data row87 col26" >0</td>
      <td id="T_1666d_row87_col27" class="data row87 col27" >0</td>
      <td id="T_1666d_row87_col28" class="data row87 col28" >0</td>
      <td id="T_1666d_row87_col29" class="data row87 col29" >3,970</td>
      <td id="T_1666d_row87_col30" class="data row87 col30" >63,543</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row88" class="row_heading level2 row88" >A</th>
      <td id="T_1666d_row88_col0" class="data row88 col0" >241,825</td>
      <td id="T_1666d_row88_col1" class="data row88 col1" >234,607</td>
      <td id="T_1666d_row88_col2" class="data row88 col2" >1,263</td>
      <td id="T_1666d_row88_col3" class="data row88 col3" >112,617</td>
      <td id="T_1666d_row88_col4" class="data row88 col4" >0</td>
      <td id="T_1666d_row88_col5" class="data row88 col5" >0</td>
      <td id="T_1666d_row88_col6" class="data row88 col6" >0</td>
      <td id="T_1666d_row88_col7" class="data row88 col7" >0</td>
      <td id="T_1666d_row88_col8" class="data row88 col8" >0</td>
      <td id="T_1666d_row88_col9" class="data row88 col9" >0</td>
      <td id="T_1666d_row88_col10" class="data row88 col10" >0</td>
      <td id="T_1666d_row88_col11" class="data row88 col11" >0</td>
      <td id="T_1666d_row88_col12" class="data row88 col12" >32,313</td>
      <td id="T_1666d_row88_col13" class="data row88 col13" >0</td>
      <td id="T_1666d_row88_col14" class="data row88 col14" >0</td>
      <td id="T_1666d_row88_col15" class="data row88 col15" >0</td>
      <td id="T_1666d_row88_col16" class="data row88 col16" >0</td>
      <td id="T_1666d_row88_col17" class="data row88 col17" >0</td>
      <td id="T_1666d_row88_col18" class="data row88 col18" >0</td>
      <td id="T_1666d_row88_col19" class="data row88 col19" >0</td>
      <td id="T_1666d_row88_col20" class="data row88 col20" >0</td>
      <td id="T_1666d_row88_col21" class="data row88 col21" >0</td>
      <td id="T_1666d_row88_col22" class="data row88 col22" >50,070</td>
      <td id="T_1666d_row88_col23" class="data row88 col23" >83,616</td>
      <td id="T_1666d_row88_col24" class="data row88 col24" >0</td>
      <td id="T_1666d_row88_col25" class="data row88 col25" >0</td>
      <td id="T_1666d_row88_col26" class="data row88 col26" >0</td>
      <td id="T_1666d_row88_col27" class="data row88 col27" >0</td>
      <td id="T_1666d_row88_col28" class="data row88 col28" >0</td>
      <td id="T_1666d_row88_col29" class="data row88 col29" >0</td>
      <td id="T_1666d_row88_col30" class="data row88 col30" >756,311</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row89" class="row_heading level2 row89" >I</th>
      <td id="T_1666d_row89_col0" class="data row89 col0" >0</td>
      <td id="T_1666d_row89_col1" class="data row89 col1" >0</td>
      <td id="T_1666d_row89_col2" class="data row89 col2" >334,315</td>
      <td id="T_1666d_row89_col3" class="data row89 col3" >185,844</td>
      <td id="T_1666d_row89_col4" class="data row89 col4" >0</td>
      <td id="T_1666d_row89_col5" class="data row89 col5" >0</td>
      <td id="T_1666d_row89_col6" class="data row89 col6" >309,660</td>
      <td id="T_1666d_row89_col7" class="data row89 col7" >0</td>
      <td id="T_1666d_row89_col8" class="data row89 col8" >0</td>
      <td id="T_1666d_row89_col9" class="data row89 col9" >136,534</td>
      <td id="T_1666d_row89_col10" class="data row89 col10" >95,278</td>
      <td id="T_1666d_row89_col11" class="data row89 col11" >185,572</td>
      <td id="T_1666d_row89_col12" class="data row89 col12" >0</td>
      <td id="T_1666d_row89_col13" class="data row89 col13" >110,401</td>
      <td id="T_1666d_row89_col14" class="data row89 col14" >0</td>
      <td id="T_1666d_row89_col15" class="data row89 col15" >0</td>
      <td id="T_1666d_row89_col16" class="data row89 col16" >9,645</td>
      <td id="T_1666d_row89_col17" class="data row89 col17" >3,586</td>
      <td id="T_1666d_row89_col18" class="data row89 col18" >0</td>
      <td id="T_1666d_row89_col19" class="data row89 col19" >0</td>
      <td id="T_1666d_row89_col20" class="data row89 col20" >235,200</td>
      <td id="T_1666d_row89_col21" class="data row89 col21" >55,938</td>
      <td id="T_1666d_row89_col22" class="data row89 col22" >168,441</td>
      <td id="T_1666d_row89_col23" class="data row89 col23" >0</td>
      <td id="T_1666d_row89_col24" class="data row89 col24" >499</td>
      <td id="T_1666d_row89_col25" class="data row89 col25" >0</td>
      <td id="T_1666d_row89_col26" class="data row89 col26" >0</td>
      <td id="T_1666d_row89_col27" class="data row89 col27" >0</td>
      <td id="T_1666d_row89_col28" class="data row89 col28" >2,483</td>
      <td id="T_1666d_row89_col29" class="data row89 col29" >0</td>
      <td id="T_1666d_row89_col30" class="data row89 col30" >1,833,396</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row90" class="row_heading level2 row90" >E</th>
      <td id="T_1666d_row90_col0" class="data row90 col0" >0</td>
      <td id="T_1666d_row90_col1" class="data row90 col1" >0</td>
      <td id="T_1666d_row90_col2" class="data row90 col2" >0</td>
      <td id="T_1666d_row90_col3" class="data row90 col3" >0</td>
      <td id="T_1666d_row90_col4" class="data row90 col4" >0</td>
      <td id="T_1666d_row90_col5" class="data row90 col5" >0</td>
      <td id="T_1666d_row90_col6" class="data row90 col6" >0</td>
      <td id="T_1666d_row90_col7" class="data row90 col7" >0</td>
      <td id="T_1666d_row90_col8" class="data row90 col8" >0</td>
      <td id="T_1666d_row90_col9" class="data row90 col9" >0</td>
      <td id="T_1666d_row90_col10" class="data row90 col10" >0</td>
      <td id="T_1666d_row90_col11" class="data row90 col11" >0</td>
      <td id="T_1666d_row90_col12" class="data row90 col12" >0</td>
      <td id="T_1666d_row90_col13" class="data row90 col13" >0</td>
      <td id="T_1666d_row90_col14" class="data row90 col14" >0</td>
      <td id="T_1666d_row90_col15" class="data row90 col15" >0</td>
      <td id="T_1666d_row90_col16" class="data row90 col16" >0</td>
      <td id="T_1666d_row90_col17" class="data row90 col17" >0</td>
      <td id="T_1666d_row90_col18" class="data row90 col18" >0</td>
      <td id="T_1666d_row90_col19" class="data row90 col19" >0</td>
      <td id="T_1666d_row90_col20" class="data row90 col20" >0</td>
      <td id="T_1666d_row90_col21" class="data row90 col21" >0</td>
      <td id="T_1666d_row90_col22" class="data row90 col22" >0</td>
      <td id="T_1666d_row90_col23" class="data row90 col23" >0</td>
      <td id="T_1666d_row90_col24" class="data row90 col24" >0</td>
      <td id="T_1666d_row90_col25" class="data row90 col25" >0</td>
      <td id="T_1666d_row90_col26" class="data row90 col26" >0</td>
      <td id="T_1666d_row90_col27" class="data row90 col27" >2</td>
      <td id="T_1666d_row90_col28" class="data row90 col28" >0</td>
      <td id="T_1666d_row90_col29" class="data row90 col29" >0</td>
      <td id="T_1666d_row90_col30" class="data row90 col30" >2</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row91" class="row_heading level2 row91" >C</th>
      <td id="T_1666d_row91_col0" class="data row91 col0" >0</td>
      <td id="T_1666d_row91_col1" class="data row91 col1" >0</td>
      <td id="T_1666d_row91_col2" class="data row91 col2" >0</td>
      <td id="T_1666d_row91_col3" class="data row91 col3" >0</td>
      <td id="T_1666d_row91_col4" class="data row91 col4" >0</td>
      <td id="T_1666d_row91_col5" class="data row91 col5" >0</td>
      <td id="T_1666d_row91_col6" class="data row91 col6" >0</td>
      <td id="T_1666d_row91_col7" class="data row91 col7" >83,012</td>
      <td id="T_1666d_row91_col8" class="data row91 col8" >0</td>
      <td id="T_1666d_row91_col9" class="data row91 col9" >0</td>
      <td id="T_1666d_row91_col10" class="data row91 col10" >83,553</td>
      <td id="T_1666d_row91_col11" class="data row91 col11" >0</td>
      <td id="T_1666d_row91_col12" class="data row91 col12" >0</td>
      <td id="T_1666d_row91_col13" class="data row91 col13" >65,915</td>
      <td id="T_1666d_row91_col14" class="data row91 col14" >0</td>
      <td id="T_1666d_row91_col15" class="data row91 col15" >67,760</td>
      <td id="T_1666d_row91_col16" class="data row91 col16" >0</td>
      <td id="T_1666d_row91_col17" class="data row91 col17" >4,370</td>
      <td id="T_1666d_row91_col18" class="data row91 col18" >0</td>
      <td id="T_1666d_row91_col19" class="data row91 col19" >0</td>
      <td id="T_1666d_row91_col20" class="data row91 col20" >0</td>
      <td id="T_1666d_row91_col21" class="data row91 col21" >0</td>
      <td id="T_1666d_row91_col22" class="data row91 col22" >0</td>
      <td id="T_1666d_row91_col23" class="data row91 col23" >960</td>
      <td id="T_1666d_row91_col24" class="data row91 col24" >0</td>
      <td id="T_1666d_row91_col25" class="data row91 col25" >0</td>
      <td id="T_1666d_row91_col26" class="data row91 col26" >0</td>
      <td id="T_1666d_row91_col27" class="data row91 col27" >41,291</td>
      <td id="T_1666d_row91_col28" class="data row91 col28" >198,598</td>
      <td id="T_1666d_row91_col29" class="data row91 col29" >136,850</td>
      <td id="T_1666d_row91_col30" class="data row91 col30" >682,309</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row92" class="row_heading level2 row92" >F</th>
      <td id="T_1666d_row92_col0" class="data row92 col0" >0</td>
      <td id="T_1666d_row92_col1" class="data row92 col1" >0</td>
      <td id="T_1666d_row92_col2" class="data row92 col2" >0</td>
      <td id="T_1666d_row92_col3" class="data row92 col3" >0</td>
      <td id="T_1666d_row92_col4" class="data row92 col4" >0</td>
      <td id="T_1666d_row92_col5" class="data row92 col5" >0</td>
      <td id="T_1666d_row92_col6" class="data row92 col6" >0</td>
      <td id="T_1666d_row92_col7" class="data row92 col7" >0</td>
      <td id="T_1666d_row92_col8" class="data row92 col8" >0</td>
      <td id="T_1666d_row92_col9" class="data row92 col9" >0</td>
      <td id="T_1666d_row92_col10" class="data row92 col10" >0</td>
      <td id="T_1666d_row92_col11" class="data row92 col11" >0</td>
      <td id="T_1666d_row92_col12" class="data row92 col12" >0</td>
      <td id="T_1666d_row92_col13" class="data row92 col13" >0</td>
      <td id="T_1666d_row92_col14" class="data row92 col14" >0</td>
      <td id="T_1666d_row92_col15" class="data row92 col15" >97,715</td>
      <td id="T_1666d_row92_col16" class="data row92 col16" >0</td>
      <td id="T_1666d_row92_col17" class="data row92 col17" >0</td>
      <td id="T_1666d_row92_col18" class="data row92 col18" >0</td>
      <td id="T_1666d_row92_col19" class="data row92 col19" >0</td>
      <td id="T_1666d_row92_col20" class="data row92 col20" >0</td>
      <td id="T_1666d_row92_col21" class="data row92 col21" >50,922</td>
      <td id="T_1666d_row92_col22" class="data row92 col22" >0</td>
      <td id="T_1666d_row92_col23" class="data row92 col23" >0</td>
      <td id="T_1666d_row92_col24" class="data row92 col24" >234,701</td>
      <td id="T_1666d_row92_col25" class="data row92 col25" >0</td>
      <td id="T_1666d_row92_col26" class="data row92 col26" >0</td>
      <td id="T_1666d_row92_col27" class="data row92 col27" >0</td>
      <td id="T_1666d_row92_col28" class="data row92 col28" >510</td>
      <td id="T_1666d_row92_col29" class="data row92 col29" >101,300</td>
      <td id="T_1666d_row92_col30" class="data row92 col30" >485,148</td>
    </tr>
    <tr>
      <th id="T_1666d_level0_row93" class="row_heading level0 row93" rowspan="18">QaC3</th>
      <th id="T_1666d_level1_row93" class="row_heading level1 row93" rowspan="9">Sday_1</th>
      <th id="T_1666d_level2_row93" class="row_heading level2 row93" >B</th>
      <td id="T_1666d_row93_col0" class="data row93 col0" >0</td>
      <td id="T_1666d_row93_col1" class="data row93 col1" >0</td>
      <td id="T_1666d_row93_col2" class="data row93 col2" >0</td>
      <td id="T_1666d_row93_col3" class="data row93 col3" >0</td>
      <td id="T_1666d_row93_col4" class="data row93 col4" >0</td>
      <td id="T_1666d_row93_col5" class="data row93 col5" >0</td>
      <td id="T_1666d_row93_col6" class="data row93 col6" >0</td>
      <td id="T_1666d_row93_col7" class="data row93 col7" >4,468</td>
      <td id="T_1666d_row93_col8" class="data row93 col8" >103,765</td>
      <td id="T_1666d_row93_col9" class="data row93 col9" >0</td>
      <td id="T_1666d_row93_col10" class="data row93 col10" >0</td>
      <td id="T_1666d_row93_col11" class="data row93 col11" >0</td>
      <td id="T_1666d_row93_col12" class="data row93 col12" >0</td>
      <td id="T_1666d_row93_col13" class="data row93 col13" >0</td>
      <td id="T_1666d_row93_col14" class="data row93 col14" >120,096</td>
      <td id="T_1666d_row93_col15" class="data row93 col15" >0</td>
      <td id="T_1666d_row93_col16" class="data row93 col16" >0</td>
      <td id="T_1666d_row93_col17" class="data row93 col17" >38,903</td>
      <td id="T_1666d_row93_col18" class="data row93 col18" >0</td>
      <td id="T_1666d_row93_col19" class="data row93 col19" >0</td>
      <td id="T_1666d_row93_col20" class="data row93 col20" >126,000</td>
      <td id="T_1666d_row93_col21" class="data row93 col21" >105,177</td>
      <td id="T_1666d_row93_col22" class="data row93 col22" >0</td>
      <td id="T_1666d_row93_col23" class="data row93 col23" >0</td>
      <td id="T_1666d_row93_col24" class="data row93 col24" >0</td>
      <td id="T_1666d_row93_col25" class="data row93 col25" >0</td>
      <td id="T_1666d_row93_col26" class="data row93 col26" >0</td>
      <td id="T_1666d_row93_col27" class="data row93 col27" >0</td>
      <td id="T_1666d_row93_col28" class="data row93 col28" >0</td>
      <td id="T_1666d_row93_col29" class="data row93 col29" >99,509</td>
      <td id="T_1666d_row93_col30" class="data row93 col30" >597,918</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row94" class="row_heading level2 row94" >H</th>
      <td id="T_1666d_row94_col0" class="data row94 col0" >0</td>
      <td id="T_1666d_row94_col1" class="data row94 col1" >0</td>
      <td id="T_1666d_row94_col2" class="data row94 col2" >0</td>
      <td id="T_1666d_row94_col3" class="data row94 col3" >0</td>
      <td id="T_1666d_row94_col4" class="data row94 col4" >0</td>
      <td id="T_1666d_row94_col5" class="data row94 col5" >0</td>
      <td id="T_1666d_row94_col6" class="data row94 col6" >0</td>
      <td id="T_1666d_row94_col7" class="data row94 col7" >0</td>
      <td id="T_1666d_row94_col8" class="data row94 col8" >0</td>
      <td id="T_1666d_row94_col9" class="data row94 col9" >0</td>
      <td id="T_1666d_row94_col10" class="data row94 col10" >0</td>
      <td id="T_1666d_row94_col11" class="data row94 col11" >0</td>
      <td id="T_1666d_row94_col12" class="data row94 col12" >0</td>
      <td id="T_1666d_row94_col13" class="data row94 col13" >0</td>
      <td id="T_1666d_row94_col14" class="data row94 col14" >0</td>
      <td id="T_1666d_row94_col15" class="data row94 col15" >2,208</td>
      <td id="T_1666d_row94_col16" class="data row94 col16" >0</td>
      <td id="T_1666d_row94_col17" class="data row94 col17" >0</td>
      <td id="T_1666d_row94_col18" class="data row94 col18" >0</td>
      <td id="T_1666d_row94_col19" class="data row94 col19" >0</td>
      <td id="T_1666d_row94_col20" class="data row94 col20" >0</td>
      <td id="T_1666d_row94_col21" class="data row94 col21" >0</td>
      <td id="T_1666d_row94_col22" class="data row94 col22" >0</td>
      <td id="T_1666d_row94_col23" class="data row94 col23" >0</td>
      <td id="T_1666d_row94_col24" class="data row94 col24" >4,112</td>
      <td id="T_1666d_row94_col25" class="data row94 col25" >0</td>
      <td id="T_1666d_row94_col26" class="data row94 col26" >0</td>
      <td id="T_1666d_row94_col27" class="data row94 col27" >0</td>
      <td id="T_1666d_row94_col28" class="data row94 col28" >0</td>
      <td id="T_1666d_row94_col29" class="data row94 col29" >0</td>
      <td id="T_1666d_row94_col30" class="data row94 col30" >6,320</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row95" class="row_heading level2 row95" >D</th>
      <td id="T_1666d_row95_col0" class="data row95 col0" >4,814</td>
      <td id="T_1666d_row95_col1" class="data row95 col1" >0</td>
      <td id="T_1666d_row95_col2" class="data row95 col2" >0</td>
      <td id="T_1666d_row95_col3" class="data row95 col3" >11,243</td>
      <td id="T_1666d_row95_col4" class="data row95 col4" >0</td>
      <td id="T_1666d_row95_col5" class="data row95 col5" >0</td>
      <td id="T_1666d_row95_col6" class="data row95 col6" >7,109</td>
      <td id="T_1666d_row95_col7" class="data row95 col7" >0</td>
      <td id="T_1666d_row95_col8" class="data row95 col8" >0</td>
      <td id="T_1666d_row95_col9" class="data row95 col9" >3,576</td>
      <td id="T_1666d_row95_col10" class="data row95 col10" >0</td>
      <td id="T_1666d_row95_col11" class="data row95 col11" >0</td>
      <td id="T_1666d_row95_col12" class="data row95 col12" >0</td>
      <td id="T_1666d_row95_col13" class="data row95 col13" >0</td>
      <td id="T_1666d_row95_col14" class="data row95 col14" >0</td>
      <td id="T_1666d_row95_col15" class="data row95 col15" >2,192</td>
      <td id="T_1666d_row95_col16" class="data row95 col16" >0</td>
      <td id="T_1666d_row95_col17" class="data row95 col17" >10,700</td>
      <td id="T_1666d_row95_col18" class="data row95 col18" >0</td>
      <td id="T_1666d_row95_col19" class="data row95 col19" >0</td>
      <td id="T_1666d_row95_col20" class="data row95 col20" >11,200</td>
      <td id="T_1666d_row95_col21" class="data row95 col21" >4,920</td>
      <td id="T_1666d_row95_col22" class="data row95 col22" >0</td>
      <td id="T_1666d_row95_col23" class="data row95 col23" >0</td>
      <td id="T_1666d_row95_col24" class="data row95 col24" >0</td>
      <td id="T_1666d_row95_col25" class="data row95 col25" >0</td>
      <td id="T_1666d_row95_col26" class="data row95 col26" >0</td>
      <td id="T_1666d_row95_col27" class="data row95 col27" >0</td>
      <td id="T_1666d_row95_col28" class="data row95 col28" >23,059</td>
      <td id="T_1666d_row95_col29" class="data row95 col29" >0</td>
      <td id="T_1666d_row95_col30" class="data row95 col30" >78,813</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row96" class="row_heading level2 row96" >G</th>
      <td id="T_1666d_row96_col0" class="data row96 col0" >0</td>
      <td id="T_1666d_row96_col1" class="data row96 col1" >0</td>
      <td id="T_1666d_row96_col2" class="data row96 col2" >0</td>
      <td id="T_1666d_row96_col3" class="data row96 col3" >0</td>
      <td id="T_1666d_row96_col4" class="data row96 col4" >0</td>
      <td id="T_1666d_row96_col5" class="data row96 col5" >0</td>
      <td id="T_1666d_row96_col6" class="data row96 col6" >0</td>
      <td id="T_1666d_row96_col7" class="data row96 col7" >60,596</td>
      <td id="T_1666d_row96_col8" class="data row96 col8" >0</td>
      <td id="T_1666d_row96_col9" class="data row96 col9" >0</td>
      <td id="T_1666d_row96_col10" class="data row96 col10" >499</td>
      <td id="T_1666d_row96_col11" class="data row96 col11" >0</td>
      <td id="T_1666d_row96_col12" class="data row96 col12" >0</td>
      <td id="T_1666d_row96_col13" class="data row96 col13" >0</td>
      <td id="T_1666d_row96_col14" class="data row96 col14" >0</td>
      <td id="T_1666d_row96_col15" class="data row96 col15" >0</td>
      <td id="T_1666d_row96_col16" class="data row96 col16" >0</td>
      <td id="T_1666d_row96_col17" class="data row96 col17" >0</td>
      <td id="T_1666d_row96_col18" class="data row96 col18" >0</td>
      <td id="T_1666d_row96_col19" class="data row96 col19" >0</td>
      <td id="T_1666d_row96_col20" class="data row96 col20" >0</td>
      <td id="T_1666d_row96_col21" class="data row96 col21" >0</td>
      <td id="T_1666d_row96_col22" class="data row96 col22" >0</td>
      <td id="T_1666d_row96_col23" class="data row96 col23" >0</td>
      <td id="T_1666d_row96_col24" class="data row96 col24" >37,118</td>
      <td id="T_1666d_row96_col25" class="data row96 col25" >0</td>
      <td id="T_1666d_row96_col26" class="data row96 col26" >0</td>
      <td id="T_1666d_row96_col27" class="data row96 col27" >0</td>
      <td id="T_1666d_row96_col28" class="data row96 col28" >0</td>
      <td id="T_1666d_row96_col29" class="data row96 col29" >0</td>
      <td id="T_1666d_row96_col30" class="data row96 col30" >98,213</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row97" class="row_heading level2 row97" >A</th>
      <td id="T_1666d_row97_col0" class="data row97 col0" >170,015</td>
      <td id="T_1666d_row97_col1" class="data row97 col1" >75,872</td>
      <td id="T_1666d_row97_col2" class="data row97 col2" >252,000</td>
      <td id="T_1666d_row97_col3" class="data row97 col3" >0</td>
      <td id="T_1666d_row97_col4" class="data row97 col4" >0</td>
      <td id="T_1666d_row97_col5" class="data row97 col5" >0</td>
      <td id="T_1666d_row97_col6" class="data row97 col6" >0</td>
      <td id="T_1666d_row97_col7" class="data row97 col7" >0</td>
      <td id="T_1666d_row97_col8" class="data row97 col8" >0</td>
      <td id="T_1666d_row97_col9" class="data row97 col9" >26,149</td>
      <td id="T_1666d_row97_col10" class="data row97 col10" >0</td>
      <td id="T_1666d_row97_col11" class="data row97 col11" >0</td>
      <td id="T_1666d_row97_col12" class="data row97 col12" >0</td>
      <td id="T_1666d_row97_col13" class="data row97 col13" >9,105</td>
      <td id="T_1666d_row97_col14" class="data row97 col14" >0</td>
      <td id="T_1666d_row97_col15" class="data row97 col15" >10,251</td>
      <td id="T_1666d_row97_col16" class="data row97 col16" >0</td>
      <td id="T_1666d_row97_col17" class="data row97 col17" >88,972</td>
      <td id="T_1666d_row97_col18" class="data row97 col18" >0</td>
      <td id="T_1666d_row97_col19" class="data row97 col19" >0</td>
      <td id="T_1666d_row97_col20" class="data row97 col20" >0</td>
      <td id="T_1666d_row97_col21" class="data row97 col21" >0</td>
      <td id="T_1666d_row97_col22" class="data row97 col22" >0</td>
      <td id="T_1666d_row97_col23" class="data row97 col23" >0</td>
      <td id="T_1666d_row97_col24" class="data row97 col24" >0</td>
      <td id="T_1666d_row97_col25" class="data row97 col25" >0</td>
      <td id="T_1666d_row97_col26" class="data row97 col26" >0</td>
      <td id="T_1666d_row97_col27" class="data row97 col27" >0</td>
      <td id="T_1666d_row97_col28" class="data row97 col28" >0</td>
      <td id="T_1666d_row97_col29" class="data row97 col29" >0</td>
      <td id="T_1666d_row97_col30" class="data row97 col30" >632,364</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row98" class="row_heading level2 row98" >I</th>
      <td id="T_1666d_row98_col0" class="data row98 col0" >0</td>
      <td id="T_1666d_row98_col1" class="data row98 col1" >63,042</td>
      <td id="T_1666d_row98_col2" class="data row98 col2" >0</td>
      <td id="T_1666d_row98_col3" class="data row98 col3" >267,078</td>
      <td id="T_1666d_row98_col4" class="data row98 col4" >0</td>
      <td id="T_1666d_row98_col5" class="data row98 col5" >0</td>
      <td id="T_1666d_row98_col6" class="data row98 col6" >0</td>
      <td id="T_1666d_row98_col7" class="data row98 col7" >133,568</td>
      <td id="T_1666d_row98_col8" class="data row98 col8" >0</td>
      <td id="T_1666d_row98_col9" class="data row98 col9" >142,040</td>
      <td id="T_1666d_row98_col10" class="data row98 col10" >138,346</td>
      <td id="T_1666d_row98_col11" class="data row98 col11" >14,682</td>
      <td id="T_1666d_row98_col12" class="data row98 col12" >235,200</td>
      <td id="T_1666d_row98_col13" class="data row98 col13" >223,060</td>
      <td id="T_1666d_row98_col14" class="data row98 col14" >73,925</td>
      <td id="T_1666d_row98_col15" class="data row98 col15" >0</td>
      <td id="T_1666d_row98_col16" class="data row98 col16" >235,198</td>
      <td id="T_1666d_row98_col17" class="data row98 col17" >500</td>
      <td id="T_1666d_row98_col18" class="data row98 col18" >0</td>
      <td id="T_1666d_row98_col19" class="data row98 col19" >0</td>
      <td id="T_1666d_row98_col20" class="data row98 col20" >0</td>
      <td id="T_1666d_row98_col21" class="data row98 col21" >65,444</td>
      <td id="T_1666d_row98_col22" class="data row98 col22" >1,999</td>
      <td id="T_1666d_row98_col23" class="data row98 col23" >138,353</td>
      <td id="T_1666d_row98_col24" class="data row98 col24" >0</td>
      <td id="T_1666d_row98_col25" class="data row98 col25" >0</td>
      <td id="T_1666d_row98_col26" class="data row98 col26" >0</td>
      <td id="T_1666d_row98_col27" class="data row98 col27" >0</td>
      <td id="T_1666d_row98_col28" class="data row98 col28" >0</td>
      <td id="T_1666d_row98_col29" class="data row98 col29" >0</td>
      <td id="T_1666d_row98_col30" class="data row98 col30" >1,732,435</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row99" class="row_heading level2 row99" >E</th>
      <td id="T_1666d_row99_col0" class="data row99 col0" >0</td>
      <td id="T_1666d_row99_col1" class="data row99 col1" >0</td>
      <td id="T_1666d_row99_col2" class="data row99 col2" >0</td>
      <td id="T_1666d_row99_col3" class="data row99 col3" >0</td>
      <td id="T_1666d_row99_col4" class="data row99 col4" >0</td>
      <td id="T_1666d_row99_col5" class="data row99 col5" >0</td>
      <td id="T_1666d_row99_col6" class="data row99 col6" >0</td>
      <td id="T_1666d_row99_col7" class="data row99 col7" >0</td>
      <td id="T_1666d_row99_col8" class="data row99 col8" >0</td>
      <td id="T_1666d_row99_col9" class="data row99 col9" >0</td>
      <td id="T_1666d_row99_col10" class="data row99 col10" >0</td>
      <td id="T_1666d_row99_col11" class="data row99 col11" >0</td>
      <td id="T_1666d_row99_col12" class="data row99 col12" >0</td>
      <td id="T_1666d_row99_col13" class="data row99 col13" >0</td>
      <td id="T_1666d_row99_col14" class="data row99 col14" >0</td>
      <td id="T_1666d_row99_col15" class="data row99 col15" >0</td>
      <td id="T_1666d_row99_col16" class="data row99 col16" >0</td>
      <td id="T_1666d_row99_col17" class="data row99 col17" >0</td>
      <td id="T_1666d_row99_col18" class="data row99 col18" >0</td>
      <td id="T_1666d_row99_col19" class="data row99 col19" >0</td>
      <td id="T_1666d_row99_col20" class="data row99 col20" >0</td>
      <td id="T_1666d_row99_col21" class="data row99 col21" >0</td>
      <td id="T_1666d_row99_col22" class="data row99 col22" >0</td>
      <td id="T_1666d_row99_col23" class="data row99 col23" >0</td>
      <td id="T_1666d_row99_col24" class="data row99 col24" >1</td>
      <td id="T_1666d_row99_col25" class="data row99 col25" >0</td>
      <td id="T_1666d_row99_col26" class="data row99 col26" >0</td>
      <td id="T_1666d_row99_col27" class="data row99 col27" >0</td>
      <td id="T_1666d_row99_col28" class="data row99 col28" >0</td>
      <td id="T_1666d_row99_col29" class="data row99 col29" >0</td>
      <td id="T_1666d_row99_col30" class="data row99 col30" >1</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row100" class="row_heading level2 row100" >C</th>
      <td id="T_1666d_row100_col0" class="data row100 col0" >0</td>
      <td id="T_1666d_row100_col1" class="data row100 col1" >0</td>
      <td id="T_1666d_row100_col2" class="data row100 col2" >0</td>
      <td id="T_1666d_row100_col3" class="data row100 col3" >0</td>
      <td id="T_1666d_row100_col4" class="data row100 col4" >0</td>
      <td id="T_1666d_row100_col5" class="data row100 col5" >0</td>
      <td id="T_1666d_row100_col6" class="data row100 col6" >176,006</td>
      <td id="T_1666d_row100_col7" class="data row100 col7" >56,891</td>
      <td id="T_1666d_row100_col8" class="data row100 col8" >0</td>
      <td id="T_1666d_row100_col9" class="data row100 col9" >22,103</td>
      <td id="T_1666d_row100_col10" class="data row100 col10" >0</td>
      <td id="T_1666d_row100_col11" class="data row100 col11" >132,311</td>
      <td id="T_1666d_row100_col12" class="data row100 col12" >0</td>
      <td id="T_1666d_row100_col13" class="data row100 col13" >0</td>
      <td id="T_1666d_row100_col14" class="data row100 col14" >688</td>
      <td id="T_1666d_row100_col15" class="data row100 col15" >59,854</td>
      <td id="T_1666d_row100_col16" class="data row100 col16" >0</td>
      <td id="T_1666d_row100_col17" class="data row100 col17" >0</td>
      <td id="T_1666d_row100_col18" class="data row100 col18" >0</td>
      <td id="T_1666d_row100_col19" class="data row100 col19" >0</td>
      <td id="T_1666d_row100_col20" class="data row100 col20" >0</td>
      <td id="T_1666d_row100_col21" class="data row100 col21" >0</td>
      <td id="T_1666d_row100_col22" class="data row100 col22" >164,770</td>
      <td id="T_1666d_row100_col23" class="data row100 col23" >0</td>
      <td id="T_1666d_row100_col24" class="data row100 col24" >32,735</td>
      <td id="T_1666d_row100_col25" class="data row100 col25" >0</td>
      <td id="T_1666d_row100_col26" class="data row100 col26" >0</td>
      <td id="T_1666d_row100_col27" class="data row100 col27" >201,600</td>
      <td id="T_1666d_row100_col28" class="data row100 col28" >0</td>
      <td id="T_1666d_row100_col29" class="data row100 col29" >0</td>
      <td id="T_1666d_row100_col30" class="data row100 col30" >846,958</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row101" class="row_heading level2 row101" >F</th>
      <td id="T_1666d_row101_col0" class="data row101 col0" >0</td>
      <td id="T_1666d_row101_col1" class="data row101 col1" >70,996</td>
      <td id="T_1666d_row101_col2" class="data row101 col2" >0</td>
      <td id="T_1666d_row101_col3" class="data row101 col3" >0</td>
      <td id="T_1666d_row101_col4" class="data row101 col4" >0</td>
      <td id="T_1666d_row101_col5" class="data row101 col5" >0</td>
      <td id="T_1666d_row101_col6" class="data row101 col6" >0</td>
      <td id="T_1666d_row101_col7" class="data row101 col7" >0</td>
      <td id="T_1666d_row101_col8" class="data row101 col8" >0</td>
      <td id="T_1666d_row101_col9" class="data row101 col9" >0</td>
      <td id="T_1666d_row101_col10" class="data row101 col10" >6</td>
      <td id="T_1666d_row101_col11" class="data row101 col11" >0</td>
      <td id="T_1666d_row101_col12" class="data row101 col12" >0</td>
      <td id="T_1666d_row101_col13" class="data row101 col13" >0</td>
      <td id="T_1666d_row101_col14" class="data row101 col14" >0</td>
      <td id="T_1666d_row101_col15" class="data row101 col15" >0</td>
      <td id="T_1666d_row101_col16" class="data row101 col16" >2</td>
      <td id="T_1666d_row101_col17" class="data row101 col17" >0</td>
      <td id="T_1666d_row101_col18" class="data row101 col18" >0</td>
      <td id="T_1666d_row101_col19" class="data row101 col19" >0</td>
      <td id="T_1666d_row101_col20" class="data row101 col20" >0</td>
      <td id="T_1666d_row101_col21" class="data row101 col21" >0</td>
      <td id="T_1666d_row101_col22" class="data row101 col22" >0</td>
      <td id="T_1666d_row101_col23" class="data row101 col23" >0</td>
      <td id="T_1666d_row101_col24" class="data row101 col24" >0</td>
      <td id="T_1666d_row101_col25" class="data row101 col25" >0</td>
      <td id="T_1666d_row101_col26" class="data row101 col26" >0</td>
      <td id="T_1666d_row101_col27" class="data row101 col27" >0</td>
      <td id="T_1666d_row101_col28" class="data row101 col28" >0</td>
      <td id="T_1666d_row101_col29" class="data row101 col29" >102,522</td>
      <td id="T_1666d_row101_col30" class="data row101 col30" >173,526</td>
    </tr>
    <tr>
      <th id="T_1666d_level1_row102" class="row_heading level1 row102" rowspan="9">Sday_2</th>
      <th id="T_1666d_level2_row102" class="row_heading level2 row102" >B</th>
      <td id="T_1666d_row102_col0" class="data row102 col0" >22,917</td>
      <td id="T_1666d_row102_col1" class="data row102 col1" >0</td>
      <td id="T_1666d_row102_col2" class="data row102 col2" >0</td>
      <td id="T_1666d_row102_col3" class="data row102 col3" >0</td>
      <td id="T_1666d_row102_col4" class="data row102 col4" >0</td>
      <td id="T_1666d_row102_col5" class="data row102 col5" >0</td>
      <td id="T_1666d_row102_col6" class="data row102 col6" >169,164</td>
      <td id="T_1666d_row102_col7" class="data row102 col7" >0</td>
      <td id="T_1666d_row102_col8" class="data row102 col8" >103,765</td>
      <td id="T_1666d_row102_col9" class="data row102 col9" >0</td>
      <td id="T_1666d_row102_col10" class="data row102 col10" >0</td>
      <td id="T_1666d_row102_col11" class="data row102 col11" >0</td>
      <td id="T_1666d_row102_col12" class="data row102 col12" >0</td>
      <td id="T_1666d_row102_col13" class="data row102 col13" >0</td>
      <td id="T_1666d_row102_col14" class="data row102 col14" >134,503</td>
      <td id="T_1666d_row102_col15" class="data row102 col15" >0</td>
      <td id="T_1666d_row102_col16" class="data row102 col16" >0</td>
      <td id="T_1666d_row102_col17" class="data row102 col17" >8,153</td>
      <td id="T_1666d_row102_col18" class="data row102 col18" >0</td>
      <td id="T_1666d_row102_col19" class="data row102 col19" >0</td>
      <td id="T_1666d_row102_col20" class="data row102 col20" >53,365</td>
      <td id="T_1666d_row102_col21" class="data row102 col21" >14,116</td>
      <td id="T_1666d_row102_col22" class="data row102 col22" >0</td>
      <td id="T_1666d_row102_col23" class="data row102 col23" >0</td>
      <td id="T_1666d_row102_col24" class="data row102 col24" >0</td>
      <td id="T_1666d_row102_col25" class="data row102 col25" >0</td>
      <td id="T_1666d_row102_col26" class="data row102 col26" >0</td>
      <td id="T_1666d_row102_col27" class="data row102 col27" >0</td>
      <td id="T_1666d_row102_col28" class="data row102 col28" >0</td>
      <td id="T_1666d_row102_col29" class="data row102 col29" >0</td>
      <td id="T_1666d_row102_col30" class="data row102 col30" >505,983</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row103" class="row_heading level2 row103" >H</th>
      <td id="T_1666d_row103_col0" class="data row103 col0" >0</td>
      <td id="T_1666d_row103_col1" class="data row103 col1" >0</td>
      <td id="T_1666d_row103_col2" class="data row103 col2" >0</td>
      <td id="T_1666d_row103_col3" class="data row103 col3" >0</td>
      <td id="T_1666d_row103_col4" class="data row103 col4" >0</td>
      <td id="T_1666d_row103_col5" class="data row103 col5" >0</td>
      <td id="T_1666d_row103_col6" class="data row103 col6" >0</td>
      <td id="T_1666d_row103_col7" class="data row103 col7" >0</td>
      <td id="T_1666d_row103_col8" class="data row103 col8" >0</td>
      <td id="T_1666d_row103_col9" class="data row103 col9" >0</td>
      <td id="T_1666d_row103_col10" class="data row103 col10" >0</td>
      <td id="T_1666d_row103_col11" class="data row103 col11" >0</td>
      <td id="T_1666d_row103_col12" class="data row103 col12" >0</td>
      <td id="T_1666d_row103_col13" class="data row103 col13" >0</td>
      <td id="T_1666d_row103_col14" class="data row103 col14" >0</td>
      <td id="T_1666d_row103_col15" class="data row103 col15" >0</td>
      <td id="T_1666d_row103_col16" class="data row103 col16" >0</td>
      <td id="T_1666d_row103_col17" class="data row103 col17" >0</td>
      <td id="T_1666d_row103_col18" class="data row103 col18" >0</td>
      <td id="T_1666d_row103_col19" class="data row103 col19" >0</td>
      <td id="T_1666d_row103_col20" class="data row103 col20" >0</td>
      <td id="T_1666d_row103_col21" class="data row103 col21" >0</td>
      <td id="T_1666d_row103_col22" class="data row103 col22" >0</td>
      <td id="T_1666d_row103_col23" class="data row103 col23" >0</td>
      <td id="T_1666d_row103_col24" class="data row103 col24" >0</td>
      <td id="T_1666d_row103_col25" class="data row103 col25" >0</td>
      <td id="T_1666d_row103_col26" class="data row103 col26" >0</td>
      <td id="T_1666d_row103_col27" class="data row103 col27" >0</td>
      <td id="T_1666d_row103_col28" class="data row103 col28" >0</td>
      <td id="T_1666d_row103_col29" class="data row103 col29" >2,171</td>
      <td id="T_1666d_row103_col30" class="data row103 col30" >2,171</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row104" class="row_heading level2 row104" >D</th>
      <td id="T_1666d_row104_col0" class="data row104 col0" >0</td>
      <td id="T_1666d_row104_col1" class="data row104 col1" >0</td>
      <td id="T_1666d_row104_col2" class="data row104 col2" >0</td>
      <td id="T_1666d_row104_col3" class="data row104 col3" >0</td>
      <td id="T_1666d_row104_col4" class="data row104 col4" >0</td>
      <td id="T_1666d_row104_col5" class="data row104 col5" >0</td>
      <td id="T_1666d_row104_col6" class="data row104 col6" >1,608</td>
      <td id="T_1666d_row104_col7" class="data row104 col7" >0</td>
      <td id="T_1666d_row104_col8" class="data row104 col8" >0</td>
      <td id="T_1666d_row104_col9" class="data row104 col9" >15,336</td>
      <td id="T_1666d_row104_col10" class="data row104 col10" >0</td>
      <td id="T_1666d_row104_col11" class="data row104 col11" >0</td>
      <td id="T_1666d_row104_col12" class="data row104 col12" >0</td>
      <td id="T_1666d_row104_col13" class="data row104 col13" >0</td>
      <td id="T_1666d_row104_col14" class="data row104 col14" >0</td>
      <td id="T_1666d_row104_col15" class="data row104 col15" >0</td>
      <td id="T_1666d_row104_col16" class="data row104 col16" >0</td>
      <td id="T_1666d_row104_col17" class="data row104 col17" >0</td>
      <td id="T_1666d_row104_col18" class="data row104 col18" >0</td>
      <td id="T_1666d_row104_col19" class="data row104 col19" >0</td>
      <td id="T_1666d_row104_col20" class="data row104 col20" >11,200</td>
      <td id="T_1666d_row104_col21" class="data row104 col21" >518</td>
      <td id="T_1666d_row104_col22" class="data row104 col22" >0</td>
      <td id="T_1666d_row104_col23" class="data row104 col23" >0</td>
      <td id="T_1666d_row104_col24" class="data row104 col24" >23,058</td>
      <td id="T_1666d_row104_col25" class="data row104 col25" >0</td>
      <td id="T_1666d_row104_col26" class="data row104 col26" >0</td>
      <td id="T_1666d_row104_col27" class="data row104 col27" >0</td>
      <td id="T_1666d_row104_col28" class="data row104 col28" >23,059</td>
      <td id="T_1666d_row104_col29" class="data row104 col29" >3,421</td>
      <td id="T_1666d_row104_col30" class="data row104 col30" >78,200</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row105" class="row_heading level2 row105" >G</th>
      <td id="T_1666d_row105_col0" class="data row105 col0" >0</td>
      <td id="T_1666d_row105_col1" class="data row105 col1" >0</td>
      <td id="T_1666d_row105_col2" class="data row105 col2" >0</td>
      <td id="T_1666d_row105_col3" class="data row105 col3" >0</td>
      <td id="T_1666d_row105_col4" class="data row105 col4" >0</td>
      <td id="T_1666d_row105_col5" class="data row105 col5" >0</td>
      <td id="T_1666d_row105_col6" class="data row105 col6" >0</td>
      <td id="T_1666d_row105_col7" class="data row105 col7" >0</td>
      <td id="T_1666d_row105_col8" class="data row105 col8" >0</td>
      <td id="T_1666d_row105_col9" class="data row105 col9" >0</td>
      <td id="T_1666d_row105_col10" class="data row105 col10" >1</td>
      <td id="T_1666d_row105_col11" class="data row105 col11" >0</td>
      <td id="T_1666d_row105_col12" class="data row105 col12" >0</td>
      <td id="T_1666d_row105_col13" class="data row105 col13" >0</td>
      <td id="T_1666d_row105_col14" class="data row105 col14" >12,836</td>
      <td id="T_1666d_row105_col15" class="data row105 col15" >0</td>
      <td id="T_1666d_row105_col16" class="data row105 col16" >0</td>
      <td id="T_1666d_row105_col17" class="data row105 col17" >51,698</td>
      <td id="T_1666d_row105_col18" class="data row105 col18" >0</td>
      <td id="T_1666d_row105_col19" class="data row105 col19" >0</td>
      <td id="T_1666d_row105_col20" class="data row105 col20" >0</td>
      <td id="T_1666d_row105_col21" class="data row105 col21" >0</td>
      <td id="T_1666d_row105_col22" class="data row105 col22" >0</td>
      <td id="T_1666d_row105_col23" class="data row105 col23" >46,076</td>
      <td id="T_1666d_row105_col24" class="data row105 col24" >0</td>
      <td id="T_1666d_row105_col25" class="data row105 col25" >0</td>
      <td id="T_1666d_row105_col26" class="data row105 col26" >0</td>
      <td id="T_1666d_row105_col27" class="data row105 col27" >0</td>
      <td id="T_1666d_row105_col28" class="data row105 col28" >0</td>
      <td id="T_1666d_row105_col29" class="data row105 col29" >11,852</td>
      <td id="T_1666d_row105_col30" class="data row105 col30" >122,463</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row106" class="row_heading level2 row106" >A</th>
      <td id="T_1666d_row106_col0" class="data row106 col0" >170,015</td>
      <td id="T_1666d_row106_col1" class="data row106 col1" >246,911</td>
      <td id="T_1666d_row106_col2" class="data row106 col2" >252,000</td>
      <td id="T_1666d_row106_col3" class="data row106 col3" >0</td>
      <td id="T_1666d_row106_col4" class="data row106 col4" >0</td>
      <td id="T_1666d_row106_col5" class="data row106 col5" >0</td>
      <td id="T_1666d_row106_col6" class="data row106 col6" >0</td>
      <td id="T_1666d_row106_col7" class="data row106 col7" >0</td>
      <td id="T_1666d_row106_col8" class="data row106 col8" >0</td>
      <td id="T_1666d_row106_col9" class="data row106 col9" >14,389</td>
      <td id="T_1666d_row106_col10" class="data row106 col10" >0</td>
      <td id="T_1666d_row106_col11" class="data row106 col11" >0</td>
      <td id="T_1666d_row106_col12" class="data row106 col12" >0</td>
      <td id="T_1666d_row106_col13" class="data row106 col13" >0</td>
      <td id="T_1666d_row106_col14" class="data row106 col14" >0</td>
      <td id="T_1666d_row106_col15" class="data row106 col15" >103,765</td>
      <td id="T_1666d_row106_col16" class="data row106 col16" >0</td>
      <td id="T_1666d_row106_col17" class="data row106 col17" >0</td>
      <td id="T_1666d_row106_col18" class="data row106 col18" >0</td>
      <td id="T_1666d_row106_col19" class="data row106 col19" >0</td>
      <td id="T_1666d_row106_col20" class="data row106 col20" >72,635</td>
      <td id="T_1666d_row106_col21" class="data row106 col21" >0</td>
      <td id="T_1666d_row106_col22" class="data row106 col22" >0</td>
      <td id="T_1666d_row106_col23" class="data row106 col23" >0</td>
      <td id="T_1666d_row106_col24" class="data row106 col24" >72,131</td>
      <td id="T_1666d_row106_col25" class="data row106 col25" >0</td>
      <td id="T_1666d_row106_col26" class="data row106 col26" >0</td>
      <td id="T_1666d_row106_col27" class="data row106 col27" >0</td>
      <td id="T_1666d_row106_col28" class="data row106 col28" >0</td>
      <td id="T_1666d_row106_col29" class="data row106 col29" >0</td>
      <td id="T_1666d_row106_col30" class="data row106 col30" >931,846</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row107" class="row_heading level2 row107" >I</th>
      <td id="T_1666d_row107_col0" class="data row107 col0" >78,757</td>
      <td id="T_1666d_row107_col1" class="data row107 col1" >0</td>
      <td id="T_1666d_row107_col2" class="data row107 col2" >0</td>
      <td id="T_1666d_row107_col3" class="data row107 col3" >267,078</td>
      <td id="T_1666d_row107_col4" class="data row107 col4" >0</td>
      <td id="T_1666d_row107_col5" class="data row107 col5" >0</td>
      <td id="T_1666d_row107_col6" class="data row107 col6" >0</td>
      <td id="T_1666d_row107_col7" class="data row107 col7" >136,073</td>
      <td id="T_1666d_row107_col8" class="data row107 col8" >0</td>
      <td id="T_1666d_row107_col9" class="data row107 col9" >44,724</td>
      <td id="T_1666d_row107_col10" class="data row107 col10" >0</td>
      <td id="T_1666d_row107_col11" class="data row107 col11" >300,821</td>
      <td id="T_1666d_row107_col12" class="data row107 col12" >235,200</td>
      <td id="T_1666d_row107_col13" class="data row107 col13" >235,200</td>
      <td id="T_1666d_row107_col14" class="data row107 col14" >34,468</td>
      <td id="T_1666d_row107_col15" class="data row107 col15" >0</td>
      <td id="T_1666d_row107_col16" class="data row107 col16" >234,700</td>
      <td id="T_1666d_row107_col17" class="data row107 col17" >0</td>
      <td id="T_1666d_row107_col18" class="data row107 col18" >0</td>
      <td id="T_1666d_row107_col19" class="data row107 col19" >0</td>
      <td id="T_1666d_row107_col20" class="data row107 col20" >0</td>
      <td id="T_1666d_row107_col21" class="data row107 col21" >4,364</td>
      <td id="T_1666d_row107_col22" class="data row107 col22" >0</td>
      <td id="T_1666d_row107_col23" class="data row107 col23" >231,738</td>
      <td id="T_1666d_row107_col24" class="data row107 col24" >0</td>
      <td id="T_1666d_row107_col25" class="data row107 col25" >0</td>
      <td id="T_1666d_row107_col26" class="data row107 col26" >0</td>
      <td id="T_1666d_row107_col27" class="data row107 col27" >0</td>
      <td id="T_1666d_row107_col28" class="data row107 col28" >0</td>
      <td id="T_1666d_row107_col29" class="data row107 col29" >0</td>
      <td id="T_1666d_row107_col30" class="data row107 col30" >1,803,123</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row108" class="row_heading level2 row108" >E</th>
      <td id="T_1666d_row108_col0" class="data row108 col0" >0</td>
      <td id="T_1666d_row108_col1" class="data row108 col1" >0</td>
      <td id="T_1666d_row108_col2" class="data row108 col2" >0</td>
      <td id="T_1666d_row108_col3" class="data row108 col3" >0</td>
      <td id="T_1666d_row108_col4" class="data row108 col4" >0</td>
      <td id="T_1666d_row108_col5" class="data row108 col5" >0</td>
      <td id="T_1666d_row108_col6" class="data row108 col6" >0</td>
      <td id="T_1666d_row108_col7" class="data row108 col7" >0</td>
      <td id="T_1666d_row108_col8" class="data row108 col8" >0</td>
      <td id="T_1666d_row108_col9" class="data row108 col9" >0</td>
      <td id="T_1666d_row108_col10" class="data row108 col10" >0</td>
      <td id="T_1666d_row108_col11" class="data row108 col11" >0</td>
      <td id="T_1666d_row108_col12" class="data row108 col12" >0</td>
      <td id="T_1666d_row108_col13" class="data row108 col13" >0</td>
      <td id="T_1666d_row108_col14" class="data row108 col14" >0</td>
      <td id="T_1666d_row108_col15" class="data row108 col15" >0</td>
      <td id="T_1666d_row108_col16" class="data row108 col16" >0</td>
      <td id="T_1666d_row108_col17" class="data row108 col17" >0</td>
      <td id="T_1666d_row108_col18" class="data row108 col18" >0</td>
      <td id="T_1666d_row108_col19" class="data row108 col19" >0</td>
      <td id="T_1666d_row108_col20" class="data row108 col20" >0</td>
      <td id="T_1666d_row108_col21" class="data row108 col21" >0</td>
      <td id="T_1666d_row108_col22" class="data row108 col22" >0</td>
      <td id="T_1666d_row108_col23" class="data row108 col23" >0</td>
      <td id="T_1666d_row108_col24" class="data row108 col24" >1,016</td>
      <td id="T_1666d_row108_col25" class="data row108 col25" >0</td>
      <td id="T_1666d_row108_col26" class="data row108 col26" >0</td>
      <td id="T_1666d_row108_col27" class="data row108 col27" >0</td>
      <td id="T_1666d_row108_col28" class="data row108 col28" >0</td>
      <td id="T_1666d_row108_col29" class="data row108 col29" >0</td>
      <td id="T_1666d_row108_col30" class="data row108 col30" >1,016</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row109" class="row_heading level2 row109" >C</th>
      <td id="T_1666d_row109_col0" class="data row109 col0" >0</td>
      <td id="T_1666d_row109_col1" class="data row109 col1" >0</td>
      <td id="T_1666d_row109_col2" class="data row109 col2" >0</td>
      <td id="T_1666d_row109_col3" class="data row109 col3" >0</td>
      <td id="T_1666d_row109_col4" class="data row109 col4" >0</td>
      <td id="T_1666d_row109_col5" class="data row109 col5" >0</td>
      <td id="T_1666d_row109_col6" class="data row109 col6" >0</td>
      <td id="T_1666d_row109_col7" class="data row109 col7" >59,476</td>
      <td id="T_1666d_row109_col8" class="data row109 col8" >0</td>
      <td id="T_1666d_row109_col9" class="data row109 col9" >47,565</td>
      <td id="T_1666d_row109_col10" class="data row109 col10" >81,769</td>
      <td id="T_1666d_row109_col11" class="data row109 col11" >21,107</td>
      <td id="T_1666d_row109_col12" class="data row109 col12" >0</td>
      <td id="T_1666d_row109_col13" class="data row109 col13" >0</td>
      <td id="T_1666d_row109_col14" class="data row109 col14" >0</td>
      <td id="T_1666d_row109_col15" class="data row109 col15" >0</td>
      <td id="T_1666d_row109_col16" class="data row109 col16" >0</td>
      <td id="T_1666d_row109_col17" class="data row109 col17" >0</td>
      <td id="T_1666d_row109_col18" class="data row109 col18" >0</td>
      <td id="T_1666d_row109_col19" class="data row109 col19" >0</td>
      <td id="T_1666d_row109_col20" class="data row109 col20" >0</td>
      <td id="T_1666d_row109_col21" class="data row109 col21" >125,344</td>
      <td id="T_1666d_row109_col22" class="data row109 col22" >103,231</td>
      <td id="T_1666d_row109_col23" class="data row109 col23" >0</td>
      <td id="T_1666d_row109_col24" class="data row109 col24" >0</td>
      <td id="T_1666d_row109_col25" class="data row109 col25" >0</td>
      <td id="T_1666d_row109_col26" class="data row109 col26" >0</td>
      <td id="T_1666d_row109_col27" class="data row109 col27" >201,600</td>
      <td id="T_1666d_row109_col28" class="data row109 col28" >0</td>
      <td id="T_1666d_row109_col29" class="data row109 col29" >170,485</td>
      <td id="T_1666d_row109_col30" class="data row109 col30" >810,577</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row110" class="row_heading level2 row110" >F</th>
      <td id="T_1666d_row110_col0" class="data row110 col0" >0</td>
      <td id="T_1666d_row110_col1" class="data row110 col1" >6,785</td>
      <td id="T_1666d_row110_col2" class="data row110 col2" >0</td>
      <td id="T_1666d_row110_col3" class="data row110 col3" >0</td>
      <td id="T_1666d_row110_col4" class="data row110 col4" >0</td>
      <td id="T_1666d_row110_col5" class="data row110 col5" >0</td>
      <td id="T_1666d_row110_col6" class="data row110 col6" >0</td>
      <td id="T_1666d_row110_col7" class="data row110 col7" >0</td>
      <td id="T_1666d_row110_col8" class="data row110 col8" >0</td>
      <td id="T_1666d_row110_col9" class="data row110 col9" >0</td>
      <td id="T_1666d_row110_col10" class="data row110 col10" >2,071</td>
      <td id="T_1666d_row110_col11" class="data row110 col11" >0</td>
      <td id="T_1666d_row110_col12" class="data row110 col12" >0</td>
      <td id="T_1666d_row110_col13" class="data row110 col13" >0</td>
      <td id="T_1666d_row110_col14" class="data row110 col14" >0</td>
      <td id="T_1666d_row110_col15" class="data row110 col15" >0</td>
      <td id="T_1666d_row110_col16" class="data row110 col16" >500</td>
      <td id="T_1666d_row110_col17" class="data row110 col17" >138,166</td>
      <td id="T_1666d_row110_col18" class="data row110 col18" >0</td>
      <td id="T_1666d_row110_col19" class="data row110 col19" >0</td>
      <td id="T_1666d_row110_col20" class="data row110 col20" >0</td>
      <td id="T_1666d_row110_col21" class="data row110 col21" >0</td>
      <td id="T_1666d_row110_col22" class="data row110 col22" >63,149</td>
      <td id="T_1666d_row110_col23" class="data row110 col23" >0</td>
      <td id="T_1666d_row110_col24" class="data row110 col24" >0</td>
      <td id="T_1666d_row110_col25" class="data row110 col25" >0</td>
      <td id="T_1666d_row110_col26" class="data row110 col26" >0</td>
      <td id="T_1666d_row110_col27" class="data row110 col27" >0</td>
      <td id="T_1666d_row110_col28" class="data row110 col28" >0</td>
      <td id="T_1666d_row110_col29" class="data row110 col29" >0</td>
      <td id="T_1666d_row110_col30" class="data row110 col30" >210,671</td>
    </tr>
    <tr>
      <th id="T_1666d_level0_row111" class="row_heading level0 row111" rowspan="7">Pac1</th>
      <th id="T_1666d_level1_row111" class="row_heading level1 row111" rowspan="7">Day</th>
      <th id="T_1666d_level2_row111" class="row_heading level2 row111" >B</th>
      <td id="T_1666d_row111_col0" class="data row111 col0" >0</td>
      <td id="T_1666d_row111_col1" class="data row111 col1" >201,273</td>
      <td id="T_1666d_row111_col2" class="data row111 col2" >409,994</td>
      <td id="T_1666d_row111_col3" class="data row111 col3" >0</td>
      <td id="T_1666d_row111_col4" class="data row111 col4" >0</td>
      <td id="T_1666d_row111_col5" class="data row111 col5" >0</td>
      <td id="T_1666d_row111_col6" class="data row111 col6" >5,547</td>
      <td id="T_1666d_row111_col7" class="data row111 col7" >0</td>
      <td id="T_1666d_row111_col8" class="data row111 col8" >844,091</td>
      <td id="T_1666d_row111_col9" class="data row111 col9" >0</td>
      <td id="T_1666d_row111_col10" class="data row111 col10" >0</td>
      <td id="T_1666d_row111_col11" class="data row111 col11" >0</td>
      <td id="T_1666d_row111_col12" class="data row111 col12" >204,081</td>
      <td id="T_1666d_row111_col13" class="data row111 col13" >2,669</td>
      <td id="T_1666d_row111_col14" class="data row111 col14" >0</td>
      <td id="T_1666d_row111_col15" class="data row111 col15" >0</td>
      <td id="T_1666d_row111_col16" class="data row111 col16" >0</td>
      <td id="T_1666d_row111_col17" class="data row111 col17" >680,048</td>
      <td id="T_1666d_row111_col18" class="data row111 col18" >0</td>
      <td id="T_1666d_row111_col19" class="data row111 col19" >0</td>
      <td id="T_1666d_row111_col20" class="data row111 col20" >0</td>
      <td id="T_1666d_row111_col21" class="data row111 col21" >722,219</td>
      <td id="T_1666d_row111_col22" class="data row111 col22" >30,078</td>
      <td id="T_1666d_row111_col23" class="data row111 col23" >0</td>
      <td id="T_1666d_row111_col24" class="data row111 col24" >0</td>
      <td id="T_1666d_row111_col25" class="data row111 col25" >0</td>
      <td id="T_1666d_row111_col26" class="data row111 col26" >0</td>
      <td id="T_1666d_row111_col27" class="data row111 col27" >0</td>
      <td id="T_1666d_row111_col28" class="data row111 col28" >0</td>
      <td id="T_1666d_row111_col29" class="data row111 col29" >0</td>
      <td id="T_1666d_row111_col30" class="data row111 col30" >3,100,000</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row112" class="row_heading level2 row112" >D</th>
      <td id="T_1666d_row112_col0" class="data row112 col0" >184,633</td>
      <td id="T_1666d_row112_col1" class="data row112 col1" >89,905</td>
      <td id="T_1666d_row112_col2" class="data row112 col2" >0</td>
      <td id="T_1666d_row112_col3" class="data row112 col3" >0</td>
      <td id="T_1666d_row112_col4" class="data row112 col4" >0</td>
      <td id="T_1666d_row112_col5" class="data row112 col5" >0</td>
      <td id="T_1666d_row112_col6" class="data row112 col6" >1,620</td>
      <td id="T_1666d_row112_col7" class="data row112 col7" >0</td>
      <td id="T_1666d_row112_col8" class="data row112 col8" >0</td>
      <td id="T_1666d_row112_col9" class="data row112 col9" >4,102</td>
      <td id="T_1666d_row112_col10" class="data row112 col10" >0</td>
      <td id="T_1666d_row112_col11" class="data row112 col11" >14,740</td>
      <td id="T_1666d_row112_col12" class="data row112 col12" >0</td>
      <td id="T_1666d_row112_col13" class="data row112 col13" >0</td>
      <td id="T_1666d_row112_col14" class="data row112 col14" >0</td>
      <td id="T_1666d_row112_col15" class="data row112 col15" >0</td>
      <td id="T_1666d_row112_col16" class="data row112 col16" >0</td>
      <td id="T_1666d_row112_col17" class="data row112 col17" >0</td>
      <td id="T_1666d_row112_col18" class="data row112 col18" >0</td>
      <td id="T_1666d_row112_col19" class="data row112 col19" >0</td>
      <td id="T_1666d_row112_col20" class="data row112 col20" >0</td>
      <td id="T_1666d_row112_col21" class="data row112 col21" >0</td>
      <td id="T_1666d_row112_col22" class="data row112 col22" >0</td>
      <td id="T_1666d_row112_col23" class="data row112 col23" >0</td>
      <td id="T_1666d_row112_col24" class="data row112 col24" >0</td>
      <td id="T_1666d_row112_col25" class="data row112 col25" >0</td>
      <td id="T_1666d_row112_col26" class="data row112 col26" >0</td>
      <td id="T_1666d_row112_col27" class="data row112 col27" >0</td>
      <td id="T_1666d_row112_col28" class="data row112 col28" >0</td>
      <td id="T_1666d_row112_col29" class="data row112 col29" >0</td>
      <td id="T_1666d_row112_col30" class="data row112 col30" >295,000</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row113" class="row_heading level2 row113" >A</th>
      <td id="T_1666d_row113_col0" class="data row113 col0" >43,590</td>
      <td id="T_1666d_row113_col1" class="data row113 col1" >0</td>
      <td id="T_1666d_row113_col2" class="data row113 col2" >0</td>
      <td id="T_1666d_row113_col3" class="data row113 col3" >1,000</td>
      <td id="T_1666d_row113_col4" class="data row113 col4" >0</td>
      <td id="T_1666d_row113_col5" class="data row113 col5" >0</td>
      <td id="T_1666d_row113_col6" class="data row113 col6" >0</td>
      <td id="T_1666d_row113_col7" class="data row113 col7" >0</td>
      <td id="T_1666d_row113_col8" class="data row113 col8" >0</td>
      <td id="T_1666d_row113_col9" class="data row113 col9" >0</td>
      <td id="T_1666d_row113_col10" class="data row113 col10" >0</td>
      <td id="T_1666d_row113_col11" class="data row113 col11" >0</td>
      <td id="T_1666d_row113_col12" class="data row113 col12" >2,000</td>
      <td id="T_1666d_row113_col13" class="data row113 col13" >0</td>
      <td id="T_1666d_row113_col14" class="data row113 col14" >0</td>
      <td id="T_1666d_row113_col15" class="data row113 col15" >0</td>
      <td id="T_1666d_row113_col16" class="data row113 col16" >0</td>
      <td id="T_1666d_row113_col17" class="data row113 col17" >0</td>
      <td id="T_1666d_row113_col18" class="data row113 col18" >0</td>
      <td id="T_1666d_row113_col19" class="data row113 col19" >0</td>
      <td id="T_1666d_row113_col20" class="data row113 col20" >0</td>
      <td id="T_1666d_row113_col21" class="data row113 col21" >0</td>
      <td id="T_1666d_row113_col22" class="data row113 col22" >0</td>
      <td id="T_1666d_row113_col23" class="data row113 col23" >0</td>
      <td id="T_1666d_row113_col24" class="data row113 col24" >0</td>
      <td id="T_1666d_row113_col25" class="data row113 col25" >0</td>
      <td id="T_1666d_row113_col26" class="data row113 col26" >0</td>
      <td id="T_1666d_row113_col27" class="data row113 col27" >0</td>
      <td id="T_1666d_row113_col28" class="data row113 col28" >0</td>
      <td id="T_1666d_row113_col29" class="data row113 col29" >0</td>
      <td id="T_1666d_row113_col30" class="data row113 col30" >46,590</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row114" class="row_heading level2 row114" >I</th>
      <td id="T_1666d_row114_col0" class="data row114 col0" >188,570</td>
      <td id="T_1666d_row114_col1" class="data row114 col1" >375,852</td>
      <td id="T_1666d_row114_col2" class="data row114 col2" >410,919</td>
      <td id="T_1666d_row114_col3" class="data row114 col3" >714,146</td>
      <td id="T_1666d_row114_col4" class="data row114 col4" >0</td>
      <td id="T_1666d_row114_col5" class="data row114 col5" >0</td>
      <td id="T_1666d_row114_col6" class="data row114 col6" >651,017</td>
      <td id="T_1666d_row114_col7" class="data row114 col7" >2,584</td>
      <td id="T_1666d_row114_col8" class="data row114 col8" >86,931</td>
      <td id="T_1666d_row114_col9" class="data row114 col9" >713,847</td>
      <td id="T_1666d_row114_col10" class="data row114 col10" >720,000</td>
      <td id="T_1666d_row114_col11" class="data row114 col11" >697,890</td>
      <td id="T_1666d_row114_col12" class="data row114 col12" >550,221</td>
      <td id="T_1666d_row114_col13" class="data row114 col13" >717,998</td>
      <td id="T_1666d_row114_col14" class="data row114 col14" >377,193</td>
      <td id="T_1666d_row114_col15" class="data row114 col15" >83,229</td>
      <td id="T_1666d_row114_col16" class="data row114 col16" >719,108</td>
      <td id="T_1666d_row114_col17" class="data row114 col17" >0</td>
      <td id="T_1666d_row114_col18" class="data row114 col18" >0</td>
      <td id="T_1666d_row114_col19" class="data row114 col19" >0</td>
      <td id="T_1666d_row114_col20" class="data row114 col20" >710,868</td>
      <td id="T_1666d_row114_col21" class="data row114 col21" >177,319</td>
      <td id="T_1666d_row114_col22" class="data row114 col22" >2,902</td>
      <td id="T_1666d_row114_col23" class="data row114 col23" >719,383</td>
      <td id="T_1666d_row114_col24" class="data row114 col24" >1,000</td>
      <td id="T_1666d_row114_col25" class="data row114 col25" >0</td>
      <td id="T_1666d_row114_col26" class="data row114 col26" >0</td>
      <td id="T_1666d_row114_col27" class="data row114 col27" >138,577</td>
      <td id="T_1666d_row114_col28" class="data row114 col28" >3,446</td>
      <td id="T_1666d_row114_col29" class="data row114 col29" >0</td>
      <td id="T_1666d_row114_col30" class="data row114 col30" >8,763,000</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row115" class="row_heading level2 row115" >E</th>
      <td id="T_1666d_row115_col0" class="data row115 col0" >209,891</td>
      <td id="T_1666d_row115_col1" class="data row115 col1" >0</td>
      <td id="T_1666d_row115_col2" class="data row115 col2" >0</td>
      <td id="T_1666d_row115_col3" class="data row115 col3" >0</td>
      <td id="T_1666d_row115_col4" class="data row115 col4" >0</td>
      <td id="T_1666d_row115_col5" class="data row115 col5" >0</td>
      <td id="T_1666d_row115_col6" class="data row115 col6" >0</td>
      <td id="T_1666d_row115_col7" class="data row115 col7" >0</td>
      <td id="T_1666d_row115_col8" class="data row115 col8" >0</td>
      <td id="T_1666d_row115_col9" class="data row115 col9" >0</td>
      <td id="T_1666d_row115_col10" class="data row115 col10" >0</td>
      <td id="T_1666d_row115_col11" class="data row115 col11" >0</td>
      <td id="T_1666d_row115_col12" class="data row115 col12" >0</td>
      <td id="T_1666d_row115_col13" class="data row115 col13" >0</td>
      <td id="T_1666d_row115_col14" class="data row115 col14" >0</td>
      <td id="T_1666d_row115_col15" class="data row115 col15" >0</td>
      <td id="T_1666d_row115_col16" class="data row115 col16" >0</td>
      <td id="T_1666d_row115_col17" class="data row115 col17" >0</td>
      <td id="T_1666d_row115_col18" class="data row115 col18" >0</td>
      <td id="T_1666d_row115_col19" class="data row115 col19" >0</td>
      <td id="T_1666d_row115_col20" class="data row115 col20" >0</td>
      <td id="T_1666d_row115_col21" class="data row115 col21" >0</td>
      <td id="T_1666d_row115_col22" class="data row115 col22" >0</td>
      <td id="T_1666d_row115_col23" class="data row115 col23" >0</td>
      <td id="T_1666d_row115_col24" class="data row115 col24" >0</td>
      <td id="T_1666d_row115_col25" class="data row115 col25" >0</td>
      <td id="T_1666d_row115_col26" class="data row115 col26" >0</td>
      <td id="T_1666d_row115_col27" class="data row115 col27" >0</td>
      <td id="T_1666d_row115_col28" class="data row115 col28" >0</td>
      <td id="T_1666d_row115_col29" class="data row115 col29" >0</td>
      <td id="T_1666d_row115_col30" class="data row115 col30" >209,891</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row116" class="row_heading level2 row116" >C</th>
      <td id="T_1666d_row116_col0" class="data row116 col0" >0</td>
      <td id="T_1666d_row116_col1" class="data row116 col1" >0</td>
      <td id="T_1666d_row116_col2" class="data row116 col2" >0</td>
      <td id="T_1666d_row116_col3" class="data row116 col3" >0</td>
      <td id="T_1666d_row116_col4" class="data row116 col4" >0</td>
      <td id="T_1666d_row116_col5" class="data row116 col5" >0</td>
      <td id="T_1666d_row116_col6" class="data row116 col6" >0</td>
      <td id="T_1666d_row116_col7" class="data row116 col7" >717,416</td>
      <td id="T_1666d_row116_col8" class="data row116 col8" >0</td>
      <td id="T_1666d_row116_col9" class="data row116 col9" >0</td>
      <td id="T_1666d_row116_col10" class="data row116 col10" >0</td>
      <td id="T_1666d_row116_col11" class="data row116 col11" >0</td>
      <td id="T_1666d_row116_col12" class="data row116 col12" >14,718</td>
      <td id="T_1666d_row116_col13" class="data row116 col13" >0</td>
      <td id="T_1666d_row116_col14" class="data row116 col14" >342,807</td>
      <td id="T_1666d_row116_col15" class="data row116 col15" >636,771</td>
      <td id="T_1666d_row116_col16" class="data row116 col16" >0</td>
      <td id="T_1666d_row116_col17" class="data row116 col17" >177,473</td>
      <td id="T_1666d_row116_col18" class="data row116 col18" >0</td>
      <td id="T_1666d_row116_col19" class="data row116 col19" >0</td>
      <td id="T_1666d_row116_col20" class="data row116 col20" >9,132</td>
      <td id="T_1666d_row116_col21" class="data row116 col21" >1,016</td>
      <td id="T_1666d_row116_col22" class="data row116 col22" >694,539</td>
      <td id="T_1666d_row116_col23" class="data row116 col23" >0</td>
      <td id="T_1666d_row116_col24" class="data row116 col24" >0</td>
      <td id="T_1666d_row116_col25" class="data row116 col25" >0</td>
      <td id="T_1666d_row116_col26" class="data row116 col26" >0</td>
      <td id="T_1666d_row116_col27" class="data row116 col27" >581,423</td>
      <td id="T_1666d_row116_col28" class="data row116 col28" >716,446</td>
      <td id="T_1666d_row116_col29" class="data row116 col29" >674,259</td>
      <td id="T_1666d_row116_col30" class="data row116 col30" >4,566,000</td>
    </tr>
    <tr>
      <th id="T_1666d_level2_row117" class="row_heading level2 row117" >F</th>
      <td id="T_1666d_row117_col0" class="data row117 col0" >0</td>
      <td id="T_1666d_row117_col1" class="data row117 col1" >77,781</td>
      <td id="T_1666d_row117_col2" class="data row117 col2" >2,114</td>
      <td id="T_1666d_row117_col3" class="data row117 col3" >6,472</td>
      <td id="T_1666d_row117_col4" class="data row117 col4" >0</td>
      <td id="T_1666d_row117_col5" class="data row117 col5" >0</td>
      <td id="T_1666d_row117_col6" class="data row117 col6" >83,190</td>
      <td id="T_1666d_row117_col7" class="data row117 col7" >0</td>
      <td id="T_1666d_row117_col8" class="data row117 col8" >0</td>
      <td id="T_1666d_row117_col9" class="data row117 col9" >0</td>
      <td id="T_1666d_row117_col10" class="data row117 col10" >0</td>
      <td id="T_1666d_row117_col11" class="data row117 col11" >0</td>
      <td id="T_1666d_row117_col12" class="data row117 col12" >0</td>
      <td id="T_1666d_row117_col13" class="data row117 col13" >0</td>
      <td id="T_1666d_row117_col14" class="data row117 col14" >0</td>
      <td id="T_1666d_row117_col15" class="data row117 col15" >0</td>
      <td id="T_1666d_row117_col16" class="data row117 col16" >0</td>
      <td id="T_1666d_row117_col17" class="data row117 col17" >42,789</td>
      <td id="T_1666d_row117_col18" class="data row117 col18" >0</td>
      <td id="T_1666d_row117_col19" class="data row117 col19" >0</td>
      <td id="T_1666d_row117_col20" class="data row117 col20" >0</td>
      <td id="T_1666d_row117_col21" class="data row117 col21" >0</td>
      <td id="T_1666d_row117_col22" class="data row117 col22" >0</td>
      <td id="T_1666d_row117_col23" class="data row117 col23" >0</td>
      <td id="T_1666d_row117_col24" class="data row117 col24" >958,666</td>
      <td id="T_1666d_row117_col25" class="data row117 col25" >0</td>
      <td id="T_1666d_row117_col26" class="data row117 col26" >0</td>
      <td id="T_1666d_row117_col27" class="data row117 col27" >0</td>
      <td id="T_1666d_row117_col28" class="data row117 col28" >0</td>
      <td id="T_1666d_row117_col29" class="data row117 col29" >60,988</td>
      <td id="T_1666d_row117_col30" class="data row117 col30" >1,232,000</td>
    </tr>
  </tbody>
</table>


</details>

<!-- END TABLE:regular-schedule -->

<!-- BEGIN TABLE:ot-attainment -->

<details>
<summary>OT DEMAND ATTAINMENT</summary>

<style type="text/css">
</style>
<table id="T_31921">
  <thead>
    <tr>
      <th class="blank level0" >&nbsp;</th>
      <th id="T_31921_level0_col0" class="col_heading level0 col0" >Original Forecast</th>
      <th id="T_31921_level0_col1" class="col_heading level0 col1" >Solved Regular</th>
      <th id="T_31921_level0_col2" class="col_heading level0 col2" >For OT</th>
      <th id="T_31921_level0_col3" class="col_heading level0 col3" >Solved OT</th>
      <th id="T_31921_level0_col4" class="col_heading level0 col4" >Remaining After OT</th>
      <th id="T_31921_level0_col5" class="col_heading level0 col5" >OT Attainment</th>
    </tr>
    <tr>
      <th class="index_name level0" >Product</th>
      <th class="blank col0" >&nbsp;</th>
      <th class="blank col1" >&nbsp;</th>
      <th class="blank col2" >&nbsp;</th>
      <th class="blank col3" >&nbsp;</th>
      <th class="blank col4" >&nbsp;</th>
      <th class="blank col5" >&nbsp;</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <th id="T_31921_level0_row0" class="row_heading level0 row0" >A</th>
      <td id="T_31921_row0_col0" class="data row0 col0" >3,000,000</td>
      <td id="T_31921_row0_col1" class="data row0 col1" >46,590</td>
      <td id="T_31921_row0_col2" class="data row0 col2" >2,953,410</td>
      <td id="T_31921_row0_col3" class="data row0 col3" >1,293,512</td>
      <td id="T_31921_row0_col4" class="data row0 col4" >1,659,898</td>
      <td id="T_31921_row0_col5" class="data row0 col5" >43.8%</td>
    </tr>
    <tr>
      <th id="T_31921_level0_row1" class="row_heading level0 row1" >B</th>
      <td id="T_31921_row1_col0" class="data row1 col0" >3,100,000</td>
      <td id="T_31921_row1_col1" class="data row1 col1" >3,100,000</td>
      <td id="T_31921_row1_col2" class="data row1 col2" >0</td>
      <td id="T_31921_row1_col3" class="data row1 col3" >0</td>
      <td id="T_31921_row1_col4" class="data row1 col4" >0</td>
      <td id="T_31921_row1_col5" class="data row1 col5" >0.0%</td>
    </tr>
    <tr>
      <th id="T_31921_level0_row2" class="row_heading level0 row2" >C</th>
      <td id="T_31921_row2_col0" class="data row2 col0" >4,566,000</td>
      <td id="T_31921_row2_col1" class="data row2 col1" >4,566,000</td>
      <td id="T_31921_row2_col2" class="data row2 col2" >0</td>
      <td id="T_31921_row2_col3" class="data row2 col3" >0</td>
      <td id="T_31921_row2_col4" class="data row2 col4" >0</td>
      <td id="T_31921_row2_col5" class="data row2 col5" >0.0%</td>
    </tr>
    <tr>
      <th id="T_31921_level0_row3" class="row_heading level0 row3" >D</th>
      <td id="T_31921_row3_col0" class="data row3 col0" >295,000</td>
      <td id="T_31921_row3_col1" class="data row3 col1" >295,000</td>
      <td id="T_31921_row3_col2" class="data row3 col2" >0</td>
      <td id="T_31921_row3_col3" class="data row3 col3" >0</td>
      <td id="T_31921_row3_col4" class="data row3 col4" >0</td>
      <td id="T_31921_row3_col5" class="data row3 col5" >0.0%</td>
    </tr>
    <tr>
      <th id="T_31921_level0_row4" class="row_heading level0 row4" >E</th>
      <td id="T_31921_row4_col0" class="data row4 col0" >1,233,000</td>
      <td id="T_31921_row4_col1" class="data row4 col1" >209,891</td>
      <td id="T_31921_row4_col2" class="data row4 col2" >1,023,109</td>
      <td id="T_31921_row4_col3" class="data row4 col3" >1,023,109</td>
      <td id="T_31921_row4_col4" class="data row4 col4" >0</td>
      <td id="T_31921_row4_col5" class="data row4 col5" >100.0%</td>
    </tr>
    <tr>
      <th id="T_31921_level0_row5" class="row_heading level0 row5" >F</th>
      <td id="T_31921_row5_col0" class="data row5 col0" >1,234,000</td>
      <td id="T_31921_row5_col1" class="data row5 col1" >1,232,000</td>
      <td id="T_31921_row5_col2" class="data row5 col2" >2,000</td>
      <td id="T_31921_row5_col3" class="data row5 col3" >2,000</td>
      <td id="T_31921_row5_col4" class="data row5 col4" >0</td>
      <td id="T_31921_row5_col5" class="data row5 col5" >100.0%</td>
    </tr>
    <tr>
      <th id="T_31921_level0_row6" class="row_heading level0 row6" >G</th>
      <td id="T_31921_row6_col0" class="data row6 col0" >134,500</td>
      <td id="T_31921_row6_col1" class="data row6 col1" >0</td>
      <td id="T_31921_row6_col2" class="data row6 col2" >134,500</td>
      <td id="T_31921_row6_col3" class="data row6 col3" >134,500</td>
      <td id="T_31921_row6_col4" class="data row6 col4" >0</td>
      <td id="T_31921_row6_col5" class="data row6 col5" >100.0%</td>
    </tr>
    <tr>
      <th id="T_31921_level0_row7" class="row_heading level0 row7" >H</th>
      <td id="T_31921_row7_col0" class="data row7 col0" >1,234,000</td>
      <td id="T_31921_row7_col1" class="data row7 col1" >0</td>
      <td id="T_31921_row7_col2" class="data row7 col2" >1,234,000</td>
      <td id="T_31921_row7_col3" class="data row7 col3" >297,950</td>
      <td id="T_31921_row7_col4" class="data row7 col4" >936,050</td>
      <td id="T_31921_row7_col5" class="data row7 col5" >24.1%</td>
    </tr>
    <tr>
      <th id="T_31921_level0_row8" class="row_heading level0 row8" >I</th>
      <td id="T_31921_row8_col0" class="data row8 col0" >8,765,000</td>
      <td id="T_31921_row8_col1" class="data row8 col1" >8,763,000</td>
      <td id="T_31921_row8_col2" class="data row8 col2" >2,000</td>
      <td id="T_31921_row8_col3" class="data row8 col3" >2,000</td>
      <td id="T_31921_row8_col4" class="data row8 col4" >0</td>
      <td id="T_31921_row8_col5" class="data row8 col5" >100.0%</td>
    </tr>
  </tbody>
</table>


</details>

<!-- END TABLE:ot-attainment -->

<!-- BEGIN TABLE:ot-finished-output -->

<details>
<summary>FINISHED OT OUTPUT</summary>

<style type="text/css">
</style>
<table id="T_e0b74">
  <thead>
    <tr>
      <th class="blank level0" >&nbsp;</th>
      <th id="T_e0b74_level0_col0" class="col_heading level0 col0" >2026-09-01</th>
      <th id="T_e0b74_level0_col1" class="col_heading level0 col1" >2026-09-02</th>
      <th id="T_e0b74_level0_col2" class="col_heading level0 col2" >2026-09-03</th>
      <th id="T_e0b74_level0_col3" class="col_heading level0 col3" >2026-09-04</th>
      <th id="T_e0b74_level0_col4" class="col_heading level0 col4" >2026-09-05</th>
      <th id="T_e0b74_level0_col5" class="col_heading level0 col5" >2026-09-06</th>
      <th id="T_e0b74_level0_col6" class="col_heading level0 col6" >2026-09-07</th>
      <th id="T_e0b74_level0_col7" class="col_heading level0 col7" >2026-09-08</th>
      <th id="T_e0b74_level0_col8" class="col_heading level0 col8" >2026-09-09</th>
      <th id="T_e0b74_level0_col9" class="col_heading level0 col9" >2026-09-10</th>
      <th id="T_e0b74_level0_col10" class="col_heading level0 col10" >2026-09-11</th>
      <th id="T_e0b74_level0_col11" class="col_heading level0 col11" >2026-09-12</th>
      <th id="T_e0b74_level0_col12" class="col_heading level0 col12" >2026-09-13</th>
      <th id="T_e0b74_level0_col13" class="col_heading level0 col13" >2026-09-14</th>
      <th id="T_e0b74_level0_col14" class="col_heading level0 col14" >2026-09-15</th>
      <th id="T_e0b74_level0_col15" class="col_heading level0 col15" >2026-09-16</th>
      <th id="T_e0b74_level0_col16" class="col_heading level0 col16" >2026-09-17</th>
      <th id="T_e0b74_level0_col17" class="col_heading level0 col17" >2026-09-18</th>
      <th id="T_e0b74_level0_col18" class="col_heading level0 col18" >2026-09-19</th>
      <th id="T_e0b74_level0_col19" class="col_heading level0 col19" >2026-09-20</th>
      <th id="T_e0b74_level0_col20" class="col_heading level0 col20" >2026-09-21</th>
      <th id="T_e0b74_level0_col21" class="col_heading level0 col21" >2026-09-22</th>
      <th id="T_e0b74_level0_col22" class="col_heading level0 col22" >2026-09-23</th>
      <th id="T_e0b74_level0_col23" class="col_heading level0 col23" >2026-09-24</th>
      <th id="T_e0b74_level0_col24" class="col_heading level0 col24" >2026-09-25</th>
      <th id="T_e0b74_level0_col25" class="col_heading level0 col25" >2026-09-26</th>
      <th id="T_e0b74_level0_col26" class="col_heading level0 col26" >2026-09-27</th>
      <th id="T_e0b74_level0_col27" class="col_heading level0 col27" >2026-09-28</th>
      <th id="T_e0b74_level0_col28" class="col_heading level0 col28" >2026-09-29</th>
      <th id="T_e0b74_level0_col29" class="col_heading level0 col29" >2026-09-30</th>
      <th id="T_e0b74_level0_col30" class="col_heading level0 col30" >Total</th>
    </tr>
    <tr>
      <th class="index_name level0" >Product</th>
      <th class="blank col0" >&nbsp;</th>
      <th class="blank col1" >&nbsp;</th>
      <th class="blank col2" >&nbsp;</th>
      <th class="blank col3" >&nbsp;</th>
      <th class="blank col4" >&nbsp;</th>
      <th class="blank col5" >&nbsp;</th>
      <th class="blank col6" >&nbsp;</th>
      <th class="blank col7" >&nbsp;</th>
      <th class="blank col8" >&nbsp;</th>
      <th class="blank col9" >&nbsp;</th>
      <th class="blank col10" >&nbsp;</th>
      <th class="blank col11" >&nbsp;</th>
      <th class="blank col12" >&nbsp;</th>
      <th class="blank col13" >&nbsp;</th>
      <th class="blank col14" >&nbsp;</th>
      <th class="blank col15" >&nbsp;</th>
      <th class="blank col16" >&nbsp;</th>
      <th class="blank col17" >&nbsp;</th>
      <th class="blank col18" >&nbsp;</th>
      <th class="blank col19" >&nbsp;</th>
      <th class="blank col20" >&nbsp;</th>
      <th class="blank col21" >&nbsp;</th>
      <th class="blank col22" >&nbsp;</th>
      <th class="blank col23" >&nbsp;</th>
      <th class="blank col24" >&nbsp;</th>
      <th class="blank col25" >&nbsp;</th>
      <th class="blank col26" >&nbsp;</th>
      <th class="blank col27" >&nbsp;</th>
      <th class="blank col28" >&nbsp;</th>
      <th class="blank col29" >&nbsp;</th>
      <th class="blank col30" >&nbsp;</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <th id="T_e0b74_level0_row0" class="row_heading level0 row0" >A</th>
      <td id="T_e0b74_row0_col0" class="data row0 col0" >0</td>
      <td id="T_e0b74_row0_col1" class="data row0 col1" >0</td>
      <td id="T_e0b74_row0_col2" class="data row0 col2" >0</td>
      <td id="T_e0b74_row0_col3" class="data row0 col3" >0</td>
      <td id="T_e0b74_row0_col4" class="data row0 col4" >0</td>
      <td id="T_e0b74_row0_col5" class="data row0 col5" >0</td>
      <td id="T_e0b74_row0_col6" class="data row0 col6" >0</td>
      <td id="T_e0b74_row0_col7" class="data row0 col7" >0</td>
      <td id="T_e0b74_row0_col8" class="data row0 col8" >0</td>
      <td id="T_e0b74_row0_col9" class="data row0 col9" >0</td>
      <td id="T_e0b74_row0_col10" class="data row0 col10" >0</td>
      <td id="T_e0b74_row0_col11" class="data row0 col11" >0</td>
      <td id="T_e0b74_row0_col12" class="data row0 col12" >0</td>
      <td id="T_e0b74_row0_col13" class="data row0 col13" >0</td>
      <td id="T_e0b74_row0_col14" class="data row0 col14" >0</td>
      <td id="T_e0b74_row0_col15" class="data row0 col15" >0</td>
      <td id="T_e0b74_row0_col16" class="data row0 col16" >0</td>
      <td id="T_e0b74_row0_col17" class="data row0 col17" >0</td>
      <td id="T_e0b74_row0_col18" class="data row0 col18" >0</td>
      <td id="T_e0b74_row0_col19" class="data row0 col19" >0</td>
      <td id="T_e0b74_row0_col20" class="data row0 col20" >314,381</td>
      <td id="T_e0b74_row0_col21" class="data row0 col21" >0</td>
      <td id="T_e0b74_row0_col22" class="data row0 col22" >0</td>
      <td id="T_e0b74_row0_col23" class="data row0 col23" >358,998</td>
      <td id="T_e0b74_row0_col24" class="data row0 col24" >360,000</td>
      <td id="T_e0b74_row0_col25" class="data row0 col25" >0</td>
      <td id="T_e0b74_row0_col26" class="data row0 col26" >0</td>
      <td id="T_e0b74_row0_col27" class="data row0 col27" >11,250</td>
      <td id="T_e0b74_row0_col28" class="data row0 col28" >248,883</td>
      <td id="T_e0b74_row0_col29" class="data row0 col29" >0</td>
      <td id="T_e0b74_row0_col30" class="data row0 col30" >1,293,512</td>
    </tr>
    <tr>
      <th id="T_e0b74_level0_row1" class="row_heading level0 row1" >B</th>
      <td id="T_e0b74_row1_col0" class="data row1 col0" >0</td>
      <td id="T_e0b74_row1_col1" class="data row1 col1" >0</td>
      <td id="T_e0b74_row1_col2" class="data row1 col2" >0</td>
      <td id="T_e0b74_row1_col3" class="data row1 col3" >0</td>
      <td id="T_e0b74_row1_col4" class="data row1 col4" >0</td>
      <td id="T_e0b74_row1_col5" class="data row1 col5" >0</td>
      <td id="T_e0b74_row1_col6" class="data row1 col6" >0</td>
      <td id="T_e0b74_row1_col7" class="data row1 col7" >0</td>
      <td id="T_e0b74_row1_col8" class="data row1 col8" >0</td>
      <td id="T_e0b74_row1_col9" class="data row1 col9" >0</td>
      <td id="T_e0b74_row1_col10" class="data row1 col10" >0</td>
      <td id="T_e0b74_row1_col11" class="data row1 col11" >0</td>
      <td id="T_e0b74_row1_col12" class="data row1 col12" >0</td>
      <td id="T_e0b74_row1_col13" class="data row1 col13" >0</td>
      <td id="T_e0b74_row1_col14" class="data row1 col14" >0</td>
      <td id="T_e0b74_row1_col15" class="data row1 col15" >0</td>
      <td id="T_e0b74_row1_col16" class="data row1 col16" >0</td>
      <td id="T_e0b74_row1_col17" class="data row1 col17" >0</td>
      <td id="T_e0b74_row1_col18" class="data row1 col18" >0</td>
      <td id="T_e0b74_row1_col19" class="data row1 col19" >0</td>
      <td id="T_e0b74_row1_col20" class="data row1 col20" >0</td>
      <td id="T_e0b74_row1_col21" class="data row1 col21" >0</td>
      <td id="T_e0b74_row1_col22" class="data row1 col22" >0</td>
      <td id="T_e0b74_row1_col23" class="data row1 col23" >0</td>
      <td id="T_e0b74_row1_col24" class="data row1 col24" >0</td>
      <td id="T_e0b74_row1_col25" class="data row1 col25" >0</td>
      <td id="T_e0b74_row1_col26" class="data row1 col26" >0</td>
      <td id="T_e0b74_row1_col27" class="data row1 col27" >0</td>
      <td id="T_e0b74_row1_col28" class="data row1 col28" >0</td>
      <td id="T_e0b74_row1_col29" class="data row1 col29" >0</td>
      <td id="T_e0b74_row1_col30" class="data row1 col30" >0</td>
    </tr>
    <tr>
      <th id="T_e0b74_level0_row2" class="row_heading level0 row2" >C</th>
      <td id="T_e0b74_row2_col0" class="data row2 col0" >0</td>
      <td id="T_e0b74_row2_col1" class="data row2 col1" >0</td>
      <td id="T_e0b74_row2_col2" class="data row2 col2" >0</td>
      <td id="T_e0b74_row2_col3" class="data row2 col3" >0</td>
      <td id="T_e0b74_row2_col4" class="data row2 col4" >0</td>
      <td id="T_e0b74_row2_col5" class="data row2 col5" >0</td>
      <td id="T_e0b74_row2_col6" class="data row2 col6" >0</td>
      <td id="T_e0b74_row2_col7" class="data row2 col7" >0</td>
      <td id="T_e0b74_row2_col8" class="data row2 col8" >0</td>
      <td id="T_e0b74_row2_col9" class="data row2 col9" >0</td>
      <td id="T_e0b74_row2_col10" class="data row2 col10" >0</td>
      <td id="T_e0b74_row2_col11" class="data row2 col11" >0</td>
      <td id="T_e0b74_row2_col12" class="data row2 col12" >0</td>
      <td id="T_e0b74_row2_col13" class="data row2 col13" >0</td>
      <td id="T_e0b74_row2_col14" class="data row2 col14" >0</td>
      <td id="T_e0b74_row2_col15" class="data row2 col15" >0</td>
      <td id="T_e0b74_row2_col16" class="data row2 col16" >0</td>
      <td id="T_e0b74_row2_col17" class="data row2 col17" >0</td>
      <td id="T_e0b74_row2_col18" class="data row2 col18" >0</td>
      <td id="T_e0b74_row2_col19" class="data row2 col19" >0</td>
      <td id="T_e0b74_row2_col20" class="data row2 col20" >0</td>
      <td id="T_e0b74_row2_col21" class="data row2 col21" >0</td>
      <td id="T_e0b74_row2_col22" class="data row2 col22" >0</td>
      <td id="T_e0b74_row2_col23" class="data row2 col23" >0</td>
      <td id="T_e0b74_row2_col24" class="data row2 col24" >0</td>
      <td id="T_e0b74_row2_col25" class="data row2 col25" >0</td>
      <td id="T_e0b74_row2_col26" class="data row2 col26" >0</td>
      <td id="T_e0b74_row2_col27" class="data row2 col27" >0</td>
      <td id="T_e0b74_row2_col28" class="data row2 col28" >0</td>
      <td id="T_e0b74_row2_col29" class="data row2 col29" >0</td>
      <td id="T_e0b74_row2_col30" class="data row2 col30" >0</td>
    </tr>
    <tr>
      <th id="T_e0b74_level0_row3" class="row_heading level0 row3" >D</th>
      <td id="T_e0b74_row3_col0" class="data row3 col0" >0</td>
      <td id="T_e0b74_row3_col1" class="data row3 col1" >0</td>
      <td id="T_e0b74_row3_col2" class="data row3 col2" >0</td>
      <td id="T_e0b74_row3_col3" class="data row3 col3" >0</td>
      <td id="T_e0b74_row3_col4" class="data row3 col4" >0</td>
      <td id="T_e0b74_row3_col5" class="data row3 col5" >0</td>
      <td id="T_e0b74_row3_col6" class="data row3 col6" >0</td>
      <td id="T_e0b74_row3_col7" class="data row3 col7" >0</td>
      <td id="T_e0b74_row3_col8" class="data row3 col8" >0</td>
      <td id="T_e0b74_row3_col9" class="data row3 col9" >0</td>
      <td id="T_e0b74_row3_col10" class="data row3 col10" >0</td>
      <td id="T_e0b74_row3_col11" class="data row3 col11" >0</td>
      <td id="T_e0b74_row3_col12" class="data row3 col12" >0</td>
      <td id="T_e0b74_row3_col13" class="data row3 col13" >0</td>
      <td id="T_e0b74_row3_col14" class="data row3 col14" >0</td>
      <td id="T_e0b74_row3_col15" class="data row3 col15" >0</td>
      <td id="T_e0b74_row3_col16" class="data row3 col16" >0</td>
      <td id="T_e0b74_row3_col17" class="data row3 col17" >0</td>
      <td id="T_e0b74_row3_col18" class="data row3 col18" >0</td>
      <td id="T_e0b74_row3_col19" class="data row3 col19" >0</td>
      <td id="T_e0b74_row3_col20" class="data row3 col20" >0</td>
      <td id="T_e0b74_row3_col21" class="data row3 col21" >0</td>
      <td id="T_e0b74_row3_col22" class="data row3 col22" >0</td>
      <td id="T_e0b74_row3_col23" class="data row3 col23" >0</td>
      <td id="T_e0b74_row3_col24" class="data row3 col24" >0</td>
      <td id="T_e0b74_row3_col25" class="data row3 col25" >0</td>
      <td id="T_e0b74_row3_col26" class="data row3 col26" >0</td>
      <td id="T_e0b74_row3_col27" class="data row3 col27" >0</td>
      <td id="T_e0b74_row3_col28" class="data row3 col28" >0</td>
      <td id="T_e0b74_row3_col29" class="data row3 col29" >0</td>
      <td id="T_e0b74_row3_col30" class="data row3 col30" >0</td>
    </tr>
    <tr>
      <th id="T_e0b74_level0_row4" class="row_heading level0 row4" >E</th>
      <td id="T_e0b74_row4_col0" class="data row4 col0" >0</td>
      <td id="T_e0b74_row4_col1" class="data row4 col1" >0</td>
      <td id="T_e0b74_row4_col2" class="data row4 col2" >0</td>
      <td id="T_e0b74_row4_col3" class="data row4 col3" >0</td>
      <td id="T_e0b74_row4_col4" class="data row4 col4" >0</td>
      <td id="T_e0b74_row4_col5" class="data row4 col5" >0</td>
      <td id="T_e0b74_row4_col6" class="data row4 col6" >0</td>
      <td id="T_e0b74_row4_col7" class="data row4 col7" >0</td>
      <td id="T_e0b74_row4_col8" class="data row4 col8" >0</td>
      <td id="T_e0b74_row4_col9" class="data row4 col9" >0</td>
      <td id="T_e0b74_row4_col10" class="data row4 col10" >0</td>
      <td id="T_e0b74_row4_col11" class="data row4 col11" >0</td>
      <td id="T_e0b74_row4_col12" class="data row4 col12" >0</td>
      <td id="T_e0b74_row4_col13" class="data row4 col13" >0</td>
      <td id="T_e0b74_row4_col14" class="data row4 col14" >0</td>
      <td id="T_e0b74_row4_col15" class="data row4 col15" >359,999</td>
      <td id="T_e0b74_row4_col16" class="data row4 col16" >0</td>
      <td id="T_e0b74_row4_col17" class="data row4 col17" >0</td>
      <td id="T_e0b74_row4_col18" class="data row4 col18" >0</td>
      <td id="T_e0b74_row4_col19" class="data row4 col19" >0</td>
      <td id="T_e0b74_row4_col20" class="data row4 col20" >0</td>
      <td id="T_e0b74_row4_col21" class="data row4 col21" >326,162</td>
      <td id="T_e0b74_row4_col22" class="data row4 col22" >0</td>
      <td id="T_e0b74_row4_col23" class="data row4 col23" >0</td>
      <td id="T_e0b74_row4_col24" class="data row4 col24" >0</td>
      <td id="T_e0b74_row4_col25" class="data row4 col25" >0</td>
      <td id="T_e0b74_row4_col26" class="data row4 col26" >0</td>
      <td id="T_e0b74_row4_col27" class="data row4 col27" >0</td>
      <td id="T_e0b74_row4_col28" class="data row4 col28" >0</td>
      <td id="T_e0b74_row4_col29" class="data row4 col29" >336,948</td>
      <td id="T_e0b74_row4_col30" class="data row4 col30" >1,023,109</td>
    </tr>
    <tr>
      <th id="T_e0b74_level0_row5" class="row_heading level0 row5" >F</th>
      <td id="T_e0b74_row5_col0" class="data row5 col0" >0</td>
      <td id="T_e0b74_row5_col1" class="data row5 col1" >0</td>
      <td id="T_e0b74_row5_col2" class="data row5 col2" >0</td>
      <td id="T_e0b74_row5_col3" class="data row5 col3" >0</td>
      <td id="T_e0b74_row5_col4" class="data row5 col4" >0</td>
      <td id="T_e0b74_row5_col5" class="data row5 col5" >0</td>
      <td id="T_e0b74_row5_col6" class="data row5 col6" >0</td>
      <td id="T_e0b74_row5_col7" class="data row5 col7" >0</td>
      <td id="T_e0b74_row5_col8" class="data row5 col8" >0</td>
      <td id="T_e0b74_row5_col9" class="data row5 col9" >0</td>
      <td id="T_e0b74_row5_col10" class="data row5 col10" >0</td>
      <td id="T_e0b74_row5_col11" class="data row5 col11" >0</td>
      <td id="T_e0b74_row5_col12" class="data row5 col12" >0</td>
      <td id="T_e0b74_row5_col13" class="data row5 col13" >0</td>
      <td id="T_e0b74_row5_col14" class="data row5 col14" >0</td>
      <td id="T_e0b74_row5_col15" class="data row5 col15" >0</td>
      <td id="T_e0b74_row5_col16" class="data row5 col16" >0</td>
      <td id="T_e0b74_row5_col17" class="data row5 col17" >0</td>
      <td id="T_e0b74_row5_col18" class="data row5 col18" >0</td>
      <td id="T_e0b74_row5_col19" class="data row5 col19" >0</td>
      <td id="T_e0b74_row5_col20" class="data row5 col20" >0</td>
      <td id="T_e0b74_row5_col21" class="data row5 col21" >0</td>
      <td id="T_e0b74_row5_col22" class="data row5 col22" >0</td>
      <td id="T_e0b74_row5_col23" class="data row5 col23" >0</td>
      <td id="T_e0b74_row5_col24" class="data row5 col24" >0</td>
      <td id="T_e0b74_row5_col25" class="data row5 col25" >0</td>
      <td id="T_e0b74_row5_col26" class="data row5 col26" >0</td>
      <td id="T_e0b74_row5_col27" class="data row5 col27" >0</td>
      <td id="T_e0b74_row5_col28" class="data row5 col28" >2,000</td>
      <td id="T_e0b74_row5_col29" class="data row5 col29" >0</td>
      <td id="T_e0b74_row5_col30" class="data row5 col30" >2,000</td>
    </tr>
    <tr>
      <th id="T_e0b74_level0_row6" class="row_heading level0 row6" >G</th>
      <td id="T_e0b74_row6_col0" class="data row6 col0" >0</td>
      <td id="T_e0b74_row6_col1" class="data row6 col1" >0</td>
      <td id="T_e0b74_row6_col2" class="data row6 col2" >0</td>
      <td id="T_e0b74_row6_col3" class="data row6 col3" >0</td>
      <td id="T_e0b74_row6_col4" class="data row6 col4" >0</td>
      <td id="T_e0b74_row6_col5" class="data row6 col5" >0</td>
      <td id="T_e0b74_row6_col6" class="data row6 col6" >0</td>
      <td id="T_e0b74_row6_col7" class="data row6 col7" >134,500</td>
      <td id="T_e0b74_row6_col8" class="data row6 col8" >0</td>
      <td id="T_e0b74_row6_col9" class="data row6 col9" >0</td>
      <td id="T_e0b74_row6_col10" class="data row6 col10" >0</td>
      <td id="T_e0b74_row6_col11" class="data row6 col11" >0</td>
      <td id="T_e0b74_row6_col12" class="data row6 col12" >0</td>
      <td id="T_e0b74_row6_col13" class="data row6 col13" >0</td>
      <td id="T_e0b74_row6_col14" class="data row6 col14" >0</td>
      <td id="T_e0b74_row6_col15" class="data row6 col15" >0</td>
      <td id="T_e0b74_row6_col16" class="data row6 col16" >0</td>
      <td id="T_e0b74_row6_col17" class="data row6 col17" >0</td>
      <td id="T_e0b74_row6_col18" class="data row6 col18" >0</td>
      <td id="T_e0b74_row6_col19" class="data row6 col19" >0</td>
      <td id="T_e0b74_row6_col20" class="data row6 col20" >0</td>
      <td id="T_e0b74_row6_col21" class="data row6 col21" >0</td>
      <td id="T_e0b74_row6_col22" class="data row6 col22" >0</td>
      <td id="T_e0b74_row6_col23" class="data row6 col23" >0</td>
      <td id="T_e0b74_row6_col24" class="data row6 col24" >0</td>
      <td id="T_e0b74_row6_col25" class="data row6 col25" >0</td>
      <td id="T_e0b74_row6_col26" class="data row6 col26" >0</td>
      <td id="T_e0b74_row6_col27" class="data row6 col27" >0</td>
      <td id="T_e0b74_row6_col28" class="data row6 col28" >0</td>
      <td id="T_e0b74_row6_col29" class="data row6 col29" >0</td>
      <td id="T_e0b74_row6_col30" class="data row6 col30" >134,500</td>
    </tr>
    <tr>
      <th id="T_e0b74_level0_row7" class="row_heading level0 row7" >H</th>
      <td id="T_e0b74_row7_col0" class="data row7 col0" >0</td>
      <td id="T_e0b74_row7_col1" class="data row7 col1" >0</td>
      <td id="T_e0b74_row7_col2" class="data row7 col2" >0</td>
      <td id="T_e0b74_row7_col3" class="data row7 col3" >0</td>
      <td id="T_e0b74_row7_col4" class="data row7 col4" >0</td>
      <td id="T_e0b74_row7_col5" class="data row7 col5" >0</td>
      <td id="T_e0b74_row7_col6" class="data row7 col6" >0</td>
      <td id="T_e0b74_row7_col7" class="data row7 col7" >0</td>
      <td id="T_e0b74_row7_col8" class="data row7 col8" >0</td>
      <td id="T_e0b74_row7_col9" class="data row7 col9" >180,000</td>
      <td id="T_e0b74_row7_col10" class="data row7 col10" >0</td>
      <td id="T_e0b74_row7_col11" class="data row7 col11" >0</td>
      <td id="T_e0b74_row7_col12" class="data row7 col12" >0</td>
      <td id="T_e0b74_row7_col13" class="data row7 col13" >0</td>
      <td id="T_e0b74_row7_col14" class="data row7 col14" >0</td>
      <td id="T_e0b74_row7_col15" class="data row7 col15" >0</td>
      <td id="T_e0b74_row7_col16" class="data row7 col16" >0</td>
      <td id="T_e0b74_row7_col17" class="data row7 col17" >0</td>
      <td id="T_e0b74_row7_col18" class="data row7 col18" >0</td>
      <td id="T_e0b74_row7_col19" class="data row7 col19" >0</td>
      <td id="T_e0b74_row7_col20" class="data row7 col20" >45,619</td>
      <td id="T_e0b74_row7_col21" class="data row7 col21" >32,636</td>
      <td id="T_e0b74_row7_col22" class="data row7 col22" >9,564</td>
      <td id="T_e0b74_row7_col23" class="data row7 col23" >1,000</td>
      <td id="T_e0b74_row7_col24" class="data row7 col24" >0</td>
      <td id="T_e0b74_row7_col25" class="data row7 col25" >0</td>
      <td id="T_e0b74_row7_col26" class="data row7 col26" >0</td>
      <td id="T_e0b74_row7_col27" class="data row7 col27" >0</td>
      <td id="T_e0b74_row7_col28" class="data row7 col28" >8,079</td>
      <td id="T_e0b74_row7_col29" class="data row7 col29" >21,052</td>
      <td id="T_e0b74_row7_col30" class="data row7 col30" >297,950</td>
    </tr>
    <tr>
      <th id="T_e0b74_level0_row8" class="row_heading level0 row8" >I</th>
      <td id="T_e0b74_row8_col0" class="data row8 col0" >0</td>
      <td id="T_e0b74_row8_col1" class="data row8 col1" >0</td>
      <td id="T_e0b74_row8_col2" class="data row8 col2" >0</td>
      <td id="T_e0b74_row8_col3" class="data row8 col3" >0</td>
      <td id="T_e0b74_row8_col4" class="data row8 col4" >0</td>
      <td id="T_e0b74_row8_col5" class="data row8 col5" >0</td>
      <td id="T_e0b74_row8_col6" class="data row8 col6" >0</td>
      <td id="T_e0b74_row8_col7" class="data row8 col7" >0</td>
      <td id="T_e0b74_row8_col8" class="data row8 col8" >0</td>
      <td id="T_e0b74_row8_col9" class="data row8 col9" >0</td>
      <td id="T_e0b74_row8_col10" class="data row8 col10" >0</td>
      <td id="T_e0b74_row8_col11" class="data row8 col11" >0</td>
      <td id="T_e0b74_row8_col12" class="data row8 col12" >0</td>
      <td id="T_e0b74_row8_col13" class="data row8 col13" >0</td>
      <td id="T_e0b74_row8_col14" class="data row8 col14" >0</td>
      <td id="T_e0b74_row8_col15" class="data row8 col15" >0</td>
      <td id="T_e0b74_row8_col16" class="data row8 col16" >0</td>
      <td id="T_e0b74_row8_col17" class="data row8 col17" >0</td>
      <td id="T_e0b74_row8_col18" class="data row8 col18" >0</td>
      <td id="T_e0b74_row8_col19" class="data row8 col19" >0</td>
      <td id="T_e0b74_row8_col20" class="data row8 col20" >0</td>
      <td id="T_e0b74_row8_col21" class="data row8 col21" >0</td>
      <td id="T_e0b74_row8_col22" class="data row8 col22" >0</td>
      <td id="T_e0b74_row8_col23" class="data row8 col23" >0</td>
      <td id="T_e0b74_row8_col24" class="data row8 col24" >0</td>
      <td id="T_e0b74_row8_col25" class="data row8 col25" >0</td>
      <td id="T_e0b74_row8_col26" class="data row8 col26" >0</td>
      <td id="T_e0b74_row8_col27" class="data row8 col27" >0</td>
      <td id="T_e0b74_row8_col28" class="data row8 col28" >0</td>
      <td id="T_e0b74_row8_col29" class="data row8 col29" >2,000</td>
      <td id="T_e0b74_row8_col30" class="data row8 col30" >2,000</td>
    </tr>
  </tbody>
</table>


</details>

<!-- END TABLE:ot-finished-output -->

<!-- BEGIN TABLE:material-orders -->

<details>
<summary>Material orders — grouped by material, then order day</summary>

<style type="text/css">
#T_d357e th.col_heading {
  background-color: #D9E1F2;
  color: #000000;
  font-weight: bold;
  padding: 6px;
  border: 1px solid #A6A6A6;
}
#T_d357e th.row_heading.level0 {
  background-color: #FFD966;
  color: #000000;
  font-weight: bold;
  padding: 6px;
  border-right: 2px solid #7F6000;
}
#T_d357e td.row0 {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_d357e  th.row_heading.row0:not(.level0) {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_d357e .row0 {
  border-top: 3px solid #595959;
}
#T_d357e td.row1 {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_d357e  th.row_heading.row1:not(.level0) {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_d357e td.row2 {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_d357e  th.row_heading.row2:not(.level0) {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_d357e td.row3 {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_d357e  th.row_heading.row3:not(.level0) {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_d357e td.row4 {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_d357e  th.row_heading.row4:not(.level0) {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_d357e td.row5 {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_d357e  th.row_heading.row5:not(.level0) {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_d357e td.row6 {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_d357e  th.row_heading.row6:not(.level0) {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_d357e td.row7 {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_d357e  th.row_heading.row7:not(.level0) {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_d357e td.row8 {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_d357e  th.row_heading.row8:not(.level0) {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_d357e td.row9 {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_d357e  th.row_heading.row9:not(.level0) {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_d357e td.row10 {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_d357e  th.row_heading.row10:not(.level0) {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_d357e td.row11 {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_d357e  th.row_heading.row11:not(.level0) {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_d357e td.row12 {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_d357e  th.row_heading.row12:not(.level0) {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_d357e .row12 {
  border-top: 3px solid #595959;
}
#T_d357e td.row13 {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_d357e  th.row_heading.row13:not(.level0) {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_d357e td.row14 {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_d357e  th.row_heading.row14:not(.level0) {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_d357e .row14 {
  border-top: 3px solid #595959;
}
#T_d357e td.row15 {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_d357e  th.row_heading.row15:not(.level0) {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_d357e .row15 {
  border-top: 3px solid #595959;
}
#T_d357e td.row16 {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_d357e  th.row_heading.row16:not(.level0) {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
</style>
<table id="T_d357e">
  <thead>
    <tr>
      <th class="blank" >&nbsp;</th>
      <th class="blank level0" >&nbsp;</th>
      <th id="T_d357e_level0_col0" class="col_heading level0 col0" >Arrival Day</th>
      <th id="T_d357e_level0_col1" class="col_heading level0 col1" >Order Quantity</th>
      <th id="T_d357e_level0_col2" class="col_heading level0 col2" >Containers</th>
      <th id="T_d357e_level0_col3" class="col_heading level0 col3" >Tier</th>
      <th id="T_d357e_level0_col4" class="col_heading level0 col4" >Tier Lower</th>
      <th id="T_d357e_level0_col5" class="col_heading level0 col5" >Tier Upper</th>
      <th id="T_d357e_level0_col6" class="col_heading level0 col6" >Discount</th>
      <th id="T_d357e_level0_col7" class="col_heading level0 col7" >Original Unit Price</th>
      <th id="T_d357e_level0_col8" class="col_heading level0 col8" >Discounted Unit Price</th>
      <th id="T_d357e_level0_col9" class="col_heading level0 col9" >Purchase Subtotal</th>
    </tr>
    <tr>
      <th class="index_name level0" >Material</th>
      <th class="index_name level1" >Order Day</th>
      <th class="blank col0" >&nbsp;</th>
      <th class="blank col1" >&nbsp;</th>
      <th class="blank col2" >&nbsp;</th>
      <th class="blank col3" >&nbsp;</th>
      <th class="blank col4" >&nbsp;</th>
      <th class="blank col5" >&nbsp;</th>
      <th class="blank col6" >&nbsp;</th>
      <th class="blank col7" >&nbsp;</th>
      <th class="blank col8" >&nbsp;</th>
      <th class="blank col9" >&nbsp;</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <th id="T_d357e_level0_row0" class="row_heading level0 row0" rowspan="12">Paper</th>
      <th id="T_d357e_level1_row0" class="row_heading level1 row0" >1</th>
      <td id="T_d357e_row0_col0" class="data row0 col0" >7</td>
      <td id="T_d357e_row0_col1" class="data row0 col1" >12,000.00</td>
      <td id="T_d357e_row0_col2" class="data row0 col2" >4</td>
      <td id="T_d357e_row0_col3" class="data row0 col3" >T1/1</td>
      <td id="T_d357e_row0_col4" class="data row0 col4" >9,800.00</td>
      <td id="T_d357e_row0_col5" class="data row0 col5" >14,000.00</td>
      <td id="T_d357e_row0_col6" class="data row0 col6" >1.0%</td>
      <td id="T_d357e_row0_col7" class="data row0 col7" >300.00</td>
      <td id="T_d357e_row0_col8" class="data row0 col8" >297.00</td>
      <td id="T_d357e_row0_col9" class="data row0 col9" >3,564,000.00</td>
    </tr>
    <tr>
      <th id="T_d357e_level1_row1" class="row_heading level1 row1" >2</th>
      <td id="T_d357e_row1_col0" class="data row1 col0" >8</td>
      <td id="T_d357e_row1_col1" class="data row1 col1" >12,000.00</td>
      <td id="T_d357e_row1_col2" class="data row1 col2" >4</td>
      <td id="T_d357e_row1_col3" class="data row1 col3" >T1/1</td>
      <td id="T_d357e_row1_col4" class="data row1 col4" >9,800.00</td>
      <td id="T_d357e_row1_col5" class="data row1 col5" >14,000.00</td>
      <td id="T_d357e_row1_col6" class="data row1 col6" >1.0%</td>
      <td id="T_d357e_row1_col7" class="data row1 col7" >300.00</td>
      <td id="T_d357e_row1_col8" class="data row1 col8" >297.00</td>
      <td id="T_d357e_row1_col9" class="data row1 col9" >3,564,000.00</td>
    </tr>
    <tr>
      <th id="T_d357e_level1_row2" class="row_heading level1 row2" >3</th>
      <td id="T_d357e_row2_col0" class="data row2 col0" >9</td>
      <td id="T_d357e_row2_col1" class="data row2 col1" >12,000.00</td>
      <td id="T_d357e_row2_col2" class="data row2 col2" >4</td>
      <td id="T_d357e_row2_col3" class="data row2 col3" >T1/1</td>
      <td id="T_d357e_row2_col4" class="data row2 col4" >9,800.00</td>
      <td id="T_d357e_row2_col5" class="data row2 col5" >14,000.00</td>
      <td id="T_d357e_row2_col6" class="data row2 col6" >1.0%</td>
      <td id="T_d357e_row2_col7" class="data row2 col7" >300.00</td>
      <td id="T_d357e_row2_col8" class="data row2 col8" >297.00</td>
      <td id="T_d357e_row2_col9" class="data row2 col9" >3,564,000.00</td>
    </tr>
    <tr>
      <th id="T_d357e_level1_row3" class="row_heading level1 row3" >4</th>
      <td id="T_d357e_row3_col0" class="data row3 col0" >10</td>
      <td id="T_d357e_row3_col1" class="data row3 col1" >12,000.00</td>
      <td id="T_d357e_row3_col2" class="data row3 col2" >4</td>
      <td id="T_d357e_row3_col3" class="data row3 col3" >T1/1</td>
      <td id="T_d357e_row3_col4" class="data row3 col4" >9,800.00</td>
      <td id="T_d357e_row3_col5" class="data row3 col5" >14,000.00</td>
      <td id="T_d357e_row3_col6" class="data row3 col6" >1.0%</td>
      <td id="T_d357e_row3_col7" class="data row3 col7" >300.00</td>
      <td id="T_d357e_row3_col8" class="data row3 col8" >297.00</td>
      <td id="T_d357e_row3_col9" class="data row3 col9" >3,564,000.00</td>
    </tr>
    <tr>
      <th id="T_d357e_level1_row4" class="row_heading level1 row4" >5</th>
      <td id="T_d357e_row4_col0" class="data row4 col0" >11</td>
      <td id="T_d357e_row4_col1" class="data row4 col1" >12,000.00</td>
      <td id="T_d357e_row4_col2" class="data row4 col2" >4</td>
      <td id="T_d357e_row4_col3" class="data row4 col3" >T1/1</td>
      <td id="T_d357e_row4_col4" class="data row4 col4" >9,800.00</td>
      <td id="T_d357e_row4_col5" class="data row4 col5" >14,000.00</td>
      <td id="T_d357e_row4_col6" class="data row4 col6" >1.0%</td>
      <td id="T_d357e_row4_col7" class="data row4 col7" >300.00</td>
      <td id="T_d357e_row4_col8" class="data row4 col8" >297.00</td>
      <td id="T_d357e_row4_col9" class="data row4 col9" >3,564,000.00</td>
    </tr>
    <tr>
      <th id="T_d357e_level1_row5" class="row_heading level1 row5" >6</th>
      <td id="T_d357e_row5_col0" class="data row5 col0" >12</td>
      <td id="T_d357e_row5_col1" class="data row5 col1" >12,000.00</td>
      <td id="T_d357e_row5_col2" class="data row5 col2" >4</td>
      <td id="T_d357e_row5_col3" class="data row5 col3" >T1/1</td>
      <td id="T_d357e_row5_col4" class="data row5 col4" >9,800.00</td>
      <td id="T_d357e_row5_col5" class="data row5 col5" >14,000.00</td>
      <td id="T_d357e_row5_col6" class="data row5 col6" >1.0%</td>
      <td id="T_d357e_row5_col7" class="data row5 col7" >300.00</td>
      <td id="T_d357e_row5_col8" class="data row5 col8" >297.00</td>
      <td id="T_d357e_row5_col9" class="data row5 col9" >3,564,000.00</td>
    </tr>
    <tr>
      <th id="T_d357e_level1_row6" class="row_heading level1 row6" >7</th>
      <td id="T_d357e_row6_col0" class="data row6 col0" >13</td>
      <td id="T_d357e_row6_col1" class="data row6 col1" >12,000.00</td>
      <td id="T_d357e_row6_col2" class="data row6 col2" >4</td>
      <td id="T_d357e_row6_col3" class="data row6 col3" >T1/1</td>
      <td id="T_d357e_row6_col4" class="data row6 col4" >9,800.00</td>
      <td id="T_d357e_row6_col5" class="data row6 col5" >14,000.00</td>
      <td id="T_d357e_row6_col6" class="data row6 col6" >1.0%</td>
      <td id="T_d357e_row6_col7" class="data row6 col7" >300.00</td>
      <td id="T_d357e_row6_col8" class="data row6 col8" >297.00</td>
      <td id="T_d357e_row6_col9" class="data row6 col9" >3,564,000.00</td>
    </tr>
    <tr>
      <th id="T_d357e_level1_row7" class="row_heading level1 row7" >8</th>
      <td id="T_d357e_row7_col0" class="data row7 col0" >14</td>
      <td id="T_d357e_row7_col1" class="data row7 col1" >12,000.00</td>
      <td id="T_d357e_row7_col2" class="data row7 col2" >4</td>
      <td id="T_d357e_row7_col3" class="data row7 col3" >T1/1</td>
      <td id="T_d357e_row7_col4" class="data row7 col4" >9,800.00</td>
      <td id="T_d357e_row7_col5" class="data row7 col5" >14,000.00</td>
      <td id="T_d357e_row7_col6" class="data row7 col6" >1.0%</td>
      <td id="T_d357e_row7_col7" class="data row7 col7" >300.00</td>
      <td id="T_d357e_row7_col8" class="data row7 col8" >297.00</td>
      <td id="T_d357e_row7_col9" class="data row7 col9" >3,564,000.00</td>
    </tr>
    <tr>
      <th id="T_d357e_level1_row8" class="row_heading level1 row8" >12</th>
      <td id="T_d357e_row8_col0" class="data row8 col0" >18</td>
      <td id="T_d357e_row8_col1" class="data row8 col1" >12,000.00</td>
      <td id="T_d357e_row8_col2" class="data row8 col2" >4</td>
      <td id="T_d357e_row8_col3" class="data row8 col3" >T1/1</td>
      <td id="T_d357e_row8_col4" class="data row8 col4" >9,800.00</td>
      <td id="T_d357e_row8_col5" class="data row8 col5" >14,000.00</td>
      <td id="T_d357e_row8_col6" class="data row8 col6" >1.0%</td>
      <td id="T_d357e_row8_col7" class="data row8 col7" >300.00</td>
      <td id="T_d357e_row8_col8" class="data row8 col8" >297.00</td>
      <td id="T_d357e_row8_col9" class="data row8 col9" >3,564,000.00</td>
    </tr>
    <tr>
      <th id="T_d357e_level1_row9" class="row_heading level1 row9" >14</th>
      <td id="T_d357e_row9_col0" class="data row9 col0" >20</td>
      <td id="T_d357e_row9_col1" class="data row9 col1" >2,810.81</td>
      <td id="T_d357e_row9_col2" class="data row9 col2" >1</td>
      <td id="T_d357e_row9_col3" class="data row9 col3" >T0/1</td>
      <td id="T_d357e_row9_col4" class="data row9 col4" >0.00</td>
      <td id="T_d357e_row9_col5" class="data row9 col5" >9,800.00</td>
      <td id="T_d357e_row9_col6" class="data row9 col6" >0.0%</td>
      <td id="T_d357e_row9_col7" class="data row9 col7" >300.00</td>
      <td id="T_d357e_row9_col8" class="data row9 col8" >300.00</td>
      <td id="T_d357e_row9_col9" class="data row9 col9" >843,243.00</td>
    </tr>
    <tr>
      <th id="T_d357e_level1_row10" class="row_heading level1 row10" >19</th>
      <td id="T_d357e_row10_col0" class="data row10 col0" >25</td>
      <td id="T_d357e_row10_col1" class="data row10 col1" >9,000.00</td>
      <td id="T_d357e_row10_col2" class="data row10 col2" >3</td>
      <td id="T_d357e_row10_col3" class="data row10 col3" >T0/1</td>
      <td id="T_d357e_row10_col4" class="data row10 col4" >0.00</td>
      <td id="T_d357e_row10_col5" class="data row10 col5" >9,800.00</td>
      <td id="T_d357e_row10_col6" class="data row10 col6" >0.0%</td>
      <td id="T_d357e_row10_col7" class="data row10 col7" >300.00</td>
      <td id="T_d357e_row10_col8" class="data row10 col8" >300.00</td>
      <td id="T_d357e_row10_col9" class="data row10 col9" >2,700,000.00</td>
    </tr>
    <tr>
      <th id="T_d357e_level1_row11" class="row_heading level1 row11" >22</th>
      <td id="T_d357e_row11_col0" class="data row11 col0" >28</td>
      <td id="T_d357e_row11_col1" class="data row11 col1" >12,000.00</td>
      <td id="T_d357e_row11_col2" class="data row11 col2" >4</td>
      <td id="T_d357e_row11_col3" class="data row11 col3" >T1/1</td>
      <td id="T_d357e_row11_col4" class="data row11 col4" >9,800.00</td>
      <td id="T_d357e_row11_col5" class="data row11 col5" >14,000.00</td>
      <td id="T_d357e_row11_col6" class="data row11 col6" >1.0%</td>
      <td id="T_d357e_row11_col7" class="data row11 col7" >300.00</td>
      <td id="T_d357e_row11_col8" class="data row11 col8" >297.00</td>
      <td id="T_d357e_row11_col9" class="data row11 col9" >3,564,000.00</td>
    </tr>
    <tr>
      <th id="T_d357e_level0_row12" class="row_heading level0 row12" rowspan="2">String</th>
      <th id="T_d357e_level1_row12" class="row_heading level1 row12" >3</th>
      <td id="T_d357e_row12_col0" class="data row12 col0" >7</td>
      <td id="T_d357e_row12_col1" class="data row12 col1" >25,000.00</td>
      <td id="T_d357e_row12_col2" class="data row12 col2" >5</td>
      <td id="T_d357e_row12_col3" class="data row12 col3" >T1/1</td>
      <td id="T_d357e_row12_col4" class="data row12 col4" >14,000.00</td>
      <td id="T_d357e_row12_col5" class="data row12 col5" >28,000.00</td>
      <td id="T_d357e_row12_col6" class="data row12 col6" >2.0%</td>
      <td id="T_d357e_row12_col7" class="data row12 col7" >20.00</td>
      <td id="T_d357e_row12_col8" class="data row12 col8" >19.60</td>
      <td id="T_d357e_row12_col9" class="data row12 col9" >490,000.00</td>
    </tr>
    <tr>
      <th id="T_d357e_level1_row13" class="row_heading level1 row13" >24</th>
      <td id="T_d357e_row13_col0" class="data row13 col0" >28</td>
      <td id="T_d357e_row13_col1" class="data row13 col1" >3,719.96</td>
      <td id="T_d357e_row13_col2" class="data row13 col2" >1</td>
      <td id="T_d357e_row13_col3" class="data row13 col3" >T0/1</td>
      <td id="T_d357e_row13_col4" class="data row13 col4" >0.00</td>
      <td id="T_d357e_row13_col5" class="data row13 col5" >14,000.00</td>
      <td id="T_d357e_row13_col6" class="data row13 col6" >0.0%</td>
      <td id="T_d357e_row13_col7" class="data row13 col7" >20.00</td>
      <td id="T_d357e_row13_col8" class="data row13 col8" >20.00</td>
      <td id="T_d357e_row13_col9" class="data row13 col9" >74,399.24</td>
    </tr>
    <tr>
      <th id="T_d357e_level0_row14" class="row_heading level0 row14" >Glue</th>
      <th id="T_d357e_level1_row14" class="row_heading level1 row14" >1</th>
      <td id="T_d357e_row14_col0" class="data row14 col0" >13</td>
      <td id="T_d357e_row14_col1" class="data row14 col1" >660.89</td>
      <td id="T_d357e_row14_col2" class="data row14 col2" >1</td>
      <td id="T_d357e_row14_col3" class="data row14 col3" >T0/2</td>
      <td id="T_d357e_row14_col4" class="data row14 col4" >0.00</td>
      <td id="T_d357e_row14_col5" class="data row14 col5" >4,000.00</td>
      <td id="T_d357e_row14_col6" class="data row14 col6" >1.0%</td>
      <td id="T_d357e_row14_col7" class="data row14 col7" >40.00</td>
      <td id="T_d357e_row14_col8" class="data row14 col8" >39.60</td>
      <td id="T_d357e_row14_col9" class="data row14 col9" >26,171.28</td>
    </tr>
    <tr>
      <th id="T_d357e_level0_row15" class="row_heading level0 row15" rowspan="2">Poly Etylen</th>
      <th id="T_d357e_level1_row15" class="row_heading level1 row15" >2</th>
      <td id="T_d357e_row15_col0" class="data row15 col0" >8</td>
      <td id="T_d357e_row15_col1" class="data row15 col1" >10,000.00</td>
      <td id="T_d357e_row15_col2" class="data row15 col2" >5</td>
      <td id="T_d357e_row15_col3" class="data row15 col3" >T1/1</td>
      <td id="T_d357e_row15_col4" class="data row15 col4" >9,800.00</td>
      <td id="T_d357e_row15_col5" class="data row15 col5" >14,000.00</td>
      <td id="T_d357e_row15_col6" class="data row15 col6" >1.0%</td>
      <td id="T_d357e_row15_col7" class="data row15 col7" >500.00</td>
      <td id="T_d357e_row15_col8" class="data row15 col8" >495.00</td>
      <td id="T_d357e_row15_col9" class="data row15 col9" >4,950,000.00</td>
    </tr>
    <tr>
      <th id="T_d357e_level1_row16" class="row_heading level1 row16" >8</th>
      <td id="T_d357e_row16_col0" class="data row16 col0" >14</td>
      <td id="T_d357e_row16_col1" class="data row16 col1" >10,110.00</td>
      <td id="T_d357e_row16_col2" class="data row16 col2" >6</td>
      <td id="T_d357e_row16_col3" class="data row16 col3" >T1/1</td>
      <td id="T_d357e_row16_col4" class="data row16 col4" >9,800.00</td>
      <td id="T_d357e_row16_col5" class="data row16 col5" >14,000.00</td>
      <td id="T_d357e_row16_col6" class="data row16 col6" >1.0%</td>
      <td id="T_d357e_row16_col7" class="data row16 col7" >500.00</td>
      <td id="T_d357e_row16_col8" class="data row16 col8" >495.00</td>
      <td id="T_d357e_row16_col9" class="data row16 col9" >5,004,450.00</td>
    </tr>
  </tbody>
</table>


</details>

<details>
<summary>Material orders — grouped by order day, then material</summary>

<style type="text/css">
#T_c148e th.col_heading {
  background-color: #D9E1F2;
  color: #000000;
  font-weight: bold;
  padding: 6px;
  border: 1px solid #A6A6A6;
}
#T_c148e th.row_heading.level0 {
  background-color: #FFD966;
  color: #000000;
  font-weight: bold;
  padding: 6px;
  border-right: 2px solid #7F6000;
}
#T_c148e td.row0 {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_c148e  th.row_heading.row0:not(.level0) {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_c148e .row0 {
  border-top: 3px solid #595959;
}
#T_c148e td.row1 {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_c148e  th.row_heading.row1:not(.level0) {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_c148e td.row2 {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_c148e  th.row_heading.row2:not(.level0) {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_c148e .row2 {
  border-top: 3px solid #595959;
}
#T_c148e td.row3 {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_c148e  th.row_heading.row3:not(.level0) {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_c148e td.row4 {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_c148e  th.row_heading.row4:not(.level0) {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_c148e .row4 {
  border-top: 3px solid #595959;
}
#T_c148e td.row5 {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_c148e  th.row_heading.row5:not(.level0) {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_c148e td.row6 {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_c148e  th.row_heading.row6:not(.level0) {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_c148e .row6 {
  border-top: 3px solid #595959;
}
#T_c148e td.row7 {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_c148e  th.row_heading.row7:not(.level0) {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_c148e .row7 {
  border-top: 3px solid #595959;
}
#T_c148e td.row8 {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_c148e  th.row_heading.row8:not(.level0) {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_c148e .row8 {
  border-top: 3px solid #595959;
}
#T_c148e td.row9 {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_c148e  th.row_heading.row9:not(.level0) {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_c148e .row9 {
  border-top: 3px solid #595959;
}
#T_c148e td.row10 {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_c148e  th.row_heading.row10:not(.level0) {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_c148e .row10 {
  border-top: 3px solid #595959;
}
#T_c148e td.row11 {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_c148e  th.row_heading.row11:not(.level0) {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_c148e td.row12 {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_c148e  th.row_heading.row12:not(.level0) {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_c148e .row12 {
  border-top: 3px solid #595959;
}
#T_c148e td.row13 {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_c148e  th.row_heading.row13:not(.level0) {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_c148e .row13 {
  border-top: 3px solid #595959;
}
#T_c148e td.row14 {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_c148e  th.row_heading.row14:not(.level0) {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_c148e .row14 {
  border-top: 3px solid #595959;
}
#T_c148e td.row15 {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_c148e  th.row_heading.row15:not(.level0) {
  background-color: #DDEBF7;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_c148e .row15 {
  border-top: 3px solid #595959;
}
#T_c148e td.row16 {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_c148e  th.row_heading.row16:not(.level0) {
  background-color: #FFFFFF;
  color: #000000;
  padding: 6px;
  border-bottom: 1px solid #D9D9D9;
}
#T_c148e .row16 {
  border-top: 3px solid #595959;
}
</style>
<table id="T_c148e">
  <thead>
    <tr>
      <th class="blank" >&nbsp;</th>
      <th class="blank level0" >&nbsp;</th>
      <th id="T_c148e_level0_col0" class="col_heading level0 col0" >Arrival Day</th>
      <th id="T_c148e_level0_col1" class="col_heading level0 col1" >Order Quantity</th>
      <th id="T_c148e_level0_col2" class="col_heading level0 col2" >Containers</th>
      <th id="T_c148e_level0_col3" class="col_heading level0 col3" >Tier</th>
      <th id="T_c148e_level0_col4" class="col_heading level0 col4" >Tier Lower</th>
      <th id="T_c148e_level0_col5" class="col_heading level0 col5" >Tier Upper</th>
      <th id="T_c148e_level0_col6" class="col_heading level0 col6" >Discount</th>
      <th id="T_c148e_level0_col7" class="col_heading level0 col7" >Original Unit Price</th>
      <th id="T_c148e_level0_col8" class="col_heading level0 col8" >Discounted Unit Price</th>
      <th id="T_c148e_level0_col9" class="col_heading level0 col9" >Purchase Subtotal</th>
    </tr>
    <tr>
      <th class="index_name level0" >Order Day</th>
      <th class="index_name level1" >Material</th>
      <th class="blank col0" >&nbsp;</th>
      <th class="blank col1" >&nbsp;</th>
      <th class="blank col2" >&nbsp;</th>
      <th class="blank col3" >&nbsp;</th>
      <th class="blank col4" >&nbsp;</th>
      <th class="blank col5" >&nbsp;</th>
      <th class="blank col6" >&nbsp;</th>
      <th class="blank col7" >&nbsp;</th>
      <th class="blank col8" >&nbsp;</th>
      <th class="blank col9" >&nbsp;</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <th id="T_c148e_level0_row0" class="row_heading level0 row0" rowspan="2">1</th>
      <th id="T_c148e_level1_row0" class="row_heading level1 row0" >Paper</th>
      <td id="T_c148e_row0_col0" class="data row0 col0" >7</td>
      <td id="T_c148e_row0_col1" class="data row0 col1" >12,000.00</td>
      <td id="T_c148e_row0_col2" class="data row0 col2" >4</td>
      <td id="T_c148e_row0_col3" class="data row0 col3" >T1/1</td>
      <td id="T_c148e_row0_col4" class="data row0 col4" >9,800.00</td>
      <td id="T_c148e_row0_col5" class="data row0 col5" >14,000.00</td>
      <td id="T_c148e_row0_col6" class="data row0 col6" >1.0%</td>
      <td id="T_c148e_row0_col7" class="data row0 col7" >300.00</td>
      <td id="T_c148e_row0_col8" class="data row0 col8" >297.00</td>
      <td id="T_c148e_row0_col9" class="data row0 col9" >3,564,000.00</td>
    </tr>
    <tr>
      <th id="T_c148e_level1_row1" class="row_heading level1 row1" >Glue</th>
      <td id="T_c148e_row1_col0" class="data row1 col0" >13</td>
      <td id="T_c148e_row1_col1" class="data row1 col1" >660.89</td>
      <td id="T_c148e_row1_col2" class="data row1 col2" >1</td>
      <td id="T_c148e_row1_col3" class="data row1 col3" >T0/2</td>
      <td id="T_c148e_row1_col4" class="data row1 col4" >0.00</td>
      <td id="T_c148e_row1_col5" class="data row1 col5" >4,000.00</td>
      <td id="T_c148e_row1_col6" class="data row1 col6" >1.0%</td>
      <td id="T_c148e_row1_col7" class="data row1 col7" >40.00</td>
      <td id="T_c148e_row1_col8" class="data row1 col8" >39.60</td>
      <td id="T_c148e_row1_col9" class="data row1 col9" >26,171.28</td>
    </tr>
    <tr>
      <th id="T_c148e_level0_row2" class="row_heading level0 row2" rowspan="2">2</th>
      <th id="T_c148e_level1_row2" class="row_heading level1 row2" >Paper</th>
      <td id="T_c148e_row2_col0" class="data row2 col0" >8</td>
      <td id="T_c148e_row2_col1" class="data row2 col1" >12,000.00</td>
      <td id="T_c148e_row2_col2" class="data row2 col2" >4</td>
      <td id="T_c148e_row2_col3" class="data row2 col3" >T1/1</td>
      <td id="T_c148e_row2_col4" class="data row2 col4" >9,800.00</td>
      <td id="T_c148e_row2_col5" class="data row2 col5" >14,000.00</td>
      <td id="T_c148e_row2_col6" class="data row2 col6" >1.0%</td>
      <td id="T_c148e_row2_col7" class="data row2 col7" >300.00</td>
      <td id="T_c148e_row2_col8" class="data row2 col8" >297.00</td>
      <td id="T_c148e_row2_col9" class="data row2 col9" >3,564,000.00</td>
    </tr>
    <tr>
      <th id="T_c148e_level1_row3" class="row_heading level1 row3" >Poly Etylen</th>
      <td id="T_c148e_row3_col0" class="data row3 col0" >8</td>
      <td id="T_c148e_row3_col1" class="data row3 col1" >10,000.00</td>
      <td id="T_c148e_row3_col2" class="data row3 col2" >5</td>
      <td id="T_c148e_row3_col3" class="data row3 col3" >T1/1</td>
      <td id="T_c148e_row3_col4" class="data row3 col4" >9,800.00</td>
      <td id="T_c148e_row3_col5" class="data row3 col5" >14,000.00</td>
      <td id="T_c148e_row3_col6" class="data row3 col6" >1.0%</td>
      <td id="T_c148e_row3_col7" class="data row3 col7" >500.00</td>
      <td id="T_c148e_row3_col8" class="data row3 col8" >495.00</td>
      <td id="T_c148e_row3_col9" class="data row3 col9" >4,950,000.00</td>
    </tr>
    <tr>
      <th id="T_c148e_level0_row4" class="row_heading level0 row4" rowspan="2">3</th>
      <th id="T_c148e_level1_row4" class="row_heading level1 row4" >Paper</th>
      <td id="T_c148e_row4_col0" class="data row4 col0" >9</td>
      <td id="T_c148e_row4_col1" class="data row4 col1" >12,000.00</td>
      <td id="T_c148e_row4_col2" class="data row4 col2" >4</td>
      <td id="T_c148e_row4_col3" class="data row4 col3" >T1/1</td>
      <td id="T_c148e_row4_col4" class="data row4 col4" >9,800.00</td>
      <td id="T_c148e_row4_col5" class="data row4 col5" >14,000.00</td>
      <td id="T_c148e_row4_col6" class="data row4 col6" >1.0%</td>
      <td id="T_c148e_row4_col7" class="data row4 col7" >300.00</td>
      <td id="T_c148e_row4_col8" class="data row4 col8" >297.00</td>
      <td id="T_c148e_row4_col9" class="data row4 col9" >3,564,000.00</td>
    </tr>
    <tr>
      <th id="T_c148e_level1_row5" class="row_heading level1 row5" >String</th>
      <td id="T_c148e_row5_col0" class="data row5 col0" >7</td>
      <td id="T_c148e_row5_col1" class="data row5 col1" >25,000.00</td>
      <td id="T_c148e_row5_col2" class="data row5 col2" >5</td>
      <td id="T_c148e_row5_col3" class="data row5 col3" >T1/1</td>
      <td id="T_c148e_row5_col4" class="data row5 col4" >14,000.00</td>
      <td id="T_c148e_row5_col5" class="data row5 col5" >28,000.00</td>
      <td id="T_c148e_row5_col6" class="data row5 col6" >2.0%</td>
      <td id="T_c148e_row5_col7" class="data row5 col7" >20.00</td>
      <td id="T_c148e_row5_col8" class="data row5 col8" >19.60</td>
      <td id="T_c148e_row5_col9" class="data row5 col9" >490,000.00</td>
    </tr>
    <tr>
      <th id="T_c148e_level0_row6" class="row_heading level0 row6" >4</th>
      <th id="T_c148e_level1_row6" class="row_heading level1 row6" >Paper</th>
      <td id="T_c148e_row6_col0" class="data row6 col0" >10</td>
      <td id="T_c148e_row6_col1" class="data row6 col1" >12,000.00</td>
      <td id="T_c148e_row6_col2" class="data row6 col2" >4</td>
      <td id="T_c148e_row6_col3" class="data row6 col3" >T1/1</td>
      <td id="T_c148e_row6_col4" class="data row6 col4" >9,800.00</td>
      <td id="T_c148e_row6_col5" class="data row6 col5" >14,000.00</td>
      <td id="T_c148e_row6_col6" class="data row6 col6" >1.0%</td>
      <td id="T_c148e_row6_col7" class="data row6 col7" >300.00</td>
      <td id="T_c148e_row6_col8" class="data row6 col8" >297.00</td>
      <td id="T_c148e_row6_col9" class="data row6 col9" >3,564,000.00</td>
    </tr>
    <tr>
      <th id="T_c148e_level0_row7" class="row_heading level0 row7" >5</th>
      <th id="T_c148e_level1_row7" class="row_heading level1 row7" >Paper</th>
      <td id="T_c148e_row7_col0" class="data row7 col0" >11</td>
      <td id="T_c148e_row7_col1" class="data row7 col1" >12,000.00</td>
      <td id="T_c148e_row7_col2" class="data row7 col2" >4</td>
      <td id="T_c148e_row7_col3" class="data row7 col3" >T1/1</td>
      <td id="T_c148e_row7_col4" class="data row7 col4" >9,800.00</td>
      <td id="T_c148e_row7_col5" class="data row7 col5" >14,000.00</td>
      <td id="T_c148e_row7_col6" class="data row7 col6" >1.0%</td>
      <td id="T_c148e_row7_col7" class="data row7 col7" >300.00</td>
      <td id="T_c148e_row7_col8" class="data row7 col8" >297.00</td>
      <td id="T_c148e_row7_col9" class="data row7 col9" >3,564,000.00</td>
    </tr>
    <tr>
      <th id="T_c148e_level0_row8" class="row_heading level0 row8" >6</th>
      <th id="T_c148e_level1_row8" class="row_heading level1 row8" >Paper</th>
      <td id="T_c148e_row8_col0" class="data row8 col0" >12</td>
      <td id="T_c148e_row8_col1" class="data row8 col1" >12,000.00</td>
      <td id="T_c148e_row8_col2" class="data row8 col2" >4</td>
      <td id="T_c148e_row8_col3" class="data row8 col3" >T1/1</td>
      <td id="T_c148e_row8_col4" class="data row8 col4" >9,800.00</td>
      <td id="T_c148e_row8_col5" class="data row8 col5" >14,000.00</td>
      <td id="T_c148e_row8_col6" class="data row8 col6" >1.0%</td>
      <td id="T_c148e_row8_col7" class="data row8 col7" >300.00</td>
      <td id="T_c148e_row8_col8" class="data row8 col8" >297.00</td>
      <td id="T_c148e_row8_col9" class="data row8 col9" >3,564,000.00</td>
    </tr>
    <tr>
      <th id="T_c148e_level0_row9" class="row_heading level0 row9" >7</th>
      <th id="T_c148e_level1_row9" class="row_heading level1 row9" >Paper</th>
      <td id="T_c148e_row9_col0" class="data row9 col0" >13</td>
      <td id="T_c148e_row9_col1" class="data row9 col1" >12,000.00</td>
      <td id="T_c148e_row9_col2" class="data row9 col2" >4</td>
      <td id="T_c148e_row9_col3" class="data row9 col3" >T1/1</td>
      <td id="T_c148e_row9_col4" class="data row9 col4" >9,800.00</td>
      <td id="T_c148e_row9_col5" class="data row9 col5" >14,000.00</td>
      <td id="T_c148e_row9_col6" class="data row9 col6" >1.0%</td>
      <td id="T_c148e_row9_col7" class="data row9 col7" >300.00</td>
      <td id="T_c148e_row9_col8" class="data row9 col8" >297.00</td>
      <td id="T_c148e_row9_col9" class="data row9 col9" >3,564,000.00</td>
    </tr>
    <tr>
      <th id="T_c148e_level0_row10" class="row_heading level0 row10" rowspan="2">8</th>
      <th id="T_c148e_level1_row10" class="row_heading level1 row10" >Paper</th>
      <td id="T_c148e_row10_col0" class="data row10 col0" >14</td>
      <td id="T_c148e_row10_col1" class="data row10 col1" >12,000.00</td>
      <td id="T_c148e_row10_col2" class="data row10 col2" >4</td>
      <td id="T_c148e_row10_col3" class="data row10 col3" >T1/1</td>
      <td id="T_c148e_row10_col4" class="data row10 col4" >9,800.00</td>
      <td id="T_c148e_row10_col5" class="data row10 col5" >14,000.00</td>
      <td id="T_c148e_row10_col6" class="data row10 col6" >1.0%</td>
      <td id="T_c148e_row10_col7" class="data row10 col7" >300.00</td>
      <td id="T_c148e_row10_col8" class="data row10 col8" >297.00</td>
      <td id="T_c148e_row10_col9" class="data row10 col9" >3,564,000.00</td>
    </tr>
    <tr>
      <th id="T_c148e_level1_row11" class="row_heading level1 row11" >Poly Etylen</th>
      <td id="T_c148e_row11_col0" class="data row11 col0" >14</td>
      <td id="T_c148e_row11_col1" class="data row11 col1" >10,110.00</td>
      <td id="T_c148e_row11_col2" class="data row11 col2" >6</td>
      <td id="T_c148e_row11_col3" class="data row11 col3" >T1/1</td>
      <td id="T_c148e_row11_col4" class="data row11 col4" >9,800.00</td>
      <td id="T_c148e_row11_col5" class="data row11 col5" >14,000.00</td>
      <td id="T_c148e_row11_col6" class="data row11 col6" >1.0%</td>
      <td id="T_c148e_row11_col7" class="data row11 col7" >500.00</td>
      <td id="T_c148e_row11_col8" class="data row11 col8" >495.00</td>
      <td id="T_c148e_row11_col9" class="data row11 col9" >5,004,450.00</td>
    </tr>
    <tr>
      <th id="T_c148e_level0_row12" class="row_heading level0 row12" >12</th>
      <th id="T_c148e_level1_row12" class="row_heading level1 row12" >Paper</th>
      <td id="T_c148e_row12_col0" class="data row12 col0" >18</td>
      <td id="T_c148e_row12_col1" class="data row12 col1" >12,000.00</td>
      <td id="T_c148e_row12_col2" class="data row12 col2" >4</td>
      <td id="T_c148e_row12_col3" class="data row12 col3" >T1/1</td>
      <td id="T_c148e_row12_col4" class="data row12 col4" >9,800.00</td>
      <td id="T_c148e_row12_col5" class="data row12 col5" >14,000.00</td>
      <td id="T_c148e_row12_col6" class="data row12 col6" >1.0%</td>
      <td id="T_c148e_row12_col7" class="data row12 col7" >300.00</td>
      <td id="T_c148e_row12_col8" class="data row12 col8" >297.00</td>
      <td id="T_c148e_row12_col9" class="data row12 col9" >3,564,000.00</td>
    </tr>
    <tr>
      <th id="T_c148e_level0_row13" class="row_heading level0 row13" >14</th>
      <th id="T_c148e_level1_row13" class="row_heading level1 row13" >Paper</th>
      <td id="T_c148e_row13_col0" class="data row13 col0" >20</td>
      <td id="T_c148e_row13_col1" class="data row13 col1" >2,810.81</td>
      <td id="T_c148e_row13_col2" class="data row13 col2" >1</td>
      <td id="T_c148e_row13_col3" class="data row13 col3" >T0/1</td>
      <td id="T_c148e_row13_col4" class="data row13 col4" >0.00</td>
      <td id="T_c148e_row13_col5" class="data row13 col5" >9,800.00</td>
      <td id="T_c148e_row13_col6" class="data row13 col6" >0.0%</td>
      <td id="T_c148e_row13_col7" class="data row13 col7" >300.00</td>
      <td id="T_c148e_row13_col8" class="data row13 col8" >300.00</td>
      <td id="T_c148e_row13_col9" class="data row13 col9" >843,243.00</td>
    </tr>
    <tr>
      <th id="T_c148e_level0_row14" class="row_heading level0 row14" >19</th>
      <th id="T_c148e_level1_row14" class="row_heading level1 row14" >Paper</th>
      <td id="T_c148e_row14_col0" class="data row14 col0" >25</td>
      <td id="T_c148e_row14_col1" class="data row14 col1" >9,000.00</td>
      <td id="T_c148e_row14_col2" class="data row14 col2" >3</td>
      <td id="T_c148e_row14_col3" class="data row14 col3" >T0/1</td>
      <td id="T_c148e_row14_col4" class="data row14 col4" >0.00</td>
      <td id="T_c148e_row14_col5" class="data row14 col5" >9,800.00</td>
      <td id="T_c148e_row14_col6" class="data row14 col6" >0.0%</td>
      <td id="T_c148e_row14_col7" class="data row14 col7" >300.00</td>
      <td id="T_c148e_row14_col8" class="data row14 col8" >300.00</td>
      <td id="T_c148e_row14_col9" class="data row14 col9" >2,700,000.00</td>
    </tr>
    <tr>
      <th id="T_c148e_level0_row15" class="row_heading level0 row15" >22</th>
      <th id="T_c148e_level1_row15" class="row_heading level1 row15" >Paper</th>
      <td id="T_c148e_row15_col0" class="data row15 col0" >28</td>
      <td id="T_c148e_row15_col1" class="data row15 col1" >12,000.00</td>
      <td id="T_c148e_row15_col2" class="data row15 col2" >4</td>
      <td id="T_c148e_row15_col3" class="data row15 col3" >T1/1</td>
      <td id="T_c148e_row15_col4" class="data row15 col4" >9,800.00</td>
      <td id="T_c148e_row15_col5" class="data row15 col5" >14,000.00</td>
      <td id="T_c148e_row15_col6" class="data row15 col6" >1.0%</td>
      <td id="T_c148e_row15_col7" class="data row15 col7" >300.00</td>
      <td id="T_c148e_row15_col8" class="data row15 col8" >297.00</td>
      <td id="T_c148e_row15_col9" class="data row15 col9" >3,564,000.00</td>
    </tr>
    <tr>
      <th id="T_c148e_level0_row16" class="row_heading level0 row16" >24</th>
      <th id="T_c148e_level1_row16" class="row_heading level1 row16" >String</th>
      <td id="T_c148e_row16_col0" class="data row16 col0" >28</td>
      <td id="T_c148e_row16_col1" class="data row16 col1" >3,719.96</td>
      <td id="T_c148e_row16_col2" class="data row16 col2" >1</td>
      <td id="T_c148e_row16_col3" class="data row16 col3" >T0/1</td>
      <td id="T_c148e_row16_col4" class="data row16 col4" >0.00</td>
      <td id="T_c148e_row16_col5" class="data row16 col5" >14,000.00</td>
      <td id="T_c148e_row16_col6" class="data row16 col6" >0.0%</td>
      <td id="T_c148e_row16_col7" class="data row16 col7" >20.00</td>
      <td id="T_c148e_row16_col8" class="data row16 col8" >20.00</td>
      <td id="T_c148e_row16_col9" class="data row16 col9" >74,399.24</td>
    </tr>
  </tbody>
</table>


</details>

<!-- END TABLE:material-orders -->

<!-- BEGIN TABLE:warehouse -->

<details>
<summary>Warehouse occupation — ending inventory</summary>

<div>
<style scoped>
    .dataframe tbody tr th:only-of-type {
        vertical-align: middle;
    }

    .dataframe tbody tr th {
        vertical-align: top;
    }

    .dataframe thead th {
        text-align: right;
    }
</style>
<table border="1" class="dataframe">
  <thead>
    <tr style="text-align: right;">
      <th></th>
      <th>Used Space</th>
      <th>Remaining Space</th>
      <th>Total Capacity</th>
      <th>% Free</th>
    </tr>
    <tr>
      <th>Day</th>
      <th></th>
      <th></th>
      <th></th>
      <th></th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <th>Day 1</th>
      <td>995.98</td>
      <td>1004.02</td>
      <td>2000.00</td>
      <td>50.2%</td>
    </tr>
    <tr>
      <th>Day 2</th>
      <td>859.03</td>
      <td>1140.97</td>
      <td>2000.00</td>
      <td>57.0%</td>
    </tr>
    <tr>
      <th>Day 3</th>
      <td>727.47</td>
      <td>1272.53</td>
      <td>2000.00</td>
      <td>63.6%</td>
    </tr>
    <tr>
      <th>Day 4</th>
      <td>640.88</td>
      <td>1359.12</td>
      <td>2000.00</td>
      <td>68.0%</td>
    </tr>
    <tr>
      <th>Day 5</th>
      <td>640.88</td>
      <td>1359.12</td>
      <td>2000.00</td>
      <td>68.0%</td>
    </tr>
    <tr>
      <th>Day 6</th>
      <td>640.88</td>
      <td>1359.12</td>
      <td>2000.00</td>
      <td>68.0%</td>
    </tr>
    <tr>
      <th>Day 7</th>
      <td>920.90</td>
      <td>1079.10</td>
      <td>2000.00</td>
      <td>54.0%</td>
    </tr>
    <tr>
      <th>Day 8</th>
      <td>1068.85</td>
      <td>931.15</td>
      <td>2000.00</td>
      <td>46.6%</td>
    </tr>
    <tr>
      <th>Day 9</th>
      <td>1009.60</td>
      <td>990.40</td>
      <td>2000.00</td>
      <td>49.5%</td>
    </tr>
    <tr>
      <th>Day 10</th>
      <td>1042.01</td>
      <td>957.99</td>
      <td>2000.00</td>
      <td>47.9%</td>
    </tr>
    <tr>
      <th>Day 11</th>
      <td>1075.61</td>
      <td>924.39</td>
      <td>2000.00</td>
      <td>46.2%</td>
    </tr>
    <tr>
      <th>Day 12</th>
      <td>1104.93</td>
      <td>895.07</td>
      <td>2000.00</td>
      <td>44.8%</td>
    </tr>
    <tr>
      <th>Day 13</th>
      <td>1129.60</td>
      <td>870.40</td>
      <td>2000.00</td>
      <td>43.5%</td>
    </tr>
    <tr>
      <th>Day 14</th>
      <td>1264.00</td>
      <td>736.00</td>
      <td>2000.00</td>
      <td>36.8%</td>
    </tr>
    <tr>
      <th>Day 15</th>
      <td>1184.46</td>
      <td>815.54</td>
      <td>2000.00</td>
      <td>40.8%</td>
    </tr>
    <tr>
      <th>Day 16</th>
      <td>1110.80</td>
      <td>889.20</td>
      <td>2000.00</td>
      <td>44.5%</td>
    </tr>
    <tr>
      <th>Day 17</th>
      <td>1024.50</td>
      <td>975.50</td>
      <td>2000.00</td>
      <td>48.8%</td>
    </tr>
    <tr>
      <th>Day 18</th>
      <td>985.61</td>
      <td>1014.39</td>
      <td>2000.00</td>
      <td>50.7%</td>
    </tr>
    <tr>
      <th>Day 19</th>
      <td>985.61</td>
      <td>1014.39</td>
      <td>2000.00</td>
      <td>50.7%</td>
    </tr>
    <tr>
      <th>Day 20</th>
      <td>1013.72</td>
      <td>986.28</td>
      <td>2000.00</td>
      <td>49.3%</td>
    </tr>
    <tr>
      <th>Day 21</th>
      <td>927.50</td>
      <td>1072.50</td>
      <td>2000.00</td>
      <td>53.6%</td>
    </tr>
    <tr>
      <th>Day 22</th>
      <td>761.68</td>
      <td>1238.32</td>
      <td>2000.00</td>
      <td>61.9%</td>
    </tr>
    <tr>
      <th>Day 23</th>
      <td>685.86</td>
      <td>1314.14</td>
      <td>2000.00</td>
      <td>65.7%</td>
    </tr>
    <tr>
      <th>Day 24</th>
      <td>599.53</td>
      <td>1400.47</td>
      <td>2000.00</td>
      <td>70.0%</td>
    </tr>
    <tr>
      <th>Day 25</th>
      <td>574.37</td>
      <td>1425.63</td>
      <td>2000.00</td>
      <td>71.3%</td>
    </tr>
    <tr>
      <th>Day 26</th>
      <td>574.37</td>
      <td>1425.63</td>
      <td>2000.00</td>
      <td>71.3%</td>
    </tr>
    <tr>
      <th>Day 27</th>
      <td>574.37</td>
      <td>1425.63</td>
      <td>2000.00</td>
      <td>71.3%</td>
    </tr>
    <tr>
      <th>Day 28</th>
      <td>656.80</td>
      <td>1343.20</td>
      <td>2000.00</td>
      <td>67.2%</td>
    </tr>
    <tr>
      <th>Day 29</th>
      <td>584.74</td>
      <td>1415.26</td>
      <td>2000.00</td>
      <td>70.8%</td>
    </tr>
    <tr>
      <th>Day 30</th>
      <td>510.00</td>
      <td>1490.00</td>
      <td>2000.00</td>
      <td>74.5%</td>
    </tr>
  </tbody>
</table>
</div>

</details>

<!-- END TABLE:warehouse -->

The warehouse display reports ending occupation. The material model checks warehouse capacity after arrivals, before daily consumption.

## Scope and assumptions

This is a planning prototype with deterministic forecast inputs and processing standards. Forecasting itself is outside this repository. `Profit_Per_Product` is a supplied planning parameter rather than a complete accounting calculation.

Production flow is aggregated by day. Same-day transfer permits downstream processing against that day's upstream output; it does not establish exact within-day start times or a minute-by-minute executable sequence. Sequence-dependent setup times and machine breakdown uncertainty are not modeled.

Ending WIP is allowed by default in both production stages. The OT model keeps its WIP separate, and the material model consumes BOM quantities on the dates in its supplied schedule. Stage-specific material consumption and a joint optimization of production, OT and procurement would require extending the current models.

