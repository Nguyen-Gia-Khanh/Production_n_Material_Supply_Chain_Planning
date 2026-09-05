def style_grouped_table(df, formats=None):

    outer_index = list(df.index.get_level_values(0))

    outer_index_color = "#FFD966"
    alternating_colors = ["#FFFFFF", "#DDEBF7"]

    table_styles = [
        # Column headers
        {
            "selector": "th.col_heading",
            "props": [
                ("background-color", "#D9E1F2"),
                ("color", "#000000"),
                ("font-weight", "bold"),
                ("padding", "6px"),
                ("border", "1px solid #A6A6A6"),
            ],
        },

        # Outer index: Material for material_first
        {
            "selector": "th.row_heading.level0",
            "props": [
                ("background-color", outer_index_color),
                ("color", "#000000"),
                ("font-weight", "bold"),
                ("padding", "6px"),
                ("border-right", "2px solid #7F6000"),
            ],
        },
    ]

    for row_number, group in enumerate(outer_index):

        row_color = alternating_colors[row_number % 2]

        # Alternate data cells and all index levels except level 0
        table_styles.append(
            {
                "selector": (
                    f"td.row{row_number}, "
                    f"th.row_heading.row{row_number}:not(.level0)"
                ),
                "props": [
                    ("background-color", row_color),
                    ("color", "#000000"),
                    ("padding", "6px"),
                    ("border-bottom", "1px solid #D9D9D9"),
                ],
            }
        )

        # Thick line when a new outer-index group begins
        if (
            row_number == 0
            or group != outer_index[row_number - 1]
        ):
            table_styles.append(
                {
                    "selector": f".row{row_number}",
                    "props": [
                        ("border-top", "3px solid #595959"),
                    ],
                }
            )

    return (
        df.style
        .format(formats or {}, na_rep="-")
        .set_table_styles(table_styles)
    )


production_quantity_format = "{:,.0f}"

ot_people_format = {
    "Required Worker-Hours": "{:,.2f}",
    "Required Workers (4h Equivalent)": "{:,.2f}",
    "Solver Workers (4h Equivalent)": "{:,.2f}",
    "Line Workers": "{:,.0f}",
}

remaining_time_format = "{:,.2f}"

order_format = {
    "Order Quantity": "{:,.2f}",
    "Containers": "{:,.0f}",
    "Tier Lower": "{:,.2f}",
    "Tier Upper": "{:,.2f}",
    "Discount": "{:.1%}",
    "Original Unit Price": "{:,.2f}",
    "Discounted Unit Price": "{:,.2f}",
    "Purchase Subtotal": "{:,.2f}",
}

inventory_format = {
    "Arrivals": "{:,.2f}",
    "Production Requirement": "{:,.2f}",
    "Starting Inventory": "{:,.2f}",
    "Ending Inventory": "{:,.2f}",
    "Space per Unit": "{:,.2f}",
    "Starting Occupation (m²)": "{:,.2f}",
    "Ending Occupation (m²)": "{:,.2f}",
}

wip_display_format = {
    "Opening Carryover Units": "{:,.0f}",
    "Added Today": "{:,.0f}",
    "Moved Forward Today": "{:,.0f}",
    "Closing Carryover Units": "{:,.0f}",
    "Next Process Line-Hours": "{:,.2f}",
    "Remaining Route Line-Hours": "{:,.2f}",
}


def export_readme_table(
    df,
    *,
    section_id,
    title,
    readme_path="README.md",
    formats=None,
    columns_per_table=7,
    note=None,
):
    """Replace one marked README section with collapsible HTML tables.

    Pass the DataFrame behind an IPython display, not display()'s return value.
    The README must contain exactly one matching BEGIN/END TABLE marker pair.
    Wide frames are split by data columns; every part retains the full index
    and all rows. MultiIndex rows and columns are supported. `formats` accepts
    one format string, or a column-to-format-string/callable dictionary.
    Only the marked section is written; the DataFrame is never modified.
    """
    from html import escape
    from pathlib import Path
    import re

    import pandas as pd

    if not isinstance(df, pd.DataFrame):
        raise TypeError("Pass a pandas DataFrame, not a Styler or display() result.")
    if not re.fullmatch(r"[a-z0-9_-]+", section_id):
        raise ValueError("section_id must contain only a-z, 0-9, underscores or hyphens.")
    if (
        not isinstance(columns_per_table, int)
        or isinstance(columns_per_table, bool)
        or columns_per_table < 1
    ):
        raise ValueError("columns_per_table must be a positive integer.")
    if formats is not None and not isinstance(formats, (str, dict)):
        raise TypeError("formats must be a format string or a column-format dictionary.")

    path = Path(readme_path)
    source = path.read_text(encoding="utf-8")
    start = f"<!-- BEGIN TABLE:{section_id} -->"
    end = f"<!-- END TABLE:{section_id} -->"
    if source.count(start) != 1 or source.count(end) != 1:
        raise ValueError(f"README must contain exactly one marker pair for {section_id!r}.")
    start_position = source.index(start) + len(start)
    end_position = source.index(end)
    if end_position < start_position:
        raise ValueError("The END TABLE marker must follow its BEGIN TABLE marker.")

    parts = []
    if note:
        parts.append(f"<p>{escape(str(note))}</p>")

    part_count = max(1, (len(df.columns) + columns_per_table - 1) // columns_per_table)
    for part_number in range(part_count):
        first_column = part_number * columns_per_table
        piece = df.iloc[:, first_column:first_column + columns_per_table]
        label = str(title)
        if part_count > 1:
            label += f" — part {part_number + 1} of {part_count}"

        if piece.empty:
            table = "<p>No rows were produced for this table.</p>"
        else:
            formatters = {}
            for column in piece.columns:
                spec = formats if isinstance(formats, str) else (formats or {}).get(column)
                if spec is not None:
                    formatters[column] = spec.format if isinstance(spec, str) else spec
            table = piece.to_html(
                index=True,
                index_names=True,
                sparsify=False,
                max_rows=None,
                max_cols=None,
                border=0,
                escape=True,
                na_rep="—",
                float_format=lambda value: f"{value:,.2f}",
                formatters=formatters,
            )

        parts.append(
            f"<details>\n<summary>{escape(label)}</summary>\n\n"
            f"{table}\n\n</details>"
        )

    replacement = "\n\n" + "\n\n".join(parts) + "\n\n"
    path.write_text(
        source[:start_position] + replacement + source[end_position:],
        encoding="utf-8",
    )
    return path
