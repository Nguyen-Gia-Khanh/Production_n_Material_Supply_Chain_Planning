from pathlib import Path

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

def export_notebook_tables(notebook_path, output_dir="git_tables"):
    """Export saved notebook tables without changing their layout or formatting."""
    
    import json
    from pathlib import Path
    from html import escape

    # Read from notebook_path
    nb_path = Path(notebook_path)
    notebook = json.loads(nb_path.read_text(encoding="utf-8"))

    # Ensure output directory exists
    out_dir = Path(output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    blocks = []

    for cell in notebook["cells"]:
        title = "Notebook table"

        for output in cell.get("outputs", []):
            if output.get("output_type") == "stream":
                title = "".join(output.get("text", [])).strip() or title

            table_html = "".join(
                output.get("data", {}).get("text/html", [])
            )

            if "<table" in table_html:
                blocks.append(
                    f"<details>\n"
                    f"<summary>{escape(title)}</summary>\n\n"
                    f"{table_html}\n\n"
                    f"</details>"
                )

    if not blocks:
        raise ValueError(
            "No saved tables found. Save the notebook with its outputs first."
        )

    report = "\n\n".join(blocks)

    # Route output paths directly inside output_dir
    markdown_path = out_dir / f"{nb_path.stem}.tables.md"
    html_path = out_dir / f"{nb_path.stem}.tables.html"

    markdown_path.write_text(report, encoding="utf-8")
    html_path.write_text(
        "<!doctype html>\n"
        "<html><head><meta charset='utf-8'></head><body>\n"
        f"{report}\n"
        "</body></html>",
        encoding="utf-8",
    )

    print(f"Exported {len(blocks)} tables:")
    print(f"  Markdown: {markdown_path}")
    print(f"  HTML:     {html_path}")

    return markdown_path, html_path
  
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