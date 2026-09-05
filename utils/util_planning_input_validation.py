from math import ceil


def validate_input_safeguards(
    days,
    materials,
    products,
    requirement,
    warehouse_space,
    warehouse_capacity,
    production_due,
    bom,
    lead_time,
    safety_stock,
    initial_inventory,
    max_material_warehouse_share=0.80,
):
    """
    Validates material planning inputs against warehouse share limits (Safeguard 1)
    and lead-time inventory coverage (Safeguard 2). 

    Returns:
        str | None: Error message string if validation fails, None if valid.
    """
    start_day = min(days)
    validation_error = None

    def get_product_explanation(m, material_violation_units, day):
        # 1. Look for products scheduled directly on this day that use material m
        products_on_day = [
            (p, production_due.get((p, day), 0))
            for p in products
            if production_due.get((p, day), 0) > 0 and bom.get((p, m), 0) > 0
        ]

        # 2. Fallback: Check all products scheduled up to 'day'
        if not products_on_day:
            products_on_day = [
                (
                    p,
                    sum(
                        production_due.get((p, t), 0)
                        for t in range(start_day, day + 1)
                    ),
                )
                for p in products
                if sum(
                    production_due.get((p, t), 0)
                    for t in range(start_day, day + 1)
                )
                > 0
                and bom.get((p, m), 0) > 0
            ]

        # Single product match
        if len(products_on_day) == 1:
            product, planned_product_qty = products_on_day[0]
            material_per_product = bom[product, m]

            required_product_reduction = ceil(
                material_violation_units / material_per_product
            )
            maximum_product_qty = max(
                0.0, planned_product_qty - required_product_reduction
            )

            return (
                f"Scheduled product: {product}\n"
                f"Scheduled product quantity: {planned_product_qty:.2f}\n"
                f"BOM usage: {material_per_product:.2f} {m} per {product}\n"
                f"Maximum allowable {product} quantity: {maximum_product_qty:.2f}\n"
                f"Minimum product reduction required: {required_product_reduction} units"
            )

        # Multiple product matches
        elif len(products_on_day) > 1:
            contribution_lines = []
            reduction_lines = []

            for product, planned_product_qty in products_on_day:
                material_per_product = bom[product, m]
                material_contribution = planned_product_qty * material_per_product
                equivalent_reduction = ceil(
                    material_violation_units / material_per_product
                )

                contribution_lines.append(
                    f"{product}: {planned_product_qty:.2f} products × "
                    f"{material_per_product:.2f} {m}/product = {material_contribution:.2f} {m}"
                )
                reduction_lines.append(
                    f"Reduce {product} by {equivalent_reduction} units if only {product} is adjusted"
                )

            return (
                "Multiple products are scheduled, so there is no single exact product conversion.\n\n"
                "Material contributions:\n" + "\n".join(contribution_lines) +
                "\n\nPossible reductions:\n" + "\n".join(reduction_lines)
            )

        return "No scheduled products identified for this material requirement."

    # Safeguard 1: Single material space limit
    for d in days:
        for m in materials:
            required_material_units = requirement[m, d]
            required_space = required_material_units * warehouse_space[m]
            allowed_space = max_material_warehouse_share * warehouse_capacity

            if required_space > allowed_space:
                allowed_material_units = allowed_space / warehouse_space[m]
                excess_material_units = (
                    required_material_units - allowed_material_units
                )

                product_explanation = get_product_explanation(
                    m, excess_material_units, d
                )

                validation_error = (
                    f"Safeguard 1 failed (Warehouse Share Exceeded):\n"
                    f"Material: {m}\n"
                    f"Production day: {d}\n"
                    f"Material requirement: {required_material_units:.2f} units\n"
                    f"Warehouse space required: {required_space:.2f}\n"
                    f"{max_material_warehouse_share * 100:.0f}% warehouse allowance: {allowed_space:.2f}\n"
                    f"Maximum allowable {m}: {allowed_material_units:.2f} units\n"
                    f"Excess {m}: {excess_material_units:.2f} units\n\n"
                    f"{product_explanation}"
                )
                break

        if validation_error is not None:
            break

    # Safeguard 2: Initial inventory coverage before first supplier arrival
    if validation_error is None:
        for m in materials:
            first_possible_arrival = start_day + lead_time[m]
            cumulative_early_requirement = 0

            for d in days:
                if d >= first_possible_arrival:
                    break

                cumulative_early_requirement += requirement[m, d]
                minimum_initial_inventory = (
                    cumulative_early_requirement + safety_stock[m]
                )

                if initial_inventory[m] < minimum_initial_inventory:
                    shortage = minimum_initial_inventory - initial_inventory[m]

                    product_explanation = get_product_explanation(m, shortage, d)

                    validation_error = (
                        f"Safeguard 2 failed (Lead-Time Inventory Shortage):\n"
                        f"Material: {m}\n"
                        f"Problem detected by day: {d}\n"
                        f"First possible supplier arrival: day {first_possible_arrival}\n"
                        f"Initial inventory: {initial_inventory[m]:.2f}\n"
                        f"Cumulative early requirement: {cumulative_early_requirement:.2f}\n"
                        f"Safety stock: {safety_stock[m]:.2f}\n"
                        f"Minimum initial inventory needed: {minimum_initial_inventory:.2f}\n"
                        f"Shortage: {shortage:.2f} units of {m}\n\n"
                        f"{product_explanation}"
                    )
                    break

            if validation_error is not None:
                break

    return validation_error