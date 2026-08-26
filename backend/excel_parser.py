import openpyxl

def find_headers(sheet):
    """
    Scans the first 20 rows of the sheet to dynamically identify the header row.
    Matches keywords like type, product, crop, activity, event cost, budget.
    """
    for r in range(1, 21):
        row_vals = [str(sheet.cell(row=r, column=col).value).strip() for col in range(1, 15)]
        indices = {}
        for c_idx, val in enumerate(row_vals, start=1):
            if val is None or val == "None":
                continue
            val_lower = val.lower()
            if "type" in val_lower and "activity" not in val_lower:
                indices["type"] = c_idx
            elif "product" in val_lower:
                indices["product"] = c_idx
            elif "crop" in val_lower:
                indices["crop"] = c_idx
            elif "type of activity" in val_lower or ("activity" in val_lower and "budget" not in val_lower):
                indices["activity"] = c_idx
            elif "event cost" in val_lower or "cost" in val_lower:
                indices["event_cost"] = c_idx
            elif "budget for" in val_lower or "marketing budget" in val_lower or "dg activities" in val_lower:
                indices["qty"] = c_idx
            elif "total budget" in val_lower or "allocation" in val_lower:
                indices["allocation"] = c_idx
        
        # If we have at least 4 matches, this is our header row
        if len(indices) >= 4:
            default_keys = ["type", "product", "crop", "activity", "event_cost", "qty", "allocation"]
            standard_cols = {
                "type": 1,
                "product": 2,
                "crop": 3,
                "activity": 4,
                "event_cost": 5,
                "qty": 6,
                "allocation": 7
            }
            for k in default_keys:
                if k not in indices:
                    indices[k] = standard_cols[k]
            return r, indices
            
    # Default fallback to row 11
    return 11, {
        "type": 1,
        "product": 2,
        "crop": 3,
        "activity": 4,
        "event_cost": 5,
        "qty": 6,
        "allocation": 7
    }

def validate_budget_sheet(filepath):
    """
    Opens the excel sheet, dynamically locates data headers, parses data rows,
    and returns a success report with rows data and total aggregations.
    """
    try:
        # Load workbook (data_only=True to read static values of formulas if any)
        wb = openpyxl.load_workbook(filepath, data_only=True)
        sheet = wb.active
        
        header_row, col_map = find_headers(sheet)
        errors = []
        rows_data = []
        
        calculated_total_qty = 0
        calculated_total_budget = 0
        
        row_idx = header_row + 1
        while True:
            # Stop if the row is completely empty
            row_values = [sheet.cell(row=row_idx, column=col).value for col in range(1, 10)]
            if all(v is None for v in row_values):
                break
                
            row_type = sheet.cell(row=row_idx, column=col_map["type"]).value
            product = sheet.cell(row=row_idx, column=col_map["product"]).value
            crop = sheet.cell(row=row_idx, column=col_map["crop"]).value
            activity = sheet.cell(row=row_idx, column=col_map["activity"]).value
            
            # If product and type are both None, we probably hit the end
            if product is None and row_type is None:
                break
                
            # Read Event Cost, Qty, and Allocation
            try:
                event_cost = float(sheet.cell(row=row_idx, column=col_map["event_cost"]).value or 0.0)
                qty = float(sheet.cell(row=row_idx, column=col_map["qty"]).value or 0.0)
                
                cell_alloc = sheet.cell(row=row_idx, column=col_map["allocation"]).value
                if cell_alloc is None:
                    allocation = event_cost * qty
                else:
                    allocation = float(cell_alloc)
            except (ValueError, TypeError):
                errors.append(f"Row {row_idx}: Non-numeric value in budget columns")
                row_idx += 1
                continue

            calculated_total_qty += qty
            calculated_total_budget += allocation
            
            rows_data.append({
                "row": row_idx,
                "type": row_type,
                "product": product,
                "crop": crop,
                "activity": activity,
                "event_cost": event_cost,
                "qty": qty,
                "allocation": allocation
            })
            row_idx += 1

        wb.close()
        return {
            "success": len(errors) == 0,
            "errors": errors,
            "totals": {
                "totalQty": calculated_total_qty,
                "totalBudget": calculated_total_budget
            },
            "rows": rows_data
        }
    except Exception as e:
        return {
            "success": False,
            "errors": [f"Failed to parse Excel file: {e}"],
            "totals": {"totalQty": 0, "totalBudget": 0},
            "rows": []
        }
