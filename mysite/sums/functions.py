def sum_cells(worksheet, total_sum):
    """Helper function to sum all cells in an Excel sheet and update total_sum"""
    for row in worksheet.iter_rows():
        for cell in row:
            if isinstance(cell.value, int) or isinstance(cell.value, float):
                total_sum += float(cell.value)
    return total_sum
