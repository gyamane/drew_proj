from django.shortcuts import render
import openpyxl
from .functions import sum_cells


def get_sum(request):
    if "GET" == request.method:
        return render(request, 'index.html', {})
    else:
        excel_file = request.FILES["excel_file"]
        workbook = openpyxl.load_workbook(excel_file)

        # Get all Excel sheets
        sheets = workbook.sheetnames

        total_sum = 0.0

        # Iterate over each Excel sheet and sum all numerical cells
        for sheet in sheets:
            total_sum = sum_cells(workbook[sheet], total_sum)

    return render(request, 'index.html', {"total_sum": total_sum})
