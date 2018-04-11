import xlwings as xw


def lastRow(idx, workbook, col=1):
    """ Find the last row in the worksheet that contains data.

    idx: Specifies the worksheet to select. Starts counting from zero.

    workbook: Specifies the workbook

    col: The column in which to look for the last cell containing data.
    """

    ws = workbook.sheets[idx]

    lwr_r_cell = ws.cells.last_cell      # lower right cell
    lwr_row = lwr_r_cell.row             # row of the lower right cell
    lwr_cell = ws.range((lwr_row, col))  # change to your specified column

    if lwr_cell.value is None:
        lwr_cell = lwr_cell.end('up')   # go up untill you hit a non-empty cell

    return lwr_cell.row


def create_new_sheet(wb, sheet_name, prec_sheet ):
    try:
        wb.sheets[sheet_name]
    except Exception:
        wb.sheets.add(name=sheet_name, after=prec_sheet)
    else:
        wb.sheets[sheet_name].delete()
        wb.sheets.add(name=sheet_name, after=prec_sheet)
