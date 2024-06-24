# ----- Imports ----- #
import helper.excelFunctions as excel


def highlight(worksheet, odd_attachments):
    print("Highlighting...")

    # Helper variables
    row = 17
    name_col = 12
    exist_col = 13
    propose_col = 14
    cell = excel.get_cell_value(worksheet, row, 12)

    # Loops through attachment name cells
    while cell is not None:

        # Multiple proposed in single cell
        propose_cell = excel.get_cell_value(worksheet, row, propose_col)
        if propose_cell is not None:
            if len(propose_cell.split("  ")) > 2:
                excel.set_fill(worksheet, row, column=propose_col, r=255, g=0, b=0)

        # Drip Loop
        if 'drip loop' in excel.get_cell_value(worksheet, row, name_col).lower():
            excel.set_fill(worksheet, row, column=exist_col, r=255, g=0, b=0)

        # Street Light Drip
        if 'street light drip' in excel.get_cell_value(worksheet, row, name_col).lower():
            excel.set_fill(worksheet, row, column=exist_col, r=255, g=0, b=0)

        # Crossarm Brace
        if 'crossarm brace' in excel.get_cell_value(worksheet, row, name_col).lower():
            excel.set_fill(worksheet, row, column=exist_col, r=255, g=0, b=0)

        # # NWE Top Transformer
        # if 'nwe top transformer' in excel.get_cell_value(worksheet, row, name_col).lower():
        #     excel.set_fill(worksheet, row, column=exist_col, r=255, g=0, b=0)
        #
        # # NWE Base Transformer
        # if 'nwe base transformer' in excel.get_cell_value(worksheet, row, name_col).lower():
        #     excel.set_fill(worksheet, row, column=exist_col, r=255, g=0, b=0)

        # # NWE Primary Crossarm
        # if 'nwe primary crossarm' in excel.get_cell_value(worksheet, row, name_col).lower() and excel.get_cell_value(worksheet, row, propose_col) is not None:
        #     excel.set_fill(worksheet, row, column=propose_col, r=255, g=0, b=0)
        #
        # # NWE Secondary Crossarm
        # if 'nwe secondary crossarm' in excel.get_cell_value(worksheet, row, name_col).lower() and excel.get_cell_value(worksheet, row, propose_col) is not None:
        #     excel.set_fill(worksheet, row, column=propose_col, r=255, g=0, b=0)

        # NWE Neutral
        if 'nwe neutral' in excel.get_cell_value(worksheet, row, name_col).lower() and excel.get_cell_value(worksheet, row, propose_col) is not None:
            excel.set_fill(worksheet, row, column=propose_col, r=255, g=0, b=0)

        # NWE Secondary Drop
        if 'nwe secondary drop' in excel.get_cell_value(worksheet, row, name_col).lower() and excel.get_cell_value(worksheet, row, propose_col) is not None:
            excel.set_fill(worksheet, row, column=propose_col, r=255, g=0, b=0)

        # Street Light
        if 'street light' in excel.get_cell_value(worksheet, row, name_col).lower() and excel.get_cell_value(worksheet, row, propose_col) is not None:
            excel.set_fill(worksheet, row, column=propose_col, r=255, g=0, b=0)

        # # Street Light
        # if 'deadend' in excel.get_cell_value(worksheet, row, name_col).lower():
        #     excel.set_fill(worksheet, row, column=name_col, r=255, g=0, b=0)

        # Odd Attachments
        for value in odd_attachments:
            if value in cell:
                excel.set_fill(worksheet, row, column=name_col, r=255, g=0, b=0)

        # Next cell
        row += 1
        cell = excel.get_cell_value(worksheet, row, name_col)

    print("Highlighting Complete")
