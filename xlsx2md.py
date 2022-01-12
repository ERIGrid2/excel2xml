#! /usr/bin/python3

from datetime import date
import os
import sys
import chevron
from openpyxl import load_workbook
from openpyxl.cell.cell import Cell
from openpyxl.worksheet.worksheet import Worksheet

def extract_test_case(sheet_ranges : Worksheet):
    test_case = {}
    test_case['ID'] = sheet_ranges['C2'].value
    test_case['Name'] = sheet_ranges['C3'].value
    extract_generic_data(sheet_ranges, test_case, start_row=4)
    return test_case

def extract_test_specification(sheet_ranges):
    pass

def extract_experiment_specification(sheet_ranges):
    pass

def extract_generic_data(sheet_ranges : Worksheet, object, start_row=1, end_row=None):
    if not 'subsections' in object:
        object['subsections'] = []

    for row_idx in range(start_row, 1000):
        if sheet_ranges['A' + str(row_idx)].value == 'Diagrams':
            break
        
        headline_cell = sheet_ranges['B' + str(row_idx)]
        if headline_cell.font.bold:
            section = {
                'section_title': headline_cell.value,
                'subsections': []
            }
            object['subsections'].append(section)
            if is_gray(get_cell_right(headline_cell)):
                continue
            else:
                section['contents'] = get_cell_right(headline_cell).value
        elif headline_cell.value == 'Description':
            section['contents'] = get_cell_right(headline_cell).value
        elif headline_cell.value == 'Diagram reference':
            section['diagrams'] = extract_diagrams(sheet_ranges, get_cell_right(headline_cell).value)
        elif headline_cell.value:
            subsection = {
                'section_title': headline_cell.value,
                'contents': get_cell_right(headline_cell).value
            }
            section['subsections'].append(subsection)

    return object

def extract_diagrams(sheet : Worksheet, diagram_reference):
    return_list = []

    if not diagram_reference:
        return return_list

    diagram_references = [x.strip() for x in diagram_reference.split(';')]

    for i in range(1, 1000):
        if sheet['A' + str(i)].value == 'Diagrams':
            diagram_id_row = i + 1

    for dia_ref in diagram_references:
        col = 3
        while sheet.cell(row=diagram_id_row, column=col).value is not None:
            if sheet.cell(row=diagram_id_row, column=col).value == dia_ref:
                return_list.append(
                    {
                        'diagram_name': sheet.cell(row=diagram_id_row+1, column=col).value,
                        'diagram_uri': sheet.cell(row=diagram_id_row+3, column=col).value
                    })
                break
            col += 1

    return return_list

def get_cell_below(cell : Cell) -> Cell:
    ws = cell.parent
    return ws[cell.column_letter() + str(cell.row + 1)]

def get_cell_right(cell : Cell) -> Cell:
    ws = cell.parent
    return ws.cell(cell.row, cell.column + 1)

def is_gray(cell : Cell):
    return type(cell.fill.fgColor.theme) == int

def extract_table(start_cell : Cell, object):
    table_columns = []
    title_cell = start_cell
    while not title_cell.font.bold:
        column = []
        cell = title_cell
        while cell.value:
            column.append(cell.value)
            cell = get_cell_right(cell)            
        table_columns.append(column)

        title_cell = get_cell_below(title_cell)

    if len(table_columns) > 0:
        object['table'] = table_columns
        

def main(filename):
    try:
        wb = load_workbook(filename)
    except:
        print("File does not exist!")
        return

    test_case = None
    test_specifications = []
    experiment_specifications = []
    for sheet in wb:
        sheet_ranges = wb[sheet.title]
        sheet_type = sheet_ranges['A1']
        if sheet_type.value == 'Test Case':
            test_case = extract_test_case(sheet_ranges)
        elif sheet_type == 'Test Specification':
            test_specifications.append(extract_test_specification(sheet_ranges))
        elif sheet_type == 'Experiment Specification':
            experiment_specifications.append(extract_experiment_specification(sheet_ranges))

    if test_case:
        mtime = date.fromtimestamp(os.path.getmtime(filename)).isoformat()
        test_case_filename = os.path.splitext(os.path.basename(filename))[0]
        test_case['title'] = '"' + test_case_filename + '"'
        test_case['linkTitle'] = '"' + test_case_filename + '"'
        test_case['date'] = '"' + mtime + '"'
        test_case['description'] = '"' + test_case['Name'] + '"'
        with open('TestCase.mustache', 'r') as template:
            md_test_case = chevron.render(template=template, data=test_case)
            print(md_test_case)

#Python3
if __name__ == '__main__':
    if len(sys.argv) > 1:
        filename = str(sys.argv[1])
    else:
        print("No arguments introduced")

    main(filename)
