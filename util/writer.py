import data
import pandas as pd

class SalesReport:
    def __init__(self) -> None:
        pass

    def add_borders(self, writer, start_row, end_row, thick_borders, thin_borders):
        writer.sheets['Sales'].conditional_format(start_row, 0, start_row, 22, {'type': 'no_blanks', 'format': thick_borders})
        writer.sheets['Sales'].conditional_format(start_row, 0, start_row, 22, {'type': 'blanks', 'format': thick_borders})
        writer.sheets['Sales'].conditional_format(start_row+1, 0, end_row, 22, {'type': 'no_blanks', 'format': thin_borders})
        writer.sheets['Sales'].conditional_format(start_row+1, 0, end_row, 22, {'type': 'blanks', 'format': thin_borders})

    def generate(self):
        
        writer = pd.ExcelWriter('report.xlsx', engine = "xlsxwriter", 
                                                engine_kwargs={'options': {'strings_to_numbers': False}})
        workbook  = writer.book

        #report formatting
        #borders
        thick_borders = workbook.add_format({'border': 2, 'border_color': '#000000'})
        thin_borders = workbook.add_format({'border': 1, 'border_color': '#A9A9A9'})

        #text formats
        title_text = workbook.add_format({'bold': True, 'font_size': 20, 'align':'center'})
        subtitle_text = workbook.add_format({'bold': True, 'font_size': 16, 'align':'center'})
        column_header_text = workbook.add_format({'bold': True, 'font_size': 14, 'align':'center'})
        bold_text = workbook.add_format({'bold': True, 'font_size': 14,})
        regular_text= workbook.add_format({'font_size': 14,})
        # date format
        date_format = workbook.add_format({'num_format': 'dd-mmm-yyyy', 'font_size': 14, 'align':'right'})
        #number format
        number = workbook.add_format({'num_format': '#,##0.00_);(#,##0.00)', 'font_size': 11})
        bold_number = workbook.add_format({'bold': True, 'num_format': '#,##0.00_);(#,##0.00)', 'font_size': 14, 'align':'right'})

        #report body
        report_data = data.sales_data()

        #set report start row and column
        start_row = 3
        start_col = 0
        #change index to start from 1
        report_data.index = [i for i in range(1, len(report_data) + 1)]
        report_data.to_excel(writer, sheet_name='Sales', index=True, 
                                    startrow=start_row, startcol=start_col, header=True)
        
        #hide gridlines
        writer.sheets['Sales'].hide_gridlines(2)

        #report header
        writer.sheets['Sales'].merge_range("A1:F1", "FURNITURE STORE SALES REPORT", title_text)
        writer.sheets['Sales'].merge_range("A2:F2", "1 JAN 2023 - 3 JAN 2023", subtitle_text)
        
        #Profit column
        writer.sheets['Sales'].write(f'F{start_row+1}', 'Profit', column_header_text)
        #calculate profit
        for i in range(start_row+2, start_row+2 + len(report_data)):
            writer.sheets['Sales'].write(f'F{i}', f'=E{i}-D{i}', number)

        #TOTAL
        last_row = len(report_data) + start_row + 2
        writer.sheets['Sales'].write(f'C{last_row}', 'Total', bold_text)
        writer.sheets['Sales'].write_formula(f'E{last_row}', f'=SUM(E{start_row+2}:E{last_row-1})', bold_number)
        writer.sheets['Sales'].write_formula(f'F{last_row}', f'=SUM(F{start_row+2}:F{last_row-1})', bold_number)

        #ADD BORDERS
        writer.sheets['Sales'].conditional_format(start_row, 0, start_row, 5, {'type': 'no_blanks', 'format': thick_borders})
        writer.sheets['Sales'].conditional_format(start_row, 0, start_row, 5, {'type': 'blanks', 'format': thick_borders})
        writer.sheets['Sales'].conditional_format(start_row+1, 0, last_row-1, 5, {'type': 'no_blanks', 'format': thin_borders})
        writer.sheets['Sales'].conditional_format(start_row+1, 0, last_row-1, 5, {'type': 'blanks', 'format': thin_borders})

        #increase table font size
        writer.sheets['Sales'].conditional_format(f'D{start_row+2}:E{last_row}', {'type': 'no_blanks', 'format': number})


        workbook.close()


SalesReport().generate()
