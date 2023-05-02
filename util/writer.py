from util import data
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
        
        writer = pd.ExcelWriter('report.xlxs', engine = "xlsxwriter", 
                                                engine_kwargs={'options': {'strings_to_numbers': False}})
        workbook  = writer.book

        #report formatting
        #borders
        thick_borders = workbook.add_format({'border': 2, 'border_color': '#000000'})
        thin_borders = workbook.add_format({'border': 1, 'border_color': '#A9A9A9'})

        #text formats
        title_text = workbook.add_format({'bold': True, 'font_size': 20, 'align':'center'})
        subtitle_text = workbook.add_format({'bold': True, 'font_size': 16, 'align':'center'})
        column_header_text = workbook.add_format({'bold': True, 'font_size': 11, 'align':'center'})
        # date format
        date_format = workbook.add_format({'num_format': 'dd-mmm-yyyy', 'font_size': 11, 'align':'right'})
        #number format
        number = workbook.add_format({'num_format': '', 'font_size': 11, 'align':'right'})
        bold_number = workbook.add_format({'bold': True, 'num_format': '', 'font_size': 11, 'align':'right'})

        #report body
        report_data = data.sales_data()
        #change index to start from 1
        report_data.index = [i for i in range(1, len(report_data) + 1)]
        report_data.to_excel(writer, sheet_name='Sales', index=True, 
                                    startrow=4, startcol=0, header=True)
        
        #hide gridlines
        writer.sheets['Sales'].hide_gridlines(2)

        #report header
        writer.sheets['Sales'].merge_range("A1:F1", "FURNITURE STORE SALES REPORT", title_text)
        writer.sheets['Sales'].merge_range("A2:F2", "1 JAN 2023 - 3 JAN 2023", subtitle_text)
        
        #Profit column
        writer.sheets['Sales'].write('F5', 'Profit', column_header_text)
        #calculate profit
        for i in range(6, 6 + len(report_data) + 1):
            writer.sheets['Sales'].write(f'F{i}', f'=E{i}-D{i}', number)



