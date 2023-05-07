from util import data
import pandas as pd

class SalesReport:
    def __init__(self) -> None:
        pass

    def add_borders(self, writer, start_row, end_row, thick_borders, thin_borders) -> None:
        writer.sheets['Sales'].conditional_format(start_row, 0, start_row, 22, {'type': 'no_blanks', 'format': thick_borders})
        writer.sheets['Sales'].conditional_format(start_row, 0, start_row, 22, {'type': 'blanks', 'format': thick_borders})
        writer.sheets['Sales'].conditional_format(start_row+1, 0, end_row, 22, {'type': 'no_blanks', 'format': thin_borders})
        writer.sheets['Sales'].conditional_format(start_row+1, 0, end_row, 22, {'type': 'blanks', 'format': thin_borders})

    def generate(self) -> None:
        
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
        column_header_text = workbook.add_format({'bold': True, 'font_size': 12, 'align':'center'})
        bold_text = workbook.add_format({'bold': True, 'font_size': 12,})

        # date format
        date_format = workbook.add_format({'num_format': 'dd-mmm-yyyy', 'font_size': 12, 'align':'right'})

        #number format
        number = workbook.add_format({'num_format': '#,##0.00_);(#,##0.00)', 'font_size': 11})
        bold_number = workbook.add_format({'bold': True, 'num_format': '#,##0.00_);(#,##0.00)', 'font_size': 12, 'align':'right'})

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

        #add number formatting to the price columns
        writer.sheets['Sales'].conditional_format(f'D{start_row+2}:E{last_row}', {'type': 'no_blanks', 'format': number})

        #change the date formatting
        writer.sheets['Sales'].conditional_format(f'B{start_row+2}:B{last_row}', {'type': 'no_blanks', 'format': date_format})

        #increase column widths
        writer.sheets['Sales'].set_column('B:F', 15)

        #SHEET 2
        #VISUALIZATIONS
        #create new sheet
        pd.DataFrame().to_excel(writer, sheet_name='Chart', index=True, 
                                    startrow=start_row, startcol=start_col, header=True)
        
        # Create a column chart showing the profit made on each item side by side
        item_profit_bar_chart = workbook.add_chart({"type": "column"})

        # Configure the first series.
        item_profit_bar_chart.add_series(
            {
                "name": "Profit",
                "categories": f"=Sales!$C$5:$C${last_row-1}",
                "values": f"=Sales!$F$5:$F${last_row-1}",
            }
        )
        # Add a chart title and some axis labels.
        item_profit_bar_chart.set_title({"name": "Profit by Items"})
        item_profit_bar_chart.set_x_axis({"name": "Items", 'major_gridlines': {'visible': False}})
        item_profit_bar_chart.set_y_axis({"name": "Profit (NGN)", 'major_gridlines': {'visible': False}})

        #turn off legend
        item_profit_bar_chart.set_legend({'none': True})

        # Insert the chart into the worksheet (with an offset).
        writer.sheets['Chart'].insert_chart("C2", item_profit_bar_chart, {"x_offset": 25, "y_offset": 10})

        #set chart size
        item_profit_bar_chart.set_size({'width': 720, 'height': 576})

        #set chart style
        item_profit_bar_chart.set_style(10)


        #Chart 2: Line Chart showing daily sales vs daily profit
        daily_sales_line_chart = workbook.add_chart({"type": "line"})

        # Configure the first series.
        daily_sales_line_chart.add_series(
            {
                "name": "Sales",
                "categories": f"=Sales!$B$5:$B${last_row-1}",
                "values": f"=Sales!$E$5:$E${last_row-1}",
            }
        )
        # Add a chart title and some axis labels.
        daily_sales_line_chart.set_title({"name": "Profit by Items"})
        daily_sales_line_chart.set_x_axis({"name": "Items", 'major_gridlines': {'visible': False}})
        daily_sales_line_chart.set_y_axis({"name": "Profit (NGN)", 'major_gridlines': {'visible': False}})

        # Insert the chart into the worksheet (with an offset).
        writer.sheets['Chart'].insert_chart("P2", daily_sales_line_chart, {"x_offset": 25, "y_offset": 10})

        #set chart size
        daily_sales_line_chart.set_size({'width': 720, 'height': 576})


        workbook.close()


SalesReport().generate()
