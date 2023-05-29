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

        # text formats
        title_text = workbook.add_format({'bold': True, 'font_size': 20, 'align':'center', 'bg_color':'blue', 'font_color': 'white'})
        subtitle_text = workbook.add_format({'bold': True, 'font_size': 16, 'align':'center', 'bg_color':'blue', 'font_color': 'white'})
        column_header_text = workbook.add_format({'bold': True, 'font_size': 12, 'align':'center'})
        bold_text = workbook.add_format({'bold': True, 'font_size': 12,})

        # date format
        date_format = workbook.add_format({'num_format': 'dd-mmm-yyyy', 'font_size': 12, 'align':'right'})

        #number format
        number = workbook.add_format({'num_format': '#,##0.00_);(#,##0.00)', 'font_size': 11})
        bold_number = workbook.add_format({'bold': True, 'num_format': '#,##0.00_);(#,##0.00)', 'font_size': 12, 'align':'right'})


        # START REPORT BODY

        # sales data
        report_data: pd.DataFrame = data.sales_data()

        # set report start row and column
        start_row = 3
        start_col = 0

        # change dataframe index to start from 1
        #write dataframe to excel 'Sales' sheet in the wookbook
        report_data.to_excel(writer, sheet_name='Sales', index=True, 
                                    startrow=start_row, startcol=start_col, header=True)
        
        #hide gridlines
        writer.sheets['Sales'].hide_gridlines(2)

        #report title
        writer.sheets['Sales'].merge_range("A1:F1", "FURNITURE STORE SALES REPORT", title_text)
        writer.sheets['Sales'].merge_range("A2:F2", "1 JAN 2023 - 3 JAN 2023", subtitle_text)
        
        #Profit column
        writer.sheets['Sales'].write(f'F{start_row+1}', 'Profit', column_header_text)
        #calculate profit
        for i in range(start_row+2, start_row+2 + len(report_data)):
            writer.sheets['Sales'].write_formula(f'F{i}', f'=E{i}-D{i}', number)

        #Total row
        last_row = len(report_data) + start_row + 2
        writer.sheets['Sales'].write(f'C{last_row}', 'Total', bold_text)
        writer.sheets['Sales'].write_formula(f'E{last_row}', f'=SUM(E{start_row+2}:E{last_row-1})', bold_number)
        writer.sheets['Sales'].write_formula(f'F{last_row}', f'=SUM(F{start_row+2}:F{last_row-1})', bold_number)

        #ADD BORDERS
        #add thick borders to report header row where cells are filled
        writer.sheets['Sales'].conditional_format(start_row, 0, start_row, 5, 
                                                  {'type': 'no_blanks', 'format': thick_borders})
        #add thick borders to report header row where cells are empty
        writer.sheets['Sales'].conditional_format(start_row, 0, start_row, 5, 
                                                  {'type': 'blanks', 'format': thick_borders})
        #add thin borders body of report where cells are filled
        writer.sheets['Sales'].conditional_format(start_row+1, 0, last_row-1, 5, 
                                                  {'type': 'no_blanks', 'format': thin_borders})
        #add thin borders body of report where cells are blank
        writer.sheets['Sales'].conditional_format(start_row+1, 0, last_row-1, 5, 
                                                  {'type': 'blanks', 'format': thin_borders})

        #add number formatting to the price columns
        writer.sheets['Sales'].conditional_format(f'D{start_row+2}:E{last_row}', {'type': 'no_blanks', 'format': number})

        #change the date formatting
        writer.sheets['Sales'].conditional_format(f'B{start_row+2}:B{last_row}', {'type': 'no_blanks', 'format': date_format})

        #increase column widths
        writer.sheets['Sales'].set_column('B:F', 15)

        #SHEET 2
        #VISUALIZATIONS
        #create new sheet
        
        #Aggregate data by items and calculate total cost and sales prices
        items_cost_sales_df = report_data.groupby('Items')[['Cost Price', 'Sale Price']].sum() 
        
        #write data to new sheey
        items_cost_sales_df.to_excel(writer, sheet_name='Chart', index=True, 
                                    startrow=start_row, startcol=start_col, header=True)
        
        chart_data_end_row = start_row + len(items_cost_sales_df) + 1

        #ADD BORDERS
        #add thick borders to report header row where cells are filled
        #add thin borders body of report where cells are filled
        writer.sheets['Chart'].conditional_format(start_row+1, 0, chart_data_end_row-1, 2, 
                                                  {'type': 'no_blanks', 'format': thin_borders})
        #add thin borders body of report where cells are blank
        writer.sheets['Chart'].conditional_format(start_row+1, 0, chart_data_end_row-1, 2, 
                                                  {'type': 'blanks', 'format': thin_borders})

        #add number formatting to the price columns
        writer.sheets['Chart'].conditional_format(f'B{start_row+2}:C{chart_data_end_row}', {'type': 'no_blanks', 'format': number})

        #increase column widths
        writer.sheets['Chart'].set_column('B:C', 15)
        
        # Create a column chart object
        items_bar_chart = workbook.add_chart({"type": "column"})

        # Configure the first series.
        items_bar_chart.add_series(
            {
                "name": "=Chart!B4",
                "categories": f"=Chart!$A$5:$A{chart_data_end_row}",
                "values": f"=Chart!$B$5:$B${chart_data_end_row}",
            }
        )
        # Configure the second series.
        items_bar_chart.add_series(
            {   
                "name":"=Chart!C4",
                "categories": f"=Chart!$A$5:$A{chart_data_end_row}",
                "values": f"=Chart!$C$5:$C${chart_data_end_row}",
            }
        )

        # Add a chart title and some axis labels.
        items_bar_chart.set_title({"name": "Cost and Sales Price by Item"})
        items_bar_chart.set_x_axis({"name": "Items", 'major_gridlines': {'visible': False}})
        items_bar_chart.set_y_axis({"name": "Price (NGN)", 'major_gridlines': {'visible': False}})

        # Set an Excel chart style.
        items_bar_chart.set_style(10)

        #set chart size
        items_bar_chart.set_size({'width': 720, 'height': 576})

        # Insert the chart into the worksheet (with an offset).
        writer.sheets['Chart'].insert_chart("F2", items_bar_chart, {"x_offset": 25, "y_offset": 10})

        # close and save workbook
        workbook.close()
