from datetime import datetime
import math

import matplotlib.pyplot as plt
import numpy as np
import openpyxl
from io import BytesIO
from openpyxl.styles import Border, Side
from openpyxl.styles import Alignment, Font
from sqlalchemy import text
from openpyxl.styles import PatternFill
from matplotlib.font_manager import FontProperties
from constant import CategoryEnum
import data_service
import platform
from PIL import Image as PILImage

lotes_report_format = {
    "title": "B2:R2",
    "summary":"B3:R3",
    "profit_summary": "B4:E4",
    "profit_summary_chart": "F4:R4",
    "by_segment_summary": "B5:E5",
    "by_segment_chart": "F5:R5",
    "financial_statements": "B6:R40",
}

argosy_report_format = {
    "title": "B2:T2",
    "summary":"B3:T35"
}

# 國內財報明細項目，資料庫名稱對應顯示文字
financial_statement_tw_items = {
    "operating_revenue": "營業收入A",
    "operating_costs": "營業成本B",
    "gross_profit": "營業毛利 C=A-B",
    "gross_profit_margin": "毛利率 D=C/A",
    "operating_expenses": "營業費用 E",
    "operating_income": "營業利益 F=C-E",
    "depreciation": "折舊",
    "amortization": "攤提",
    "ebitda": "EBITDA",
    "total_nonoperating_income": "總營業外收入 G=H+I+J",
    "interest_income": "利息收入 H",
    "net_investment_income": "投資利益淨額 I",
    "other_nonoperating_income": "'其他營業外收入 J",
    "total_nonoperating_expenses": "總營業外費用 K=L+M+N",
    "interest_expenses": "利息費用 L",
    "investment_losses": "投資損失 M",
    "other_nonoperating_expenses": "其他營業外費用 N",
    "pretax_income": "稅前純益 O=F+G-K",
    "pretax_net_profit_margin": "稅前凈利率 P=O/A",
    "income_tax_expense": "所得稅費用[利益] Q",
    "minority_interest_income": "少數股東損益 R",
    "net_income": "稅後淨利 S=O-Q-R",
    "net_profit_margin": "稅後淨利率 T=S/A"
}

# 國內財報明細項目，百分比數值的項目
financial_statement_tw_items_percentage_format = {
    "gross_profit_margin": 1,
    "pretax_net_profit_margin": 1,
    "net_profit_margin": 1
}

# 國內財報明細項目，改變底色的項目
financial_statement_tw_items_color_format = {
    "title": "E2EFDA",
    "operating_revenue": "E2EFDA",
    "gross_profit": "E2EFDA",
    "gross_profit_margin": "E2EFDA",
    "operating_income": "E2EFDA",
    "total_nonoperating_income": "E2EFDA",
    "total_nonoperating_expenses": "E2EFDA",
    "pretax_income": "E2EFDA",
    "pretax_net_profit_margin": "E2EFDA",
    "net_income": "E2EFDA",
    "net_profit_margin": "E2EFDA"
}

# 國外財報明細項目，資料庫名稱對應顯示文字
financial_statement_foreign_items = {
    "net_sales": "營業收入A",
    "qoq_growth": "成長Q/Q",
    "yoy_growth": "成長Y/Y",
    "cost_of_sales": "銷售成本 B",
    "gross_profit": "毛利潤 C=A-B",
    "gross_profit_margin": "毛利率 D=C/A",
    "selling_general_administrative_expenses": "管銷研 E",
    "selling_general_administrative_expenses_percentage": """管銷研% F=E/A""",
    "operating_income": "稅前利潤額 G=C-E",
    "pretax_profit_margin": "稅前利潤率 H=G/A",
    "dividend_payment": "股息金額 I",
    "other_income_and_expenses": "營業外收支 J",
    "total_interest_and_other_expenses": "TOTAL利息及其他費用K=I+J",
    "pretax_income": "稅前收入 L=G-K",
    "pretax_net_profit_margin": "稅前利潤率 M=L/A",
    "income_tax_expense": "稅捐 N",
    "effective_tax_rate": "稅率 O",
    "shareholders_equity": "股東權益 P",
    "net_income": "淨收入 Q=L-N-P",
    "net_profit_margin": "淨利率R=Q/A"   
}

# 國外財報明細項目，百分比數值的項目
financial_statement_foreign_items_percentage_format = {
    "qoq_growth": 1,
    "yoy_growth": 1,
    "gross_profit_margin": 1,
    "pretax_profit_margin": 1,
    "pretax_profit_margin": 1,
    "effective_tax_rate": 1,
    "net_profit_margin": 1
}

# 國外財報明細項目，改變底色的項目
financial_statement_foreign_items_color_format = {
    "title": "E2EFDA",
    "net_sales": "E2EFDA",
    "gross_profit": "E2EFDA",
    "gross_profit_margin": "E2EFDA",
    "operating_income": "E2EFDA",
    "pretax_profit_margin": "E2EFDA",
    "pretax_income": "E2EFDA",
    "pretax_net_profit_margin": "E2EFDA",
    "net_income": "E2EFDA",
    "net_profit_margin": "E2EFDA"
}

# 中文字型檔，圖表的Legend需要
os_name = platform.system()
if os_name == 'Windows':
    print("当前操作系统为 Windows")
    font = FontProperties(fname=r"C:\Windows\Fonts\mingliu.ttc", size=14)
else :
    print("当前操作系统为 Linux")
    font = FontProperties(fname=r"/usr/share/fonts/truetype/wqy/wqy-microhei.ttc", size=14)

# excel邊框格式
border_style_tb = Border(top=Side(style='thin'),bottom=Side(style='thin'),)
border_style_ltb = Border(top=Side(style='thin'),bottom=Side(style='thin'),left=Side(style='thin'),)
border_style_rtb = Border(top=Side(style='thin'),bottom=Side(style='thin'),right=Side(style='thin'),)
border_style_r = Border(right=Side(style='thin'),)
border_style_l = Border(left=Side(style='thin'),)
border_style_lt = Border(top=Side(style='thin'),left=Side(style='thin'),)
border_style_rt = Border(top=Side(style='thin'),right=Side(style='thin'),)
border_style_lb = Border(left=Side(style='thin'),bottom=Side(style='thin'),)
border_style_rb = Border(right=Side(style='thin'),bottom=Side(style='thin'),)
border_style_t = Border(top=Side(style='thin'),)
border_style_b = Border(bottom=Side(style='thin'),)
border_style_all = Border(top=Side(style='thin'),bottom=Side(style='thin'),left=Side(style='thin'),right=Side(style='thin'),)

def __get_financial_statement_items_dict(category: CategoryEnum):
    if category == CategoryEnum.TW:
        return financial_statement_tw_items
    else:
        return financial_statement_foreign_items
    
def __get_financial_statement_items_percentage_format_dict(category: CategoryEnum):
    if category == CategoryEnum.TW:
        return financial_statement_tw_items_percentage_format
    else:
        return financial_statement_foreign_items_percentage_format

def __get_financial_statement_items_color_format_dict(category: CategoryEnum):
    if category == CategoryEnum.TW:
        return financial_statement_tw_items_color_format
    else:
        return financial_statement_foreign_items_color_format        

def __is_multiple_row(sheet, cell_range):
    rows = sheet[cell_range]
    if len(rows) > 1:
        return True
    return False

def __draw_first_row_border(row):
    is_first_cell = True
    for cell in row:
        if is_first_cell:
            cell.border = border_style_lt
            is_first_cell = False
        else:
            cell.border = border_style_t
    cell.border = border_style_rt


def __draw_last_row_border(row):
    is_first_cell = True
    for cell in row:
        if is_first_cell:
            cell.border = border_style_lb
            is_first_cell = False
        else:
            cell.border = border_style_b
    cell.border = border_style_rb

def __draw_middle_row_border(row):
    is_first_cell = True
    for cell in row:
        if is_first_cell:
            cell.border = border_style_l
            is_first_cell = False
    cell.border = border_style_r

def __draw_single_row_border(row):
    is_first_cell = True
    for cell in row:
        if is_first_cell:
            cell.border = border_style_ltb
            is_first_cell = False
        else:
            cell.border = border_style_tb
    cell.border = border_style_rtb

def __draw_border(sheet, cell_range):
    if __is_multiple_row(sheet, cell_range):
        is_first_row = True
        for row in sheet[cell_range]:
            if is_first_row:
                __draw_first_row_border(row)
                is_first_row = False
            else:
                __draw_middle_row_border(row)
        __draw_last_row_border(row)

    else:
        for row in sheet[cell_range]:
            __draw_single_row_border(row)
        

def __set_cell_height(sheet, cell):
    max_height = 0
    over_length = 0
    length_boundary = 80
    if cell.value:
        lines = str(cell.value).split('\n')
        height = max(len(lines) + 1, 1) * 14  # Adjust this value as needed
        max_height = max(max_height, height)

        for line in lines:
            if len(line) > length_boundary:
                over_length += 1         
    if max_height > 0:
        sheet.row_dimensions[cell.row].height = max_height + over_length
        cell.alignment = Alignment(wrap_text=True, vertical='top')

def __set_cell_blod(cell):
    cell.font = Font(bold=True)

def __set_cell_alignment_center(cell):
    cell.alignment = Alignment(horizontal='center')

def __get_first_cell(sheet, cell_range):
    for row in sheet[cell_range]:
        for cell in row:
            return cell

def __is_numeric(data):
    numeric_types = (int, float, complex)
    return isinstance(data, numeric_types)

def __get_formated_number(number):
    return "{:,}".format(math.floor(number / 1000))

def __get_percentage_number(number):
    return "{:.1%}".format(number)

def __get_financial_statements_sum(company_code, fiscal_year, category: CategoryEnum) -> dict:
    result = {}
    data_dict_list = []
    financial_statement_items = __get_financial_statement_items_dict(category)

    if category == CategoryEnum.TW:
        data_list = data_service.get_financial_statements_by_year(company_code, fiscal_year)
    else:
        data_list = data_service.get_financial_statements_foreign_by_year(company_code, fiscal_year)


    for data in data_list:
        if data is not None:
            data_dict_list.append(data.as_dict())

    for k in financial_statement_items.keys():
        sum = 0
        for data_dict in data_dict_list:
            if k in data_dict:
                if __is_numeric(data_dict[k]):
                    sum += data_dict[k]
        result[k] = sum
    return result

def __draw_financial_statements_items_name(sheet, start_cell, items):
    row_index = start_cell.row
    column_index = start_cell.column
    sheet.column_dimensions['C'].auto_size = True
    number_of_items = len(items)
    sheet.cell(row=row_index, column=column_index).value = "科目"

    for (k, v) in items.items():
        row_index += 1
        sheet.cell(row=row_index, column=column_index).value = v
        sheet.cell(row=row_index, column=column_index).border = border_style_all

        # style
    row_index = start_cell.row
    sheet.cell(row=row_index, column=column_index).border = border_style_all
    sheet.cell(row=row_index, column=column_index).alignment = Alignment(horizontal='center', vertical='center')
    sheet.cell(row=row_index, column=column_index).fill = PatternFill(fill_type='solid', start_color=financial_statement_tw_items_color_format["title"])

    for k in items.keys():
        row_index += 1
        sheet.cell(row=row_index, column=column_index).border = border_style_all
        sheet.cell(row=row_index, column=column_index).alignment = Alignment(horizontal='center', vertical='center')
        if k in financial_statement_tw_items_color_format:
            sheet.cell(row=row_index, column=column_index).fill = PatternFill(fill_type='solid', start_color=financial_statement_tw_items_color_format[k])
        

def __draw_financial_statements_items(sheet, start_cell, comany_code, fiscal_year, quarter, category: CategoryEnum):    
    data_dict = None

    # get value
    if category == CategoryEnum.TW:    
        data = data_service.get_financial_statements_by_quarter(comany_code, fiscal_year, quarter)
    else:
        data = data_service.get_financial_statements_foreign_by_quarter(comany_code, fiscal_year, quarter)

    if data is not None:
        data_dict = data.as_dict()
    
    # draw cell
    __draw_financial_statements_cells(sheet, start_cell, fiscal_year, quarter, data_dict, category)

def __draw_financial_statements_cells(sheet, start_cell, fiscal_year, quarter, data_dict, category):
    row_index = start_cell.row
    column_index = start_cell.column
    sheet.cell(row=row_index, column=column_index).value = str(fiscal_year) + "Q" + str(quarter)

    if category == CategoryEnum.TW:
        financial_statement_items = __get_financial_statement_items_dict(category)
        financial_statement_items_percentage_format = __get_financial_statement_items_percentage_format_dict(category)
        financial_statement_items_color_format = __get_financial_statement_items_color_format_dict(category)    
    else:
        financial_statement_items = __get_financial_statement_items_dict(category)
        financial_statement_items_percentage_format = __get_financial_statement_items_percentage_format_dict(category)
        financial_statement_items_color_format = __get_financial_statement_items_color_format_dict(category)

    # value
    if data_dict is not None:
        for (k, v) in financial_statement_items.items():
            row_index += 1
            if k in data_dict and data_dict[k] is not None:
                if k in financial_statement_items_percentage_format:
                    sheet.cell(row=row_index, column=column_index).value = __get_percentage_number(data_dict[k])
                else:
                    sheet.cell(row=row_index, column=column_index).value = __get_formated_number(data_dict[k])
    
    # style
    row_index = start_cell.row
    sheet.cell(row=row_index, column=column_index).border = border_style_all
    sheet.cell(row=row_index, column=column_index).alignment = Alignment(horizontal='center', vertical='center')
    sheet.cell(row=row_index, column=column_index).fill = PatternFill(fill_type='solid', start_color=financial_statement_items_color_format["title"])

    for k in financial_statement_items.keys():
        row_index += 1
        sheet.cell(row=row_index, column=column_index).border = border_style_all
        sheet.cell(row=row_index, column=column_index).alignment = Alignment(horizontal='center', vertical='center')
        if k in financial_statement_items_color_format:
            sheet.cell(row=row_index, column=column_index).fill = PatternFill(fill_type='solid', start_color=financial_statement_items_color_format[k])

def __draw_draw_financial_statements_summary(sheet, start_cell, comany_code, fiscal_year, category):
    # get value
    data_dict = __get_financial_statements_sum(comany_code, fiscal_year, category)
    
    # draw cell
    __draw_financial_statements_sum_cells(sheet, start_cell, fiscal_year, data_dict, category)

def __draw_financial_statements_sum_cells(sheet, start_cell, fiscal_year, data_dict, category: CategoryEnum):
    row_index = start_cell.row
    column_index = start_cell.column

    # value
    sheet.cell(row=row_index, column=column_index).value = str(fiscal_year)

    if category == CategoryEnum.TW:
        financial_statement_items = __get_financial_statement_items_dict(category)
        financial_statement_items_percentage_format = __get_financial_statement_items_percentage_format_dict(category)
        financial_statement_items_color_format = __get_financial_statement_items_color_format_dict(category)
    
    else:
        financial_statement_items = __get_financial_statement_items_dict(category)
        financial_statement_items_percentage_format = __get_financial_statement_items_percentage_format_dict(category)
        financial_statement_items_color_format = __get_financial_statement_items_color_format_dict(category)
    
    for k in financial_statement_items.keys():
        row_index += 1
        if k in data_dict:
            if k in financial_statement_items_percentage_format:
                sheet.cell(row=row_index, column=column_index).value = __get_percentage_number(data_dict[k])
            else:
                sheet.cell(row=row_index, column=column_index).value = __get_formated_number(data_dict[k])
        
    
    # style
    row_index = start_cell.row
    sheet.cell(row=row_index, column=column_index).border = border_style_all
    sheet.cell(row=row_index, column=column_index).alignment = Alignment(horizontal='center', vertical='center')
    sheet.cell(row=row_index, column=column_index).fill = PatternFill(fill_type='solid', start_color=financial_statement_items_color_format["title"])

    for k in financial_statement_items.keys():
        row_index += 1
        sheet.cell(row=row_index, column=column_index).border = border_style_all
        sheet.cell(row=row_index, column=column_index).alignment = Alignment(horizontal='center', vertical='center')
        if k in financial_statement_items_color_format:
            sheet.cell(row=row_index, column=column_index).fill = PatternFill(fill_type='solid', start_color=financial_statement_items_color_format[k])         

# Todo 將起始的cell傳入    
def __draw_financial_statements_tw(sheet, comany_code, input_fiscal_year, input_quarter):
    start_cell = sheet['C7']
    row_index = start_cell.row
    column_index = start_cell.column

    __draw_financial_statements_items_name(sheet, sheet.cell(row=row_index, column=column_index), financial_statement_tw_items)

    # 畫季報
    quarter_array = __get_report_quarter(input_fiscal_year, input_quarter)
    for (year, quarter) in quarter_array:
        column_index += 1
        __draw_financial_statements_items(sheet, sheet.cell(row=row_index, column=column_index), comany_code, year, quarter, CategoryEnum.TW)

    # 畫年報
    year_array = __get_report_year(input_fiscal_year)
    for year in year_array:
        column_index += 1
        __draw_draw_financial_statements_summary(sheet, sheet.cell(row=row_index, column=column_index), comany_code, year, CategoryEnum.TW)

# Todo 將起始的cell傳入   
def __draw_financial_statements_foreign(sheet, comany_code, input_fiscal_year, input_quarter):
    start_cell = sheet['C7']
    row_index = start_cell.row
    column_index = start_cell.column

    __draw_financial_statements_items_name(sheet, sheet['C7'], financial_statement_foreign_items)

    # 畫季報
    quarter_array = __get_report_quarter(input_fiscal_year, input_quarter)
    for (year, quarter) in quarter_array:
        column_index += 1
        __draw_financial_statements_items(sheet, sheet.cell(row=row_index, column=column_index), comany_code, year, quarter, CategoryEnum.FOREIGN)

    # 畫年報
    year_array = __get_report_year(input_fiscal_year)
    for year in year_array:
        column_index += 1
        __draw_draw_financial_statements_summary(sheet, sheet.cell(row=row_index, column=column_index), comany_code, year, CategoryEnum.FOREIGN)    

def __autofit_column(ws, column_letter):
    column_index = openpyxl.utils.column_index_from_string(column_letter)
    max_length = 0
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=column_index, max_col=column_index):
        for cell in row:
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))

    # Set the column width based on the maximum length
    ws.column_dimensions[column_letter].width = max_length * 1.75


def __get_combo_chart(categories, bar_chart_data, line_chart_data_1, line_chart_data_2):
    if os_name == 'Windows':
        plt.rcParams['font.family'] = 'mingliu'
    else: 
        plt.rcParams['font.family'] = 'WenQuanYi Micro Hei'

    
    fig, ax1 = plt.subplots()

    # Plotting the bar chart
    bars = ax1.bar(categories, bar_chart_data, color='#C5E0B4', label='營業收入', align='center', width=0.5)

    # Display values on the bar chart
    for bar, value in zip(bars, bar_chart_data):
        #ax1.text(bar.get_x() + bar.get_width() / 2, value + 0.5, str(value), ha='center', va='bottom', color='black', fontsize=10)
        ax1.text(bar.get_x() + bar.get_width() / 2, 3000, __get_formated_number(value), ha='center', va='bottom', color='black', fontsize=10)

    # Create a secondary y-axis for the line charts
    ax2 = ax1.twinx()

    # Plotting the line charts
    line1, = ax2.plot(categories, line_chart_data_1, color='#FFDDA4', marker='o', label='毛利', markersize=8)
    line2, = ax2.plot(categories, line_chart_data_2, color='#5B9BD5', marker='s', label='淨利', markersize=8)

    ax2.set_ylabel('', color='r', fontproperties=font)

    # Display values on the line charts
    for x, y in zip(categories, line_chart_data_1):
        ax2.text(x, y + 5, __get_percentage_number(y/100), ha='center', va='bottom', color='black', fontsize=10)
    for x, y in zip(categories, line_chart_data_2):
        ax2.text(x, y + 5, __get_percentage_number(y/100), ha='center', va='bottom', color='black', fontsize=10)

    # Combine legends from both axes
    lines = [bars, line1, line2]
    labels = [line.get_label() for line in lines]
    ax2.legend(lines, labels, loc='upper right')

    max_value = max(bar_chart_data)
    ax1.set_ylim(0, max_value * 3)
    ax2.set_ylim(-100, 100)  # Set the y-axis limits for the line charts

    ax1.set_yticklabels([])
    ax2.set_yticklabels([])

    # Show the plot
    plt.title('By季獲利率趨勢 NT$m')
    plt.tight_layout()  # Ensures that the plots do not overlap
    #plt.show()

    buffer = BytesIO()
    plt.savefig(buffer, format='png')
    buffer.seek(0)
    return buffer

def __get_operating_combo_chart(categories, bar_chart_data_1, bar_chart_data_2, bar_chart_data_3, line_chart_data_1, line_chart_data_2):
    if os_name == 'Windows':
        plt.rcParams['font.family'] = 'mingliu'
    else:
        plt.rcParams['font.family'] = 'WenQuanYi Micro Hei'

    # Create figure and axes for bar chart
    fig, ax1 = plt.subplots()

    # 计算每个柱状图的宽度
    bar_width = 0.25
    index = np.arange(len(categories))

    # Plotting the bar chart
    bars1 = plt.bar(index - bar_width, bar_chart_data_1, width=bar_width, label='營收', color='#5B9BD5')
    bars2 = plt.bar(index, bar_chart_data_2, width=bar_width, label='毛利', color = '#ED7D31')
    bars3 = plt.bar(index + bar_width, bar_chart_data_3, width=bar_width, label='淨利', color = '#FFC000')

    # Display values on the bar chart
    for bars in [bars1, bars2, bars3]:
        for bar in bars:
            yval = bar.get_height()
            plt.text(bar.get_x() + bar.get_width() / 2, yval / 2, __get_formated_number(yval), ha='center', va='bottom')

    # Create a secondary y-axis for the line charts
    ax2 = ax1.twinx()

    # Plotting the line charts
    line1, = ax2.plot(categories, line_chart_data_1, color='#A5A5A5', marker='o', label='毛利', markersize=8)
    line2, = ax2.plot(categories, line_chart_data_2, color='#4472C4', marker='s', label='淨利', markersize=8)
    ax2.set_ylabel('', color='r', fontproperties=font)

    # Display values on the line charts
    for x, y in zip(categories, line_chart_data_1):
        ax2.text(x, y + 5, __get_percentage_number(y/100), ha='center', va='bottom', color='black', fontsize=10)
    for x, y in zip(categories, line_chart_data_2):
        ax2.text(x, y + 5, __get_percentage_number(y/100), ha='center', va='bottom', color='black', fontsize=10)

    # Combine legends from both axes
    lines = [bars1, bars2, bars3, line1, line2]
    labels = [line.get_label() for line in lines]
    ax2.legend(lines, labels, loc='lower center', bbox_to_anchor=(0.5, -0.3), ncol=5)

    max_value = max(bar_chart_data_1)
    ax1.set_ylim(0, max_value * 3)
    ax2.set_ylim(-100, 100)  # Set the y-axis limits for the line charts
    ax1.set_yticklabels([])
    ax2.set_yticklabels([])

    # Show the plot
    plt.title('營運指標趨勢')
    plt.tight_layout()  # Ensures that the plots do not overlap
    #plt.show()    

    buffer = BytesIO()
    plt.savefig(buffer, format='png')
    buffer.seek(0)
    return buffer

#def __extract_quarter(data):
#    result = []
#    for row in data:
#        result.append(str(row[0]) + "Q" + str(row[1]))
#    return result

#def __extract_quarter_data(data):
#    result = []
#    for row in data:
#        result.append(row[2]) 
#    return result

def __extract_year_data(data):
    result = []
    for row in data:
        result.append(row[0]) 
    return result

def __get_quarter_profit_chart(company_code, quarter_array, category: CategoryEnum): 
    start_fiscal_year =  quarter_array[0][0]  
    end_fiscal_year = quarter_array[-1][0]      
    if category == CategoryEnum.TW:
        financial_statements = data_service.get_financial_statements_by_year_between(company_code, start_fiscal_year, end_fiscal_year)
    else : 
        financial_statements = data_service.get_financial_statements_foreign_by_year_between(company_code, start_fiscal_year, end_fiscal_year)

    financial_statements_after_filter = []
    # 將quarter_array轉為dict作為filter不在quarter_array的資料
    dict = {f"{year}Q{quarter}": (year, quarter) for year, quarter in quarter_array}  
    for data in financial_statements:
        if dict.get(f"{data.fiscal_year}Q{data.quarter}"):
            financial_statements_after_filter.append(data)

    if category == CategoryEnum.TW:
        operating_revenue_array = [ data.operating_revenue for data in financial_statements_after_filter] 
    else:  
        operating_revenue_array = [ data.net_sales for data in financial_statements_after_filter] 

    gross_profit_margin_array = [ data.gross_profit_margin for data in financial_statements_after_filter] 
    gross_profit_margin_percentage = [i * 100 for i in gross_profit_margin_array]
    net_profit_margin_array = [ data.net_profit_margin for data in financial_statements_after_filter] 
    net_profit_margin_percentage = [i * 100 for i in net_profit_margin_array]
    quarter_array_after_filter = [f"{data.fiscal_year}Q{data.quarter}" for data in financial_statements_after_filter]

    return __get_combo_chart(quarter_array_after_filter, operating_revenue_array, gross_profit_margin_percentage, net_profit_margin_percentage)

#def __filter_data_with_quarter_array(data_array, quarter_array):
#    result = []
#    dict = {f"{year}Q{quarter}": (year, quarter) for year, quarter in quarter_array}
#    #print(dict)
#    for data in data_array:
#        if dict.get(f"{data[0]}Q{data[1]}"):
#            result.append(data)
    
#    return result

def __get_year_profit_chart(company_code, start_fiscal_year, end_fiscal_year, category: CategoryEnum):
    if category == CategoryEnum.TW:
        operating_revenue = data_service.get_year_operating_revenue(company_code, start_fiscal_year, end_fiscal_year)
        gross_profit_margin = data_service.get_year_gross_profit_margin(company_code, start_fiscal_year, end_fiscal_year)
        net_profit_margin = data_service.get_year_net_profit_margin(company_code, start_fiscal_year, end_fiscal_year)
    else:
        operating_revenue = data_service.get_year_net_sales_foreign(company_code, start_fiscal_year, end_fiscal_year)
        gross_profit_margin = data_service.get_year_gross_profit_margin_foreign(company_code, start_fiscal_year, end_fiscal_year)
        net_profit_margin = data_service.get_year_net_profit_margin_foreign(company_code, start_fiscal_year, end_fiscal_year)

    categories = [str(i) for i in range(start_fiscal_year, end_fiscal_year + 1)]
    operating_revenue_array = __extract_year_data(operating_revenue)

    gross_profit_margin_array = __extract_year_data(gross_profit_margin)
    gross_profit_margin_percentage = [i * 100 for i in gross_profit_margin_array]

    net_profit_margin_array = __extract_year_data(net_profit_margin)
    net_profit_margin_percentage = [i * 100 for i in net_profit_margin_array]

    return __get_combo_chart(categories, operating_revenue_array, gross_profit_margin_percentage, net_profit_margin_percentage)

def __get_year_operating_chart(company_code, start_fiscal_year, end_fiscal_year, category: CategoryEnum):
    if category == CategoryEnum.TW:
        operating_revenue = data_service.get_year_operating_revenue(company_code, start_fiscal_year, end_fiscal_year)
        gross_profit = data_service.get_year_gross_profit(company_code, start_fiscal_year, end_fiscal_year)
        net_income = data_service.get_year_net_income(company_code, start_fiscal_year, end_fiscal_year)
        gross_profit_margin = data_service.get_year_gross_profit_margin(company_code, start_fiscal_year, end_fiscal_year)
        net_profit_margin = data_service.get_year_net_profit_margin(company_code, start_fiscal_year, end_fiscal_year)
    else:
        operating_revenue = data_service.get_year_net_sales_foreign(company_code, start_fiscal_year, end_fiscal_year)
        gross_profit = data_service.get_year_gross_profit_foreign(company_code, start_fiscal_year, end_fiscal_year)
        net_income = data_service.get_year_net_income_foreign(company_code, start_fiscal_year, end_fiscal_year)
        gross_profit_margin = data_service.get_year_gross_profit_margin_foreign(company_code, start_fiscal_year, end_fiscal_year)
        net_profit_margin = data_service.get_year_net_profit_margin_foreign(company_code, start_fiscal_year, end_fiscal_year)

    categories = [str(i) for i in range(start_fiscal_year, end_fiscal_year + 1)]
    operating_revenue_array = __extract_year_data(operating_revenue)
    gross_profit_array = __extract_year_data(gross_profit)
    net_income_array = __extract_year_data(net_income)
    gross_profit_margin_array = __extract_year_data(gross_profit_margin)
    gross_profit_margin_percentage_array = [i * 100 for i in gross_profit_margin_array]
    net_profit_margin_array = __extract_year_data(net_profit_margin)
    net_profit_margin_percentage_arrary = [i * 100 for i in net_profit_margin_array]

    return __get_operating_combo_chart(categories, operating_revenue_array, gross_profit_array, net_income_array, gross_profit_margin_percentage_array, net_profit_margin_percentage_arrary)

def __get_report_year(input_fiscal_year):
    """
    取得要產生財報年報部分的年份
    Returns a list of integers representing the current year and the two previous and two subsequent years.
    """
    result = []
    for i in range(-1, 2):
        result.append(input_fiscal_year + i)
    
    return result

def __get_quarter(input_fiscal_year, input_quarter, diff):
    """
    Calculates the year and quarter based on a given difference.

    Args:
        diff (int): The difference in quarters from the current quarter.

    Returns:
        tuple: A tuple containing the year and quarter.

    Example:
        >>> __get_quarter(1)
        (2022, 4)
    """
    current_year = input_fiscal_year
    current_quarter = input_quarter
    

    diff_year = math.floor(abs(diff) / 4)
    remainder = abs(diff) % 4

    if diff < 0:
        remainder = - remainder
        diff_year = - diff_year

    year = current_year + diff_year
    quarter = current_quarter
    if current_quarter + remainder > 4:
        year += 1
    elif current_quarter + remainder < 1:
        year -= 1

    quarter = (current_quarter + diff) % 4 or 4 # current_quarter + diff

    return(year, quarter)

def __get_report_quarter(input_fiscal_year, input_quarter):
    quarters_array = []

    # 前8個季度(含本季度) + 後4個季度
    for i in range(-7, 5):        
        quarters_array.append(__get_quarter(input_fiscal_year, input_quarter, i))         

    return quarters_array

def __get_chart_year(input_fiscal_year):
    """
    取得要產生圖表的年份
    Returns a list of integers representing the current year and the two previous and two subsequent years.
    """
    result = []

    for i in range(-3, 0):
        result.append(input_fiscal_year + i)
    
    return result
 
def __get_chart_quarter(input_fiscal_year, input_quarter):
    quarters_array = []

    # 前8個季度(含本季度)
    for i in range(-7, 1):        
        quarters_array.append(__get_quarter(input_fiscal_year, input_quarter, i))         

    return quarters_array

def create_lotes_style_financial_statements(sheet, company_code, input_fiscal_year, input_quarter, report_text, category: CategoryEnum):
    # 需要上傳的暫存檔
    temporary_file = []

    for(k, v) in lotes_report_format.items():
        # 合併儲存格
        if not __is_multiple_row(sheet, v):
            sheet.merge_cells(v)
        # 劃格線
        __draw_border(sheet, v)
        # 填入文字
        if report_text.get(k):
            __get_first_cell(sheet, v).value = report_text[k]
            __set_cell_height(sheet, __get_first_cell(sheet, v))    
        # 處理標題格式
        if k == "title":
            __set_cell_blod(__get_first_cell(sheet, v))    
            __set_cell_alignment_center(__get_first_cell(sheet, v))
    
    # 填入財報excel資料
    if category == CategoryEnum.TW:
        __draw_financial_statements_tw(sheet, company_code, input_fiscal_year, input_quarter)
    else:
        __draw_financial_statements_foreign(sheet, company_code, input_fiscal_year, input_quarter)

    # 填入獲利分析圖表  
    # 年度分析圖    
    year_array = __get_chart_year(input_fiscal_year)    
    buffer = __get_year_profit_chart(company_code, year_array[0], year_array[-1], category)

    img = openpyxl.drawing.image.Image(buffer)
    img.width = 300
    img.height = 300
    chart_location = __get_first_cell(sheet, lotes_report_format['profit_summary_chart'])
    sheet.add_image(img, chart_location.coordinate)

    # 保存圖表
    file_name = f"output/{company_code}_year_chart.png"
    pil_image = PILImage.open(buffer)
    pil_image.save(file_name)
    temporary_file.append(file_name)

    # 季度分析圖  
    quarter_array = __get_chart_quarter(input_fiscal_year, input_quarter)
    buffer = __get_quarter_profit_chart(company_code, quarter_array, category)

    img = openpyxl.drawing.image.Image(buffer)
    img.width = 450
    img.height = 300
    row = chart_location.row
    col = chart_location.column + 5
    chart_location = sheet.cell(row, col)
    sheet.add_image(img, chart_location.coordinate)
    # 調整row高度
    # image.height的單位是pixel, cell.height的單位是point, 1 pixel = 0.75 point
    row_height = max(img.height * 0.75, sheet.row_dimensions[chart_location.row].height)
    sheet.row_dimensions[chart_location.row].height = row_height

    # 保存圖表
    file_name = f"output/{company_code}_quarter_chart.png"
    pil_image = PILImage.open(buffer)
    pil_image.save(file_name)
    temporary_file.append(file_name)

    # 調整財報"科目"欄位寬度
    __autofit_column(sheet, "C")

    return temporary_file

def create_argosy_style_financial_statements(sheet, company_code, input_fiscal_year, input_quarter, report_text, category: CategoryEnum):
    for(k, v) in argosy_report_format.items():
        # 合併儲存格
        sheet.merge_cells(v)
        # 劃格線
        __draw_border(sheet, v)
        # 填入文字
        if report_text.get(k):
            __get_first_cell(sheet, v).value = report_text[k]
            __set_cell_height(sheet, __get_first_cell(sheet, v))    
        # 處理標題格式
        if k == "title":
            __set_cell_blod(__get_first_cell(sheet, v))    
            __set_cell_alignment_center(__get_first_cell(sheet, v))

    # 填入獲利分析圖表
    year_array = __get_chart_year(input_fiscal_year)
    buffer = __get_year_operating_chart(company_code, year_array[0], year_array[-1], category)
    img = openpyxl.drawing.image.Image(buffer)
    img.width = 600
    img.height = 450
    sheet.add_image(img, "K4")

    




