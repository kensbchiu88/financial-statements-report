import os
from dotenv import dotenv_values
from constant import CategoryEnum

company_name_mapping = {
    "3533": "Lotes",
    "3217": "Argosy",
    "1385157": "TEL",
    "820313": "APH"
}

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
    "other_nonoperating_income": "其他營業外收入 J",
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
    "income_tax": "稅捐 N",
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
    "selling_general_administrative_expenses_percentage": 1,
    "pretax_profit_margin": 1,
    "pretax_net_profit_margin": 1,
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

my_setting = {}

env_variables = dotenv_values(".env")

for key, value in env_variables.items():
    my_setting[key] = value
    env_value = os.getenv(key)
    if env_value is not None:
        my_setting[key] = env_value

def get_financial_statement_items_dict(category: CategoryEnum):
    if category == CategoryEnum.TW:
        return financial_statement_tw_items
    else:
        return financial_statement_foreign_items
    
def get_financial_statement_items_percentage_format_dict(category: CategoryEnum):
    if category == CategoryEnum.TW:
        return financial_statement_tw_items_percentage_format
    else:
        return financial_statement_foreign_items_percentage_format

def get_financial_statement_items_color_format_dict(category: CategoryEnum):
    if category == CategoryEnum.TW:
        return financial_statement_tw_items_color_format
    else:
        return financial_statement_foreign_items_color_format             