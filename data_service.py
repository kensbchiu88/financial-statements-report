from constant import CategoryEnum
from database import SessionLocal, engine
from models import CrawlerNews, FinancialStatementsForeign, FinancialStatementsTw
from sqlalchemy import text
import setting
import util

# 紀錄國外財報每季的營業收入。For計算季報中的QoQ / YoY
quarter_net_sales = {}

def get_financial_statements_by_quarter(company_code, fiscal_year, quarter):
    try:
        session = SessionLocal()
        return session.query(FinancialStatementsTw).filter(FinancialStatementsTw.company_code == company_code).filter(FinancialStatementsTw.fiscal_year == fiscal_year).filter(FinancialStatementsTw.quarter == quarter).first()
    finally:
        session.close()

def get_financial_statements_by_year(company_code, fiscal_year):
    try:
        session = SessionLocal()
        return session.query(FinancialStatementsTw).filter(FinancialStatementsTw.company_code == company_code).filter(FinancialStatementsTw.fiscal_year == fiscal_year).all()
    finally:
        session.close()

def get_financial_statements_by_year_between(company_code, fiscal_year_start, fiscal_year_end):    
    try:
        session = SessionLocal()
        return session.query(FinancialStatementsTw).filter(FinancialStatementsTw.company_code == company_code).filter(FinancialStatementsTw.fiscal_year.between(fiscal_year_start, fiscal_year_end)).order_by(FinancialStatementsTw.fiscal_year, FinancialStatementsTw.quarter).all()
    finally:
        session.close()

def get_year_operating_revenue(company_code, start_fiscal_year, end_fiscal_year):
    with engine.connect() as conn:
        sql_statement = text("select sum(operating_revenue) from financial_statements_tw where company_code = :company_code and fiscal_year between :start_fiscal_year and :end_fiscal_year group by fiscal_year order by fiscal_year ")
        result = conn.execute(sql_statement, {'company_code': company_code, 'start_fiscal_year': start_fiscal_year, 'end_fiscal_year': end_fiscal_year}).fetchall()
        return result    

def get_year_gross_profit(company_code, start_fiscal_year, end_fiscal_year):
    with engine.connect() as conn:
        sql_statement = text("select sum(gross_profit) from financial_statements_tw where company_code = :company_code and fiscal_year between :start_fiscal_year and :end_fiscal_year group by fiscal_year order by fiscal_year ")
        result = conn.execute(sql_statement, {'company_code': company_code, 'start_fiscal_year': start_fiscal_year, 'end_fiscal_year': end_fiscal_year}).fetchall()
        return result    
    
def get_year_gross_profit_margin(company_code, start_fiscal_year, end_fiscal_year):
    with engine.connect() as conn:
        sql_statement = text("select ROUND((CAST(sum(gross_profit) as float) / sum(operating_revenue))::numeric, 3) from financial_statements_tw where company_code = :company_code and fiscal_year between :start_fiscal_year and :end_fiscal_year group by fiscal_year order by fiscal_year ")
        result = conn.execute(sql_statement, {'company_code': company_code, 'start_fiscal_year': start_fiscal_year, 'end_fiscal_year': end_fiscal_year}).fetchall()
        return result        

def get_year_net_income(company_code, start_fiscal_year, end_fiscal_year):
    with engine.connect() as conn:
        sql_statement = text("select sum(pretax_income) from financial_statements_tw where company_code = :company_code and fiscal_year between :start_fiscal_year and :end_fiscal_year group by fiscal_year order by fiscal_year ")
        result = conn.execute(sql_statement, {'company_code': company_code, 'start_fiscal_year': start_fiscal_year, 'end_fiscal_year': end_fiscal_year}).fetchall()
        return result    

def get_year_net_profit_margin(company_code, start_fiscal_year, end_fiscal_year):
    with engine.connect() as conn:
        sql_statement = text("select ROUND((CAST(sum(pretax_income) as float) / sum(operating_revenue))::numeric, 3) from financial_statements_tw where company_code = :company_code and fiscal_year between :start_fiscal_year and :end_fiscal_year group by fiscal_year order by fiscal_year ")
        result = conn.execute(sql_statement, {'company_code': company_code, 'start_fiscal_year': start_fiscal_year, 'end_fiscal_year': end_fiscal_year}).fetchall()
        return result    

################################################
def get_financial_statements_foreign_by_quarter(company_code, fiscal_year, quarter):
    try:
        session = SessionLocal()
        return session.query(FinancialStatementsForeign).filter(FinancialStatementsForeign.company_code == company_code).filter(FinancialStatementsForeign.fiscal_year == fiscal_year).filter(FinancialStatementsForeign.quarter == quarter).first()
    finally:
        session.close()

def get_financial_statements_foreign_by_year(company_code, fiscal_year):
    try:
        session = SessionLocal()
        return session.query(FinancialStatementsForeign).filter(FinancialStatementsForeign.company_code == company_code).filter(FinancialStatementsForeign.fiscal_year == fiscal_year).all()
    finally:
        session.close()    

def get_financial_statements_foreign_by_year_between(company_code, fiscal_year_start, fiscal_year_end):    
    try:
        session = SessionLocal()
        return session.query(FinancialStatementsForeign).filter(FinancialStatementsForeign.company_code == company_code).filter(FinancialStatementsForeign.fiscal_year.between(fiscal_year_start, fiscal_year_end)).order_by(FinancialStatementsForeign.fiscal_year, FinancialStatementsForeign.quarter).all()
    finally:
        session.close()            

#年營收
def get_year_net_sales_foreign(company_code, start_fiscal_year, end_fiscal_year):
    with engine.connect() as conn:
        sql_statement = text("select sum(net_sales) from financial_statements_foreign where company_code = :company_code and fiscal_year between :start_fiscal_year and :end_fiscal_year group by fiscal_year order by fiscal_year ")
        result = conn.execute(sql_statement, {'company_code': company_code, 'start_fiscal_year': start_fiscal_year, 'end_fiscal_year': end_fiscal_year}).fetchall()
        return result    

#年毛利
def get_year_gross_profit_foreign(company_code, start_fiscal_year, end_fiscal_year):
    with engine.connect() as conn:
        sql_statement = text("select sum(gross_profit) from financial_statements_foreign where company_code = :company_code and fiscal_year between :start_fiscal_year and :end_fiscal_year group by fiscal_year order by fiscal_year ")
        result = conn.execute(sql_statement, {'company_code': company_code, 'start_fiscal_year': start_fiscal_year, 'end_fiscal_year': end_fiscal_year}).fetchall()
        return result    

#年毛利率    
def get_year_gross_profit_margin_foreign(company_code, start_fiscal_year, end_fiscal_year):
    with engine.connect() as conn:
        sql_statement = text("select ROUND((CAST(sum(gross_profit) as float) / sum(net_sales))::numeric, 3) from financial_statements_foreign where company_code = :company_code and fiscal_year between :start_fiscal_year and :end_fiscal_year group by fiscal_year order by fiscal_year ")
        result = conn.execute(sql_statement, {'company_code': company_code, 'start_fiscal_year': start_fiscal_year, 'end_fiscal_year': end_fiscal_year}).fetchall()
        return result        

#年淨利
def get_year_net_income_foreign(company_code, start_fiscal_year, end_fiscal_year):
    with engine.connect() as conn:
        sql_statement = text("select sum(net_income) from financial_statements_foreign where company_code = :company_code and fiscal_year between :start_fiscal_year and :end_fiscal_year group by fiscal_year order by fiscal_year ")
        result = conn.execute(sql_statement, {'company_code': company_code, 'start_fiscal_year': start_fiscal_year, 'end_fiscal_year': end_fiscal_year}).fetchall()
        return result    

#年淨利率
def get_year_net_profit_margin_foreign(company_code, start_fiscal_year, end_fiscal_year):
    with engine.connect() as conn:
        sql_statement = text("select ROUND((CAST(sum(net_income) as float) / sum(net_sales))::numeric, 3) from financial_statements_foreign where company_code = :company_code and fiscal_year between :start_fiscal_year and :end_fiscal_year group by fiscal_year order by fiscal_year ")
        result = conn.execute(sql_statement, {'company_code': company_code, 'start_fiscal_year': start_fiscal_year, 'end_fiscal_year': end_fiscal_year}).fetchall()
        return result    
    
# 抓取新聞資料
def get_news_by_stock_code_and_publish_on_between(stock_code, publist_on_start, publist_on_end):
    try:
        session = SessionLocal()
        return session.query(CrawlerNews).filter(CrawlerNews.stock_code.like(f'%{stock_code}%')).filter(CrawlerNews.publish_on.between(publist_on_start, publist_on_end)).order_by(CrawlerNews.publish_on).all()
    finally:
        session.close()   

# 計算年度報表數字
def get_financial_statements_year_values(company_code, fiscal_year, category: CategoryEnum) -> dict:
    #result = {}
    #data_dict_list = []
    #financial_statement_items = __get_financial_statement_items_dict(category)

    if category == CategoryEnum.TW:
        data_list = get_financial_statements_by_year(company_code, fiscal_year)
    else:
        data_list = get_financial_statements_foreign_by_year(company_code, fiscal_year)

    # 計算年度報表中可加總的欄位
    result = __get_sum_of_numeric_field_of_financial_statements(data_list, category)

    # 計算年度報表中的不可加總部分，如毛利率等
    result = __get_simple_calculated_value_of_financial_statements(result, category)

    # 計算國外年報中的YOY
    if category == CategoryEnum.FOREIGN:
        # 計算YoY
        previous_model_list = get_financial_statements_foreign_by_year(company_code, fiscal_year - 1)
        previous_sum_dict =  __get_sum_of_numeric_field_of_financial_statements(previous_model_list, category)
        if previous_sum_dict['net_sales'] > 0 and result['net_sales'] > 0:
            result['yoy_growth'] = (result['net_sales'] - previous_sum_dict['net_sales']) / previous_sum_dict['net_sales']

    return result        

# 取得財報中季度的數值
def get_financial_statements_quarter_values(comany_code, fiscal_year, quarter, category: CategoryEnum):
    result = None
    if category == CategoryEnum.TW:    
        data = get_financial_statements_by_quarter(comany_code, fiscal_year, quarter)
    else:
        data = get_financial_statements_foreign_by_quarter(comany_code, fiscal_year, quarter)

    if data is not None:
        result = data.as_dict()

        # 計算報表中的不可加總部分，如毛利率等
        result = __get_simple_calculated_value_of_financial_statements(result, category)        

        # 計算YoY 國外報表才有QoQ YoY
        if category == CategoryEnum.FOREIGN:
            # 佔存 net_sales。 for 計算YoY QoQ
            quarter_net_sales[f"{comany_code}-{fiscal_year}Q{quarter}"] = result["net_sales"]
            
            key = f"{comany_code}-{fiscal_year-1}Q{quarter}"
            if key in quarter_net_sales:
                previous = quarter_net_sales[key]
            else:
                data = get_financial_statements_foreign_by_quarter(comany_code, fiscal_year - 1, quarter)
                if data is not None:
                    previous_dict = data.as_dict()    
                    previous = previous_dict["net_sales"]
                    # 佔存 net_sales。 for 計算YoY QoQ
                    quarter_net_sales[key] = previous

            if previous > 0:
                result["yoy_growth"] = (result["net_sales"] - previous) / previous 

            # 計算QoQ
            target_year, target_quarter = util.get_previous_quarter(fiscal_year, quarter)
            key = f"{comany_code}-{target_year}Q{target_quarter}"
            if key in quarter_net_sales:
                previous = quarter_net_sales[key]
            else:
                data = get_financial_statements_foreign_by_quarter(comany_code, target_year, target_quarter)
                if data is not None:
                    previous_dict = data.as_dict()    
                    previous = previous_dict["net_sales"]
                    # 佔存 net_sales。 for 計算YoY QoQ
                    quarter_net_sales[key] = previous

            if previous > 0:                
                result["qoq_growth"] = (result["net_sales"] - previous) / previous  
    
    return result

# 計算年度報表中可加總的欄位
def __get_sum_of_numeric_field_of_financial_statements(data_list, category:CategoryEnum):
    result = {}
    data_dict_list = []
    financial_statement_items = setting.get_financial_statement_items_dict(category)
    financial_statement_items_percentage_format = setting.get_financial_statement_items_percentage_format_dict(category)

    for data in data_list:
        if data is not None:
            data_dict_list.append(data.as_dict())

    for k in financial_statement_items.keys():
        sum = 0
        for data_dict in data_dict_list:
            if k in data_dict and k not in financial_statement_items_percentage_format:   # 計算欄位無法單純相加，故必須排除
                if util.is_numeric(data_dict[k]):
                    sum += data_dict[k]
        result[k] = sum    
    return result

# 計算報表中的計算欄位值(不可加總部分)，如毛利率等
def __get_simple_calculated_value_of_financial_statements(data:dict, category:CategoryEnum):
    result = data.copy()
    if category == CategoryEnum.TW:
        #稅後淨利 S=O-Q-R
        result['net_income'] = result['pretax_income'] - result['income_tax_expense'] - result['minority_interest_income']
        
        if data['operating_revenue'] > 0:
            result['gross_profit_margin'] = result['gross_profit'] / result['operating_revenue']
            result['pretax_net_profit_margin'] = result['pretax_income'] / result['operating_revenue']
            result['net_profit_margin'] = result['net_income']  / result['operating_revenue']
    else:
        if data['net_sales'] > 0: 
            result['gross_profit_margin'] = result['gross_profit'] / result['net_sales']
            result['selling_general_administrative_expenses_percentage'] = result['selling_general_administrative_expenses'] / data['net_sales']
            result['pretax_profit_margin'] = result['operating_income'] / result['net_sales']
            result['pretax_net_profit_margin'] = result['pretax_income'] / result['net_sales']
            result['net_profit_margin'] = result['net_income']  / result['net_sales']
            result['effective_tax_rate'] = result['income_tax'] / result['pretax_income']

    return result