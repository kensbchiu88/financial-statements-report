from database import SessionLocal, engine
from models import CrawlerNews, FinancialStatementsForeign, FinancialStatementsTw
from sqlalchemy import text

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
    
########
def get_news_by_stock_code_and_publish_on_between(stock_code, publist_on_start, publist_on_end):
    try:
        session = SessionLocal()
        return session.query(CrawlerNews).filter(CrawlerNews.stock_code.like(f'%{stock_code}%')).filter(CrawlerNews.publish_on.between(publist_on_start, publist_on_end)).order_by(CrawlerNews.publish_on).all()
    finally:
        session.close()   