from sqlalchemy import BigInteger, Column, Double, Integer, String, DateTime
from database import Base

class FinancialStatementsTw(Base):
    __tablename__ = "financial_statements_tw"
    #__table_args__ = {'schema': 'financial_statements'}  # Specify the schema name
    id = Column(Integer, primary_key=True)
    company_code = Column(String)
    company_name = Column(String)
    fiscal_year = Column(Integer)
    quarter = Column(Integer)
    operating_revenue = Column(Integer)
    operating_costs = Column(Integer)
    gross_profit = Column(Integer)
    gross_profit_margin = Column(Double)
    operating_expenses = Column(Integer)
    operating_income = Column(Integer)
    depreciation = Column(Integer)
    amortization = Column(Integer)
    ebitda = Column(Integer)
    total_nonoperating_income = Column(Integer)
    interest_income = Column(Integer)
    net_investment_income = Column(Integer)
    other_nonoperating_income = Column(Integer)
    total_nonoperating_expenses = Column(Integer)
    interest_expenses = Column(Integer)
    investment_losses = Column(Integer)
    other_nonoperating_expenses = Column(Integer)
    pretax_income = Column(Integer)
    pretax_net_profit_margin = Column(Double)
    income_tax_expense = Column(Integer)
    minority_interest_income = Column(Integer)
    net_income = Column(Integer)
    net_profit_margin = Column(Double)
    create_time = Column(DateTime)

    def as_dict(self):
        return {c.name: getattr(self, c.name) for c in self.__table__.columns}


class FinancialStatementsForeign(Base):
    __tablename__ = "financial_statements_foreign"
    #__table_args__ = {'schema': 'financial_statements'}  # Specify the schema name
    id = Column(Integer, primary_key=True)
    company_code = Column(String)
    company_name = Column(String)
    fiscal_year = Column(Integer)
    quarter = Column(Integer)
    net_sales = Column(Double)
    qoq_growth = Column(Double)
    yoy_growth = Column(Double)
    cost_of_sales = Column(Double)
    gross_profit = Column(Double)
    gross_profit_margin = Column(Double)
    selling_general_administrative_expenses = Column(Double)
    selling_general_administrative_expenses_percentage = Column(Double)
    operating_income = Column(Double)
    pretax_profit_margin = Column(Double)
    dividend_payment = Column(Double)
    other_income_and_expenses = Column(Double)
    total_interest_and_other_expenses = Column(Double)
    pretax_income = Column(Double)
    pretax_net_profit_margin = Column(Double)
    income_tax = Column(Double)
    effective_tax_rate = Column(Double)
    shareholders_equity = Column(Double)
    net_income = Column(Double)
    net_profit_margin = Column(Double)
    create_time = Column(DateTime)

    def as_dict(self):
        return {c.name: getattr(self, c.name) for c in self.__table__.columns}      

class CrawlerNews(Base):
    __tablename__ = "crawler_news"
    id = Column(BigInteger, primary_key=True)
    title = Column(String)
    summary = Column(String)
    link = Column(String)
    publish_on = Column(DateTime)
    stock_code = Column(String)
    cleared_content = Column(String)
    translate_content = Column(String)
   