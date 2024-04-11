from sqlalchemy import create_engine
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker
from setting import my_setting

db_host = my_setting['DB_HOST']
db_port = my_setting['DB_PORT']
db_name = my_setting['DB_NAME']
db_user = my_setting['DB_USER']
db_password = my_setting['DB_PASSWORD']

DATABASE_URL = f'postgresql://{db_user}:{db_password}@{db_host}:{db_port}/{db_name}'
engine = create_engine(DATABASE_URL)

#engine = create_engine("postgresql://postgres:admin@127.0.0.1:5432/postgres")

SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)

Base = declarative_base()

def get_db() -> any:    
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()
