import os
#from dotenv import load_dotenv
from dotenv import dotenv_values

company_name_mapping = {
    "3533": "Lotes",
    "3217": "Argosy",
    "tel": "TEL",
    "aph": "APH"
}

my_setting = {}

#load_dotenv()
#minio_host_env = os.getenv('MINIO_HOST')
#minio_access_key_env = os.getenv('MINIO_ACCESS_KEY')
#minio_secret_key_env = os.getenv('MIMIO_SECRET_KEY')
#bucket_name_env = os.getenv('MINIO_BUCKET') 
env_variables = dotenv_values(".env")

for key, value in env_variables.items():
    #print(f"{key} = {value}")
    my_setting[key] = value
    env_value = os.getenv(key)
    if env_value is not None:
        my_setting[key] = env_value