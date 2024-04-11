import os
from minio import Minio
from minio.error import S3Error
from datetime import datetime
from dotenv import load_dotenv
import requests
#from dotenv import dotenv_values
from setting import my_setting

#load_dotenv()
#minio_host_env = os.getenv('MINIO_HOST')
#minio_access_key_env = os.getenv('MINIO_ACCESS_KEY')
#minio_secret_key_env = os.getenv('MIMIO_SECRET_KEY')
#bucket_name_env = os.getenv('MINIO_BUCKET') 
#env_variables = dotenv_values(".env")

#for key, value in env_variables.items():
#    print(f"{key} = {value}")

def upload_to_minio(local_file_path):
    today_date = datetime.now().strftime("%Y-%m-%d")
    directory, file_name = os.path.split(local_file_path)
    file_name, file_extension = os.path.splitext(file_name)
    file_name = file_name.replace("_", "-")

    #minio_host =os.environ.get('MINIO_HOST', minio_host_env)
    #minio_access_key =os.environ.get('MINIO_ACCESS_KEY', minio_access_key_env)
    #minio_secret_key =os.environ.get('MINIO_SECRET_KEY', minio_secret_key_env)
    #bucket_name =os.environ.get('MINIO_BUCKET', bucket_name_env)

    # 初始化 MinIO 客户端
    minio_client = Minio(
        my_setting['MINIO_HOST'],
        access_key=my_setting['MINIO_ACCESS_KEY'],
        secret_key=my_setting['MINIO_SECRET_KEY'],
        secure=False  # 如果是 HTTPS 连接，请设置为 True
    )

    # 指定要上传的本地文件和在 MinIO 上的存储桶名称以及对象名称
    object_name = f"{file_name}-{today_date}{file_extension}"

    # 上传文件到 MinIO
    try:
        minio_client.fput_object(
            my_setting['MINIO_BUCKET'],
            object_name,
            local_file_path,
        )
        print(f"File uploaded successfully to MinIO: {object_name}")
        return object_name
    except S3Error as e:
        print(f"Error uploading file to MinIO: {e}")
        return None


def get_previous_quarter(year, quarter):
    """
    取得上一季的年份和季度
    
    Parameters:
    year (int): The year for which the previous quarter needs to be calculated.
    quarter (int): The quarter for which the previous quarter needs to be calculated.
    
    Returns:
    tuple: A tuple containing the year and quarter of the previous quarter.
    """    
    return_year = year
    return_quarter = quarter

    if quarter - 1 == 0:
        return_year = year - 1
        return_quarter = 4
    else:
        return_quarter = quarter - 1
    return return_year, return_quarter   

def call_send_email_api(json_data): 
    url = my_setting['SEND_EMAIL_API_URL']
    response = requests.post(url, data=json_data, headers={'Content-Type': 'application/json'})
    if response.status_code == 200:
        print('email sent:', response.status_code)
    else:
        print('Failed to send email:', response.status_code)
