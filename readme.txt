1. 透過docker 執行
    - build image: docker compose build
    - build + run: docker compose up -d 
2. 直接使用python執行程式, 
    - (default) 時間參數為上一季: python main.py 
    - 執行特定季度 : python main.py [year] [quarter].
        ex: python main.py 2023 1 