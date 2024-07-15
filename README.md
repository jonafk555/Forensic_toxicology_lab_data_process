# Forensic_toxicology_lab_data_process

## 第一次使用
- 環境建置：
1. 安裝 python ：https://www.python.org/downloads/
2. 搜尋中找到 `cmd`：![image](https://github.com/user-attachments/assets/c9dc4203-2040-45d8-9927-cf3575eb1226)
- 輸入：
  - python -m pip install --upgrade pip
  - pip install pandas openpyxl PyMuPDF PyQt5
  
- 問題詳見：https://ithelp.ithome.com.tw/m/articles/10216493 

## Data Extraction
- 把 infusion 檔案匯出成副檔名為".pdf"的檔案，須確保"compound"名稱中沒有空白。
- 匯出後執行 `data_extraction.py` 匯入檔案，並儲存。
- 找到儲存檔案即可。

## Database Process
- 執行 `database_process.py`，匯入檔案。
- 以 `Global search` 尋找想要的資料。
