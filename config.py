import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, "data", "fsin_med.db")

# Создаём папку data, если её нет
os.makedirs(os.path.join(BASE_DIR, "data"), exist_ok=True)

# Названия таблиц
TABLE_MED_PARTS = "med_parts"
TABLE_MEDICATIONS = "medications"
TABLE_STOCK = "stock"
TABLE_MOVEMENTS = "movements"
TABLE_WRITEOFFS = "writeoffs"
TABLE_PRISONERS = "prisoners"
TABLE_RECORDS = "medical_records"