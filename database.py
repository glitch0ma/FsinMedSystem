import sqlite3
import config

class Database:
    def __init__(self):
        self.db_path = config.DB_PATH
        self.create_tables()

    def get_connection(self):
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        return conn

    def create_tables(self):
        conn = self.get_connection()
        cursor = conn.cursor()

        # Медицинские части
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS med_parts (
                id INTEGER PRIMARY KEY,
                name TEXT NOT NULL,
                address TEXT NOT NULL,
                phone TEXT
            )
        ''')

        # Медикаменты
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS medications (
                id INTEGER PRIMARY KEY,
                name TEXT NOT NULL,
                form TEXT,
                dosage TEXT,
                unit TEXT NOT NULL DEFAULT 'шт',
                description TEXT
            )
        ''')

        # Остатки
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS stock (
                id INTEGER PRIMARY KEY,
                med_part_id INTEGER,
                medication_id INTEGER,
                quantity REAL NOT NULL DEFAULT 0,
                production_date TEXT,
                expiration_date TEXT,
                batch_number TEXT,
                FOREIGN KEY (med_part_id) REFERENCES med_parts(id),
                FOREIGN KEY (medication_id) REFERENCES medications(id)
            )
        ''')

        # Расходование
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS expenditures (
                id INTEGER PRIMARY KEY,
                stock_id INTEGER,
                prisoner_id INTEGER,
                quantity REAL NOT NULL,
                reason TEXT,
                comment TEXT,
                date DATETIME DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (stock_id) REFERENCES stock(id),
                FOREIGN KEY (prisoner_id) REFERENCES prisoners(id)
            )
        ''')

        # Перемещения
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS movements (
                id INTEGER PRIMARY KEY,
                from_part_id INTEGER,
                to_part_id INTEGER,
                medication_id INTEGER,
                quantity REAL NOT NULL,
                date DATETIME DEFAULT CURRENT_TIMESTAMP,
                comment TEXT,
                FOREIGN KEY (from_part_id) REFERENCES med_parts(id),
                FOREIGN KEY (to_part_id) REFERENCES med_parts(id),
                FOREIGN KEY (medication_id) REFERENCES medications(id)
            )
        ''')

        # Заключённые
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS prisoners (
                id INTEGER PRIMARY KEY,
                prisoner_id TEXT UNIQUE NOT NULL,
                full_name TEXT NOT NULL,
                birth_date TEXT,
                arrival_date TEXT,
                notes TEXT
            )
        ''')

        # Медицинские обращения (создаём только если таблицы нет)
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='medical_records'")
        if not cursor.fetchone():
            # Создаём таблицу с новой структурой
            cursor.execute('''
                CREATE TABLE medical_records (
                                                 id             INTEGER PRIMARY KEY,
                                                 prisoner_id    INTEGER,
                                                 service_date   DATETIME DEFAULT CURRENT_TIMESTAMP,
                                                 diagnosis      TEXT,
                                                 doctor_name    TEXT,
                                                 service_name   TEXT NOT NULL,
                                                 quantity       REAL DEFAULT 1,
                                                 unit_price     REAL DEFAULT 0,
                                                 total_amount   REAL DEFAULT 0,
                                                 service_number TEXT,
                                                 notes          TEXT,
                                                 user           TEXT,
                                                 FOREIGN KEY (prisoner_id) REFERENCES prisoners(id)
                )
            ''')
            print("✅ Создана новая таблица medical_records")
        else:
            # Миграция — добавляем новые поля, если их нет
            try:
                cursor.execute("ALTER TABLE medical_records ADD COLUMN diagnosis TEXT")
                cursor.execute("ALTER TABLE medical_records ADD COLUMN doctor_name TEXT")
                cursor.execute("ALTER TABLE medical_records ADD COLUMN quantity REAL DEFAULT 1")
                cursor.execute("ALTER TABLE medical_records ADD COLUMN unit_price REAL DEFAULT 0")
                cursor.execute("ALTER TABLE medical_records ADD COLUMN total_amount REAL DEFAULT 0")
                cursor.execute("ALTER TABLE medical_records ADD COLUMN service_number TEXT")
                print("✅ Добавлены новые поля в таблицу medical_records")
            except sqlite3.OperationalError:
                pass  # поля уже существуют

        print("✅ Все таблицы успешно созданы/проверены")
        conn.commit()
        conn.close()

    def add_test_data(self):
        conn = self.get_connection()
        cursor = conn.cursor()

        print("🔄 Проверка и добавление медчастей...")

        # Принудительно добавляем все 5 медчастей (если какой-то нет — добавляем)
        test_parts = [
            ("Медчасть №1", "г. Улан-Удэ, ул. Ленина 15", "+7 (495) 123-45-67"),
            ("Медчасть №2", "г. Москва, ул. Центральная 8", "+7 (495) 765-43-21"),
            ("Медчасть №3", "г. Москва, ул. Советская 45", "+7 (495) 987-65-43"),
            ("Медчасть №4", "г. Москва, ул. Лесная 12", "+7 (495) 111-22-33"),
            ("Медчасть №5", "г. Москва, ул. Парковая 7", "+7 (495) 444-55-66"),
        ]

        added_count = 0
        for name, address, phone in test_parts:
            cursor.execute("SELECT id FROM med_parts WHERE name = ?", (name,))
            if cursor.fetchone() is None:   # если такой медчасти нет
                cursor.execute("INSERT INTO med_parts (name, address, phone) VALUES (?,?,?)",
                              (name, address, phone))
                added_count += 1
                print(f"   → Добавлена: {name}")

        if added_count > 0:
            print(f"✅ Успешно добавлено {added_count} новых медчастей")
        else:
            print("ℹ️ Все 5 медчастей уже существуют в базе")

        # Добавляем тестовых заключённых
        cursor.execute("SELECT COUNT(*) FROM prisoners")
        if cursor.fetchone()[0] == 0:
            test_prisoners = [
                ("П-12345", "Иванов Иван Иванович", "1995-03-15", "2024-01-10", ""),
                ("П-54321", "Сидоров Пётр Александрович", "1998-11-20", "2023-08-05", "Аллергия на пенициллин"),
                ("П-67890", "Петров Сергей Викторович", "1997-07-12", "2025-01-15", ""),
            ]
            cursor.executemany(
                "INSERT INTO prisoners (prisoner_id, full_name, birth_date, arrival_date, notes) VALUES (?,?,?,?,?)",
                test_prisoners
            )
            print("✅ Добавлены тестовые заключённые")

        conn.commit()
        conn.close()