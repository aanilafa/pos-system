import sqlite3

def init_db():
    conn = sqlite3.connect("inventory.db")
    cursor = conn.cursor()
    
    # Products Table
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS products (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            category TEXT NOT NULL,
            price REAL NOT NULL,
            stock INTEGER NOT NULL
        )
    ''')
    
    # Sales Table
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS sales (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            timestamp DATETIME DEFAULT CURRENT_TIMESTAMP,
            product_name TEXT,
            quantity INTEGER,
            total_price REAL,
            payment_method TEXT,
            cashier TEXT
        )
    ''')
    
    # Cashier Roster Table
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS cashiers (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL UNIQUE
        )
    ''')
    
    conn.commit()
    conn.close()

if __name__ == "__main__":
    init_db()
    print("Database updated with Cashiers table!")