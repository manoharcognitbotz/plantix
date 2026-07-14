import os
import psycopg2
from psycopg2.extensions import ISOLATION_LEVEL_AUTOCOMMIT
from dotenv import load_dotenv

# Absolute path to the project directory for RPA compatibility
try:
    PROJECT_DIR = os.path.dirname(os.path.abspath(__file__))
except NameError:
    PROJECT_DIR = r"C:\UiPathProjects\BlankProcess\SE"
ENV_PATH = os.path.join(PROJECT_DIR, ".env")

def main():
    load_dotenv(ENV_PATH)
    db_host = os.getenv("DB_HOST")
    db_port = os.getenv("DB_PORT")
    db_user = os.getenv("DB_USER")
    db_password = os.getenv("DB_PASSWORD")
    db_name = os.getenv("DB_NAME")

    # -------------------------------------------------------------------------
    # Step 1: Create the database if it does not exist
    # -------------------------------------------------------------------------
    print(f"[*] Connecting to PostgreSQL server {db_host}:{db_port} as '{db_user}'...")
    try:
        conn_params = {
            "host": db_host,
            "port": db_port,
            "user": db_user,
            "password": db_password,
            "database": "postgres"
        }
        if db_host and "neon.tech" in db_host:
            conn_params["sslmode"] = "require"
        conn = psycopg2.connect(**conn_params)
        conn.set_isolation_level(ISOLATION_LEVEL_AUTOCOMMIT)
        cursor = conn.cursor()

        cursor.execute("SELECT 1 FROM pg_catalog.pg_database WHERE datname = %s;", (db_name,))
        if not cursor.fetchone():
            print(f"[*] Database '{db_name}' does not exist. Creating...")
            cursor.execute(f'CREATE DATABASE "{db_name}";')
            print(f"[+] Database '{db_name}' created successfully.")
        else:
            print(f"[*] Database '{db_name}' already exists.")

        cursor.close()
        conn.close()
    except Exception as e:
        print(f"[*] Skipping database creation check (typical for cloud-hosted databases like Neon): {e}")

    # -------------------------------------------------------------------------
    # Step 2: Connect to the target database and create all tables
    # -------------------------------------------------------------------------
    print(f"\n[*] Connecting to database '{db_name}'...")
    conn_params = {
        "host": db_host,
        "port": db_port,
        "user": db_user,
        "password": db_password,
        "database": db_name
    }
    if db_host and "neon.tech" in db_host:
        conn_params["sslmode"] = "require"
    conn = psycopg2.connect(**conn_params)
    cursor = conn.cursor()

    # --- Table 1: comparison_data (baseline items from Excel) ---
    print("[*] Creating table 'comparison_data'...")
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS comparison_data (
            id SERIAL PRIMARY KEY,
            pr_item_code TEXT NOT NULL,
            uom VARCHAR(50),
            indent_quantity NUMERIC,
            imported_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );
    """)

    # --- Table 2: processed_files (tracks which files have been processed) ---
    print("[*] Creating table 'processed_files'...")
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS processed_files (
            id SERIAL PRIMARY KEY,
            filename VARCHAR(255) UNIQUE,
            processed_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );
    """)

    # --- Table 3: raw_quotations (cached Groq LLM parsed JSON per file) ---
    print("[*] Creating table 'raw_quotations'...")
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS raw_quotations (
            filename VARCHAR(255) PRIMARY KEY,
            company_name VARCHAR(255),
            parsed_json TEXT,
            processed_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );
    """)

    # --- Table 4: supplier_slots (maps slot 1/2/3 to supplier names) ---
    print("[*] Creating table 'supplier_slots'...")
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS supplier_slots (
            slot_number INTEGER PRIMARY KEY,
            supplier_name VARCHAR(255) UNIQUE
        );
    """)
    cursor.execute("""
        INSERT INTO supplier_slots (slot_number, supplier_name)
        VALUES (1, NULL), (2, NULL), (3, NULL)
        ON CONFLICT DO NOTHING;
    """)

    # --- Table 5: supplier_items (side-by-side comparison of items across 3 suppliers) ---
    print("[*] Creating table 'supplier_items'...")
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS supplier_items (
            db_item_id INTEGER PRIMARY KEY REFERENCES comparison_data(id),
            item_description TEXT,
            qty NUMERIC,
            uom VARCHAR(50),

            supplier_1_name VARCHAR(255),
            supplier_1_rate NUMERIC,
            supplier_1_discount NUMERIC,
            supplier_1_amount NUMERIC,

            supplier_2_name VARCHAR(255),
            supplier_2_rate NUMERIC,
            supplier_2_discount NUMERIC,
            supplier_2_amount NUMERIC,

            supplier_3_name VARCHAR(255),
            supplier_3_rate NUMERIC,
            supplier_3_discount NUMERIC,
            supplier_3_amount NUMERIC
        );
    """)

    # --- Table 6: supplier_terms (terms & conditions per supplier slot) ---
    print("[*] Creating table 'supplier_terms'...")
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS supplier_terms (
            term_name VARCHAR(100) PRIMARY KEY,
            supplier_1_value TEXT,
            supplier_2_value TEXT,
            supplier_3_value TEXT
        );
    """)
    cursor.execute("""
        INSERT INTO supplier_terms (term_name) VALUES
        ('DISCOUNT'),
        ('P&F CHARGES'),
        ('FREIGHT CHARGES'),
        ('GST'),
        ('PAYMENT TERMS'),
        ('DELIVERY SCHEDULE'),
        ('TEST CERTIFICATES'),
        ('QUOTATION REF/ DATE'),
        ('CONTACT DETAILS')
        ON CONFLICT DO NOTHING;
    """)

    conn.commit()

    # -------------------------------------------------------------------------
    # Step 3: Verify all tables exist
    # -------------------------------------------------------------------------
    print("\n[+] Verifying tables...")
    cursor.execute("""
        SELECT tablename FROM pg_tables
        WHERE schemaname = 'public'
        ORDER BY tablename;
    """)
    tables = cursor.fetchall()
    for t in tables:
        cursor.execute(f'SELECT COUNT(*) FROM "{t[0]}"')
        count = cursor.fetchone()[0]
        print(f"  [OK] {t[0]} ({count} rows)")

    cursor.close()
    conn.close()
    print("\n[+] All tables created successfully. Database is ready.")


if __name__ == "__main__":
    main()
