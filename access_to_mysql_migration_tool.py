# Sample Access database path:
# D:\VIT\SY\Sem-3\DBMS\DBMS-CP\Data Migration\Northwind.accdb
import pyodbc
import pandas as pd
from sqlalchemy import create_engine, text, types
import mysql.connector
import os
import warnings
import sys

warnings.filterwarnings('ignore')

# ==========================================
# 1. UTILITY & ANALYSIS FUNCTIONS
# ==========================================

def get_access_tables(db_path):
    conn_str = (
        r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
        fr"DBQ={db_path};"
        r"ExtendedAnsiSQL=1;"
    )
    try:
        cnxn = pyodbc.connect(conn_str)
        cursor = cnxn.cursor()
        tables = [row.table_name for row in cursor.tables(tableType='TABLE') if not row.table_name.startswith('MSys')]
        cnxn.close()
        return tables
    except Exception as e:
        print(f"❌ Error connecting to Access DB: {e}")
        return []

def analyze_compatibility(db_path):
    """Hybrid analyzer to recommend best target DB."""
    print("\n🔍 Running Compatibility Analysis...")
    try:
        tables = get_access_tables(db_path)
        if not tables:
            print("⚠️ No tables found.")
            return {}

        conn_str = (
            r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
            fr"DBQ={db_path};"
            r"ExtendedAnsiSQL=1;"
        )
        cnxn = pyodbc.connect(conn_str)

        dtype_list = []
        print(f"   Scanning {min(len(tables), 5)} tables for data patterns...")
        for t in tables[:5]:
            try:
                query = f"SELECT TOP 10 * FROM [{t}]"
                df = pd.read_sql(query, cnxn)
                dtype_list.extend(list(df.dtypes.astype(str)))
            except:
                pass
        cnxn.close()

        total = len(dtype_list)
        if total == 0:
            print("⚠️ Could not read data types. Defaulting scores.")
            datatype_score = 70
        else:
            numeric = sum('int' in d or 'float' in d for d in dtype_list)
            text_cols = sum('object' in d or 'str' in d for d in dtype_list)
            date_cols = sum('datetime' in d for d in dtype_list)
            datatype_score = ((numeric + text_cols + date_cols) / total * 100)

        complexity = 100 - min(len(tables) * 2, 30)
        
        weights = [0.15, 0.10, 0.10, 0.10] 

        def hybrid_score(base):
            return round(0.30 * datatype_score + 0.20 * complexity +
                         sum(w * b for w, b in zip(weights, base)), 2)

        scores = {
            "MySQL":      hybrid_score([95, 90, 95, 90]),
            "PostgreSQL": hybrid_score([90, 92, 93, 90]),
            "Oracle":     hybrid_score([85, 80, 70, 90]),
            "SQLite":     hybrid_score([80, 85, 85, 90])
        }

        best = max(scores, key=scores.get)
        print("\n📊 Compatibility Report:")
        print(f"   Tables Found: {len(tables)}")
        print("-" * 30)
        for k, v in scores.items():
            print(f"   {k:<12}: {v}%")
        print("-" * 30)
        print(f"✅ Recommended Database: {best}")
        input("\nPress Enter to return to menu...")
        return scores

    except Exception as e:
        print(f"❌ Error during analysis: {e}")
        return {}

# ==========================================
# 2. EXTRACTION & TRANSFORMATION
# ==========================================

def extract_from_access(db_path, table_name):
    conn_str = (
        r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
        fr"DBQ={db_path};"
        r"ExtendedAnsiSQL=1;"
    )
    cnxn = pyodbc.connect(conn_str)
    query = f"SELECT * FROM [{table_name}]"
    
    df = pd.read_sql(query, cnxn)
    cnxn.close()
    print(f"Extracted {len(df)} rows from '{table_name}'")
    return df

def clean_col_name(name):
    return (name.strip()
            .replace(' ', '_')
            .replace('/', '_')
            .replace('-', '_')
            .replace('.', '')
            .lower())

def transform_data(df):
    df = df.copy()
    df.columns = [clean_col_name(col) for col in df.columns]
    
    df = df.where(pd.notnull(df), None)
    
    df.replace("", None, inplace=True)

    for col in df.columns:
        if df[col].dtype == 'bool':
            df[col] = df[col].astype(int)
            
    return df

def map_access_type_to_mysql(access_type, col_name):
    """Translates Access SQL data types to MySQL data types with overrides."""
    access_type = access_type.upper()
    col_name = col_name.lower()
    
    if any(x in col_name for x in ['address', 'note', 'desc', 'attachment', 'source', 'memo', 'image']):
        return "LONGTEXT"
    
    if 'COUNTER' in access_type:
        return "INT AUTO_INCREMENT PRIMARY KEY"
    elif 'BYTE' in access_type:
        return "TINYINT"
    elif 'BIT' in access_type:
        return "TINYINT(1)"
    elif 'INTEGER' in access_type or 'LONG' in access_type or 'SHORT' in access_type:
        return "INT"
    elif 'CURRENCY' in access_type or 'MONEY' in access_type or 'DECIMAL' in access_type:
        return "DECIMAL(19,4)"
    elif 'DATETIME' in access_type:
        return "DATETIME"
    elif 'LONGCHAR' in access_type or 'MEMO' in access_type:
        return "LONGTEXT"
    elif 'BINARY' in access_type:
        return "LONGBLOB"
    else:
        return "VARCHAR(255)"

def get_table_schema_safe(db_path, table_name):
    conn_str = (
        r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
        fr"DBQ={db_path};"
        r"ExtendedAnsiSQL=1;"
    )
    cnxn = pyodbc.connect(conn_str)
    cursor = cnxn.cursor()

    schema = []
    try:
        for col in cursor.columns(table=table_name):
            cleaned_name = clean_col_name(col.column_name)
            mysql_type = map_access_type_to_mysql(col.type_name, cleaned_name)
            
            schema.append({
                'name': cleaned_name,
                'type': mysql_type,
                'nullable': col.is_nullable,
                'is_pk': 'COUNTER' in col.type_name.upper()
            })
    except Exception as e:
        print(f"Warning: Could not fetch schema via ODBC ({e}). Will rely on Pandas.")
        return None
    finally:
        cnxn.close()

    return schema

# ==========================================
# 3. LOADING (MYSQL OPERATIONS)
# ==========================================

def create_mysql_table(schema, table_name, mysql_conn):
    try:
        cursor = mysql_conn.cursor()
        cursor.execute(f"DROP TABLE IF EXISTS `{table_name}`;")

        cols_def = []
        for col in schema:
            col_def = f"`{col['name']}` {col['type']}"
            if col['nullable'] == "NO" and "PRIMARY KEY" not in col['type']:
                col_def += " NOT NULL"
            cols_def.append(col_def)

        if not cols_def:
            return

        create_sql = f"CREATE TABLE `{table_name}` ({', '.join(cols_def)}) ENGINE=InnoDB;"
        cursor.execute(create_sql)
        mysql_conn.commit()
        print(f"✅ Table '{table_name}' schema created.")
        
    except Exception as e:
        print(f"❌ Error creating table '{table_name}': {e}")

def verify_row_count(engine, table_name):
    try:
        with engine.connect() as conn:
            result = conn.execute(text(f"SELECT COUNT(*) FROM `{table_name}`"))
            count = result.scalar()
            if count > 0:
                print(f"🧐 VERIFICATION SUCCESS: '{table_name}' has {count} rows.")
            else:
                print(f"⚠️ VERIFICATION FAILED: '{table_name}' is empty.")
    except Exception as e:
        print(f"⚠️ Verification failed: {e}")

def load_to_mysql(df, config, table_name, db_path):
    if df is None or df.empty: 
        print("⚠️ DataFrame is empty, skipping.")
        return

    try:
        table_safe = clean_col_name(table_name)
        
        mysql_conn = mysql.connector.connect(
            host=config['host'], user=config['user'], 
            password=config['password'], database=config['db']
        )
        
        cursor = mysql_conn.cursor()
        cursor.execute("SET SESSION sql_mode = '';") 
        
        schema = get_table_schema_safe(db_path, table_name)
        if schema:
            create_mysql_table(schema, table_safe, mysql_conn)
        else:
            print(f"⚠️ Using Pandas Fallback for '{table_name}' structure...")
        
        mysql_conn.close()

        # 2. Insert Data
        print(f"⏳ Inserting {len(df)} rows...")
        
        engine_url = f"mysql+mysqlconnector://{config['user']}:{config['password']}@{config['host']}/{config['db']}"
        engine = create_engine(engine_url)
        
        dtype_map = {}
        for col in df.columns:
            if any(x in col for x in ['address', 'note', 'attachment', 'desc']):
                dtype_map[col] = types.Text()

        with engine.connect() as conn:
            conn.execute(text("SET SESSION sql_mode = '';"))
            
            df.to_sql(table_safe, con=conn, if_exists='append', index=False, chunksize=1000, dtype=dtype_map)
            
            conn.commit()
        # --------------------------------
        
        print(f"✅ Data committed to '{table_safe}'.")
        
        # 3. Verify
        verify_row_count(engine, table_safe)
        
    except Exception as e:
        print(f"❌ Error migrating '{table_name}': {e}")

# ==========================================
# 4. MAIN MENU
# ==========================================

def perform_migration(db_path):
    print("\n--- MYSQL CREDENTIALS ---")
    config = {
        'host': input("MySQL Host (default localhost): ") or 'localhost',
        'user': input("MySQL User (default root): ") or 'root',
        'password': input("MySQL Password: "),
        'db': input("Target Database Name: ")
    }

    print("\n🚀 STARTING MIGRATION...")
    tables = get_access_tables(db_path)
    
    for t in tables:
        print(f"\nProcessing: {t}")
        try:
            df = extract_from_access(db_path, t)
            tdf = transform_data(df)
            load_to_mysql(tdf, config, t, db_path)
        except Exception as e:
            print(f"CRITICAL FAILURE ON TABLE {t}: {e}")
        print("-" * 40)
    print("\n✅ Migration Process Finished.")
    input("Press Enter to return to menu...")

def main():
    print("\n==========================================")
    print("   ACCESS TO MYSQL MIGRATION TOOL v5.0")
    print("==========================================")
    
    while True:
        db_path = input("\nEnter Access DB path (.accdb/.mdb): ").strip().replace('"', '')
        if os.path.exists(db_path):
            break
        print("❌ File not found.")

    while True:
        print(f"\n📂 ACTIVE DATABASE: {os.path.basename(db_path)}")
        print("1. 📊 Analyze Database Compatibility")
        print("2. 📋 List All Tables")
        print("3. 🚀 Start Migration to MySQL")
        print("4. 🔄 Change Access Database File")
        print("5. ❌ Exit")
        
        choice = input("Enter choice (1-5): ").strip()
        
        if choice == '1':
            analyze_compatibility(db_path)
            
        elif choice == '2':
            tables = get_access_tables(db_path)
            print(f"\n📋 Found {len(tables)} Tables:")
            for idx, t in enumerate(tables, 1):
                print(f"  {idx}. {t}")
            input("\nPress Enter to return to menu...")
            
        elif choice == '3':
            perform_migration(db_path)
            
        elif choice == '4':
            new_path = input("Enter NEW Access DB path: ").strip().replace('"', '')
            if os.path.exists(new_path):
                db_path = new_path
                print("✅ Database updated.")
            else:
                print("❌ File not found. Keeping previous database.")
                
        elif choice == '5':
            print("Exiting... Goodbye!")
            break
            
        else:
            print("❌ Invalid choice. Please enter 1-5.")

if __name__ == "__main__":
    main()
