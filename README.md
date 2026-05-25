# Data Migration From Microsoft Access To MySQL

This project is a command-line migration tool that reads tables from a Microsoft Access database file (`.accdb` or `.mdb`), analyzes the database, and migrates the data into a MySQL database.

The tool was built for migrating the Northwind sample Access database, but it can be used with any valid Access database file if the required drivers and Python packages are installed.

## Features

- Connects to Microsoft Access database files using ODBC.
- Lists all user-created tables in the Access database.
- Performs compatibility analysis and recommends a target database.
- Extracts table data from Access into Pandas DataFrames.
- Cleans column names for MySQL compatibility.
- Converts common Access data types into MySQL data types.
- Creates matching MySQL tables automatically.
- Inserts table data into MySQL.
- Verifies migrated row counts after insertion.
- Supports changing the Access database path from the menu.

## Project File

The main Python script is:

```text
access_to_mysql_migration_tool.py
```

## Requirements

### Software

- Windows operating system
- Python 3.x
- MySQL Server
- Microsoft Access ODBC Driver

### Python Packages

Install the required Python packages:

```bash
pip install pyodbc pandas sqlalchemy mysql-connector-python
```

### Microsoft Access ODBC Driver

The script uses this ODBC driver:

```text
Microsoft Access Driver (*.mdb, *.accdb)
```

If the driver is missing, install the Microsoft Access Database Engine Redistributable.

## Database Setup

Before starting migration, create a target database in MySQL.

Example:

```sql
CREATE DATABASE northwind_migrated;
```

The script creates tables inside this database and loads the Access data into them.

## How To Run

Open a terminal in the project folder and run:

```bash
python access_to_mysql_migration_tool.py
```

When the program asks for the Access database path, enter the full path to your `.accdb` or `.mdb` file.

Example:

```text
D:\VIT\SY\Sem-3\DBMS\DBMS-CP\Data Migration\Northwind.accdb
```

Make sure the file actually exists at the entered path. If the path is wrong, the program will show:

```text
File not found.
```

## Menu Options

After entering a valid Access database path, the program displays this menu:

```text
1. Analyze Database Compatibility
2. List All Tables
3. Start Migration to MySQL
4. Change Access Database File
5. Exit
```

### 1. Analyze Database Compatibility

This option scans the Access database and calculates compatibility scores for different database systems:

- MySQL
- PostgreSQL
- Oracle
- SQLite

It then recommends the best target database based on data types and database complexity.

### 2. List All Tables

This option displays all user tables found in the Access database.

For the Northwind database, sample tables include:

- Categories
- Customers
- Employees
- Orders
- Order Details
- Products
- Shippers
- Suppliers

### 3. Start Migration to MySQL

This option starts the actual migration process.

The program asks for:

- MySQL host
- MySQL username
- MySQL password
- Target database name

Example:

```text
MySQL Host (default localhost): localhost
MySQL User (default root): root
MySQL Password: your_password
Target Database Name: northwind_migrated
```

For each Access table, the script:

1. Extracts rows from Access.
2. Cleans column names.
3. Creates a matching MySQL table.
4. Inserts all rows into MySQL.
5. Verifies the row count.

### 4. Change Access Database File

This option lets you switch to another `.accdb` or `.mdb` file without restarting the program.

### 5. Exit

This option closes the tool.

## Sample Run

```text
==========================================
   ACCESS TO MYSQL MIGRATION TOOL v5.0
==========================================

Enter Access DB path (.accdb/.mdb): D:\VIT\SY\Sem-3\DBMS\DBMS-CP\Data Migration\Northwind.accdb

ACTIVE DATABASE: Northwind.accdb
1. Analyze Database Compatibility
2. List All Tables
3. Start Migration to MySQL
4. Change Access Database File
5. Exit
Enter choice (1-5): 2

Found 13 Tables:
  1. Categories
  2. Customers
  3. Employees
  4. Order Details
  5. Orders
  6. Products
  7. Shippers
  8. Suppliers
```

## Sample Migration Output

```text
STARTING MIGRATION...

Processing: Categories
Extracted 8 rows from 'Categories'
Table 'categories' schema created.
Inserting 8 rows...
Data committed to 'categories'.
VERIFICATION SUCCESS: 'categories' has 8 rows.

Processing: Customers
Extracted 91 rows from 'Customers'
Table 'customers' schema created.
Inserting 91 rows...
Data committed to 'customers'.
VERIFICATION SUCCESS: 'customers' has 91 rows.

Migration Process Finished.
```

## Common Errors

### File not found

This happens when the Access database path is incorrect.

Check that the full path includes every folder name correctly.

Correct example:

```text
D:\VIT\SY\Sem-3\DBMS\DBMS-CP\Data Migration\Northwind.accdb
```

Incorrect example:

```text
D:\VIT\SY\DBMS\DBMS-CP\Data Migration\Northwind.accdb
```

The incorrect path is missing the `Sem-3` folder.

### Access ODBC driver error

This happens when the Microsoft Access ODBC Driver is not installed or does not match your Python architecture.

Install the Microsoft Access Database Engine Redistributable and make sure the driver appears in ODBC Data Source Administrator.

### MySQL connection error

This can happen if:

- MySQL Server is not running.
- The username or password is incorrect.
- The target database does not exist.
- The MySQL connector package is missing.

Create the database first and verify your MySQL credentials.

## Notes

- The script drops and recreates each target table before loading data.
- Existing data in target tables with the same names can be removed.
- Column names are converted to lowercase and spaces are replaced with underscores.
- Empty strings and missing values are converted to `NULL` where possible.
- Boolean values are converted to `0` and `1`.

## Author

Project repository:

```text
Yog964/Data_Migration_From_MAccess_To_Mysql
```
