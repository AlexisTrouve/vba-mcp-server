"""
Create sample Access database with VBA macros for testing.

Requires: Windows + Microsoft Access installed
"""

import sys
from pathlib import Path

try:
    import win32com.client
except ImportError:
    print("Error: pywin32 required. Install with: pip install pywin32")
    sys.exit(1)


SAMPLE_VBA_CODE = '''
Option Explicit

Public Const APP_VERSION As String = "1.0.0"

Public Function HelloAccess() As String
    HelloAccess = "Hello from Access VBA!"
End Function

Public Function GetRecordCount(tableName As String) As Long
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    Set db = CurrentDb()
    Set rs = db.OpenRecordset("SELECT COUNT(*) FROM [" & tableName & "]")

    GetRecordCount = rs.Fields(0).Value

    rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function

ErrorHandler:
    GetRecordCount = -1
End Function

Public Sub LogMessage(message As String)
    Debug.Print "[" & Now & "] " & message
End Sub

Public Function AddNumbers(a As Double, b As Double) As Double
    AddNumbers = a + b
End Function
'''


def create_sample_accdb():
    """Create sample.accdb with test VBA code and sample data."""
    print("Creating sample Access database with VBA...")

    output_path = Path(__file__).parent / "sample.accdb"

    # Delete existing file if present
    if output_path.exists():
        output_path.unlink()
        print(f"  Deleted existing: {output_path}")

    # Launch Access
    access = win32com.client.Dispatch("Access.Application")
    access.Visible = False

    try:
        # Create new database
        access.NewCurrentDatabase(str(output_path.absolute()))
        print(f"  Created database: {output_path}")

        # Create sample table using SQL
        db = access.CurrentDb()

        # Create Employees table
        sql_create = """
        CREATE TABLE Employees (
            ID AUTOINCREMENT PRIMARY KEY,
            FirstName TEXT(50),
            LastName TEXT(50),
            Department TEXT(50),
            Salary CURRENCY
        )
        """
        db.Execute(sql_create)
        print("  Created table: Employees")

        # Insert sample data
        sample_data = [
            ("John", "Doe", "IT", 75000),
            ("Jane", "Smith", "HR", 65000),
            ("Bob", "Johnson", "Sales", 70000),
            ("Alice", "Williams", "IT", 80000),
            ("Charlie", "Brown", "Marketing", 60000),
        ]

        for first, last, dept, salary in sample_data:
            sql_insert = f"""
            INSERT INTO Employees (FirstName, LastName, Department, Salary)
            VALUES ('{first}', '{last}', '{dept}', {salary})
            """
            db.Execute(sql_insert)
        print(f"  Inserted {len(sample_data)} records")

        # Create Products table
        sql_create_products = """
        CREATE TABLE Products (
            ProductID AUTOINCREMENT PRIMARY KEY,
            ProductName TEXT(100),
            Category TEXT(50),
            Price CURRENCY,
            InStock INTEGER
        )
        """
        db.Execute(sql_create_products)
        print("  Created table: Products")

        # Insert product data
        products_data = [
            ("Laptop", "Electronics", 999.99, 50),
            ("Mouse", "Electronics", 29.99, 200),
            ("Desk Chair", "Furniture", 299.99, 30),
            ("Monitor", "Electronics", 399.99, 75),
            ("Keyboard", "Electronics", 79.99, 150),
        ]

        for name, cat, price, stock in products_data:
            sql_insert = f"""
            INSERT INTO Products (ProductName, Category, Price, InStock)
            VALUES ('{name}', '{cat}', {price}, {stock})
            """
            db.Execute(sql_insert)
        print(f"  Inserted {len(products_data)} products")

        # Add VBA module
        vb_project = access.VBE.ActiveVBProject
        vb_module = vb_project.VBComponents.Add(1)  # 1 = vbext_ct_StdModule
        vb_module.Name = "TestModule"
        vb_module.CodeModule.AddFromString(SAMPLE_VBA_CODE)
        print("  Added VBA module: TestModule")

        # Close database (auto-saves)
        access.CloseCurrentDatabase()

        print(f"\n[OK] Created: {output_path}")
        print(f"    Tables: Employees (5 rows), Products (5 rows)")
        print(f"    VBA Module: TestModule with 4 procedures")

    except Exception as e:
        print(f"\n[ERROR] {e}")
        import traceback
        traceback.print_exc()
        raise
    finally:
        try:
            access.Quit()
        except:
            pass


if __name__ == "__main__":
    create_sample_accdb()
