"""
Database utility functions for PWD Tools
Handle SQLite database operations for storing and retrieving data
"""

import sqlite3
import pandas as pd
from datetime import datetime
import os
import json

class DatabaseManager:
    """Database manager for PWD Tools application"""
    
    def __init__(self, db_path="pwd_tools.db"):
        """Initialize database manager"""
        self.db_path = db_path
        self.init_database()
    
    def init_database(self):
        """Initialize database with required tables"""
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            
            # Bills table
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS bills (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    bill_number TEXT NOT NULL,
                    bill_date DATE,
                    contractor_name TEXT,
                    project_name TEXT,
                    bill_amount REAL,
                    status TEXT,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            """)
            
            # EMD refunds table
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS emd_refunds (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    tender_number TEXT NOT NULL,
                    contractor_name TEXT,
                    emd_amount REAL,
                    deposit_date DATE,
                    refund_date DATE,
                    interest_rate REAL,
                    refund_amount REAL,
                    status TEXT,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            """)
            
            # Projects table
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS projects (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    project_name TEXT NOT NULL,
                    project_code TEXT,
                    contractor_name TEXT,
                    agreement_amount REAL,
                    start_date DATE,
                    completion_date DATE,
                    status TEXT,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            """)
            
            # Deductions table
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS deductions (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    bill_id INTEGER,
                    deduction_type TEXT,
                    amount REAL,
                    rate REAL,
                    is_statutory BOOLEAN,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    FOREIGN KEY (bill_id) REFERENCES bills (id)
                )
            """)
            
            # Settings table
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS settings (
                    key TEXT PRIMARY KEY,
                    value TEXT,
                    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            """)
            
            conn.commit()
    
    def save_bill(self, bill_data):
        """Save bill information to database"""
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            
            cursor.execute("""
                INSERT INTO bills (bill_number, bill_date, contractor_name, 
                                 project_name, bill_amount, status)
                VALUES (?, ?, ?, ?, ?, ?)
            """, (
                bill_data.get('bill_number'),
                bill_data.get('bill_date'),
                bill_data.get('contractor_name'),
                bill_data.get('project_name'),
                bill_data.get('bill_amount'),
                bill_data.get('status', 'Active')
            ))
            
            bill_id = cursor.lastrowid
            
            # Save deductions if any
            if 'deductions' in bill_data:
                for deduction in bill_data['deductions']:
                    cursor.execute("""
                        INSERT INTO deductions (bill_id, deduction_type, amount, rate, is_statutory)
                        VALUES (?, ?, ?, ?, ?)
                    """, (
                        bill_id,
                        deduction['type'],
                        deduction['amount'],
                        deduction.get('rate', 0),
                        deduction.get('statutory', False)
                    ))
            
            conn.commit()
            return bill_id
    
    def save_emd_refund(self, emd_data):
        """Save EMD refund information to database"""
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            
            cursor.execute("""
                INSERT INTO emd_refunds (tender_number, contractor_name, emd_amount,
                                       deposit_date, refund_date, interest_rate, 
                                       refund_amount, status)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                emd_data.get('tender_number'),
                emd_data.get('contractor_name'),
                emd_data.get('emd_amount'),
                emd_data.get('deposit_date'),
                emd_data.get('refund_date'),
                emd_data.get('interest_rate'),
                emd_data.get('refund_amount'),
                emd_data.get('status', 'Processed')
            ))
            
            conn.commit()
            return cursor.lastrowid
    
    def save_project(self, project_data):
        """Save project information to database"""
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            
            cursor.execute("""
                INSERT INTO projects (project_name, project_code, contractor_name,
                                    agreement_amount, start_date, completion_date, status)
                VALUES (?, ?, ?, ?, ?, ?, ?)
            """, (
                project_data.get('project_name'),
                project_data.get('project_code'),
                project_data.get('contractor_name'),
                project_data.get('agreement_amount'),
                project_data.get('start_date'),
                project_data.get('completion_date'),
                project_data.get('status', 'Active')
            ))
            
            conn.commit()
            return cursor.lastrowid
    
    def get_bills(self, limit=None):
        """Retrieve bills from database"""
        with sqlite3.connect(self.db_path) as conn:
            query = "SELECT * FROM bills ORDER BY created_at DESC"
            if limit:
                query += f" LIMIT {limit}"
            
            df = pd.read_sql_query(query, conn)
            return df
    
    def get_emd_refunds(self, limit=None):
        """Retrieve EMD refunds from database"""
        with sqlite3.connect(self.db_path) as conn:
            query = "SELECT * FROM emd_refunds ORDER BY created_at DESC"
            if limit:
                query += f" LIMIT {limit}"
            
            df = pd.read_sql_query(query, conn)
            return df
    
    def get_projects(self, limit=None):
        """Retrieve projects from database"""
        with sqlite3.connect(self.db_path) as conn:
            query = "SELECT * FROM projects ORDER BY created_at DESC"
            if limit:
                query += f" LIMIT {limit}"
            
            df = pd.read_sql_query(query, conn)
            return df
    
    def get_bill_with_deductions(self, bill_id):
        """Get bill with associated deductions"""
        with sqlite3.connect(self.db_path) as conn:
            # Get bill
            bill_query = "SELECT * FROM bills WHERE id = ?"
            bill_df = pd.read_sql_query(bill_query, conn, params=(bill_id,))
            
            # Get deductions
            deduction_query = "SELECT * FROM deductions WHERE bill_id = ?"
            deduction_df = pd.read_sql_query(deduction_query, conn, params=(bill_id,))
            
            return bill_df, deduction_df
    
    def search_records(self, table, search_term, columns=None):
        """Search records in specified table"""
        if not columns:
            # Default search columns for each table
            search_columns = {
                'bills': ['bill_number', 'contractor_name', 'project_name'],
                'emd_refunds': ['tender_number', 'contractor_name'],
                'projects': ['project_name', 'project_code', 'contractor_name']
            }
            columns = search_columns.get(table, [])
        
        if not columns:
            return pd.DataFrame()
        
        with sqlite3.connect(self.db_path) as conn:
            # Build search query
            conditions = []
            params = []
            
            for column in columns:
                conditions.append(f"{column} LIKE ?")
                params.append(f"%{search_term}%")
            
            where_clause = " OR ".join(conditions)
            query = f"SELECT * FROM {table} WHERE {where_clause} ORDER BY created_at DESC"
            
            df = pd.read_sql_query(query, conn, params=params)
            return df
    
    def update_record(self, table, record_id, updates):
        """Update record in specified table"""
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            
            # Build update query
            set_clauses = []
            params = []
            
            for column, value in updates.items():
                set_clauses.append(f"{column} = ?")
                params.append(value)
            
            params.append(record_id)
            
            set_clause = ", ".join(set_clauses)
            query = f"UPDATE {table} SET {set_clause} WHERE id = ?"
            
            cursor.execute(query, params)
            conn.commit()
            
            return cursor.rowcount > 0
    
    def delete_record(self, table, record_id):
        """Delete record from specified table"""
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            
            cursor.execute(f"DELETE FROM {table} WHERE id = ?", (record_id,))
            conn.commit()
            
            return cursor.rowcount > 0
    
    def get_statistics(self):
        """Get database statistics"""
        stats = {}
        
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            
            # Count records in each table
            tables = ['bills', 'emd_refunds', 'projects', 'deductions']
            
            for table in tables:
                cursor.execute(f"SELECT COUNT(*) FROM {table}")
                stats[f"total_{table}"] = cursor.fetchone()[0]
            
            # Additional statistics
            cursor.execute("SELECT SUM(bill_amount) FROM bills WHERE status = 'Active'")
            result = cursor.fetchone()[0]
            stats['total_active_bills_amount'] = result if result else 0
            
            cursor.execute("SELECT SUM(emd_amount) FROM emd_refunds")
            result = cursor.fetchone()[0]
            stats['total_emd_amount'] = result if result else 0
            
            cursor.execute("SELECT SUM(agreement_amount) FROM projects WHERE status = 'Active'")
            result = cursor.fetchone()[0]
            stats['total_active_projects_value'] = result if result else 0
        
        return stats
    
    def backup_database(self, backup_path=None):
        """Create database backup"""
        if not backup_path:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_path = f"pwd_tools_backup_{timestamp}.db"
        
        try:
            # Simple file copy for SQLite
            import shutil
            shutil.copy2(self.db_path, backup_path)
            return backup_path
        except Exception as e:
            raise Exception(f"Backup failed: {str(e)}")
    
    def export_to_csv(self, table, file_path=None):
        """Export table data to CSV"""
        if not file_path:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            file_path = f"{table}_export_{timestamp}.csv"
        
        with sqlite3.connect(self.db_path) as conn:
            df = pd.read_sql_query(f"SELECT * FROM {table}", conn)
            df.to_csv(file_path, index=False)
            
        return file_path
    
    def close(self):
        """Close database connection (for cleanup)"""
        # SQLite connections are closed automatically when using context manager
        pass

# Example usage functions
def get_db_manager():
    """Get database manager instance"""
    return DatabaseManager()

def init_sample_data():
    """Initialize database with sample data for testing"""
    db = DatabaseManager()
    
    # Sample bill
    sample_bill = {
        'bill_number': 'B001/2024',
        'bill_date': '2024-01-15',
        'contractor_name': 'ABC Construction',
        'project_name': 'Road Construction Project',
        'bill_amount': 500000,
        'status': 'Active',
        'deductions': [
            {'type': 'Income Tax', 'amount': 5000, 'rate': 1.0, 'statutory': True},
            {'type': 'Security Deposit', 'amount': 25000, 'rate': 5.0, 'statutory': False}
        ]
    }
    
    bill_id = db.save_bill(sample_bill)
    
    # Sample EMD refund
    sample_emd = {
        'tender_number': 'T001/2024',
        'contractor_name': 'XYZ Builders',
        'emd_amount': 50000,
        'deposit_date': '2024-01-10',
        'refund_date': '2024-02-10',
        'interest_rate': 6.0,
        'refund_amount': 50500,
        'status': 'Processed'
    }
    
    emd_id = db.save_emd_refund(sample_emd)
    
    # Sample project
    sample_project = {
        'project_name': 'Bridge Construction',
        'project_code': 'PWD/BRG/001',
        'contractor_name': 'PQR Infrastructure',
        'agreement_amount': 2000000,
        'start_date': '2024-01-01',
        'completion_date': '2024-12-31',
        'status': 'Active'
    }
    
    project_id = db.save_project(sample_project)
    
    return {
        'bill_id': bill_id,
        'emd_id': emd_id,
        'project_id': project_id
    }
