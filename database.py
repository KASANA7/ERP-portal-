import sqlite3
import os

DB_PATH = os.path.join(os.path.dirname(__file__), 'cms.db')

def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_db()
    c = conn.cursor()

    # Users
    c.execute('''CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT UNIQUE NOT NULL,
        password_hash TEXT NOT NULL,
        full_name TEXT,
        role TEXT DEFAULT 'admin',
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )''')

    # Companies
    c.execute('''CREATE TABLE IF NOT EXISTS companies (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL,
        gst TEXT,
        contact_person TEXT,
        contact_number TEXT,
        email TEXT,
        esic_pf_code TEXT,
        address TEXT,
        work_type TEXT,
        num_supervisors INTEGER DEFAULT 0,
        num_workers INTEGER DEFAULT 0,
        risk_level TEXT,
        contract_start TEXT,
        contract_expiry TEXT,
        nature_of_work TEXT,
        parta_checklist TEXT,
        partb_checklist TEXT,
        hr_status TEXT DEFAULT 'Pending',
        hr_approved_by TEXT,
        hr_approved_at TEXT,
        hr_remarks TEXT,
        safety_status TEXT DEFAULT 'Pending',
        safety_approved_by TEXT,
        safety_approved_at TEXT,
        safety_remarks TEXT,
        cluster_approved_by TEXT,
        cluster_approved_at TEXT,
        cluster_remarks TEXT,
        status TEXT DEFAULT 'Draft',
        is_active INTEGER DEFAULT 1,
        draft_step INTEGER DEFAULT 1,
        draft_data TEXT,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )''')

    # Company Documents - split Part A and Part B
    c.execute('''CREATE TABLE IF NOT EXISTS company_docs (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        company_id INTEGER,
        doc_name TEXT,
        file_path TEXT,
        part TEXT DEFAULT 'A',
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (company_id) REFERENCES companies(id)
    )''')

    # Supervisors
    c.execute('''CREATE TABLE IF NOT EXISTS supervisors (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        company_id INTEGER,
        first_name TEXT,
        middle_name TEXT,
        last_name TEXT,
        dob TEXT,
        gender TEXT,
        mobile TEXT,
        aadhar TEXT,
        address TEXT,
        languages TEXT,
        qualifications TEXT,
        work_type TEXT,
        num_workers INTEGER DEFAULT 0,
        experience_years INTEGER DEFAULT 0,
        experience_months INTEGER DEFAULT 0,
        technical_qualification TEXT,
        jobs_handled TEXT,
        risk_level TEXT,
        interview_data TEXT,
        interview_result TEXT,
        remarks TEXT,
        status TEXT DEFAULT 'Pending',
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (company_id) REFERENCES companies(id)
    )''')

    # Supervisor Documents
    c.execute('''CREATE TABLE IF NOT EXISTS supervisor_docs (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        supervisor_id INTEGER,
        doc_name TEXT,
        file_path TEXT,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (supervisor_id) REFERENCES supervisors(id)
    )''')

    # Worker Cards
    c.execute('''CREATE TABLE IF NOT EXISTS worker_cards (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        card_number TEXT UNIQUE,
        company_id INTEGER,
        supervisor_id INTEGER,
        worker_type TEXT DEFAULT 'workman',
        first_name TEXT,
        middle_name TEXT,
        last_name TEXT,
        dob TEXT,
        aadhar TEXT,
        blood_group TEXT,
        emergency_contact TEXT,
        designation TEXT,
        risk_level TEXT,
        photo_path TEXT,
        medical_valid TEXT,
        induction_valid TEXT,
        hr_status TEXT DEFAULT 'Pending',
        hr_remarks TEXT,
        safety_status TEXT DEFAULT 'Pending',
        safety_remarks TEXT,
        status TEXT DEFAULT 'Pending',
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (company_id) REFERENCES companies(id),
        FOREIGN KEY (supervisor_id) REFERENCES supervisors(id)
    )''')

    # Worker Documents
    c.execute('''CREATE TABLE IF NOT EXISTS worker_docs (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        worker_id INTEGER,
        doc_name TEXT,
        file_path TEXT,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (worker_id) REFERENCES worker_cards(id)
    )''')

    conn.commit()
    conn.close()
    print("Database ready.")

def generate_card_number(company_id):
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT COUNT(*) FROM worker_cards WHERE company_id = ?", (company_id,))
    count = c.fetchone()[0] + 1
    c.execute("SELECT name FROM companies WHERE id = ?", (company_id,))
    co = c.fetchone()
    conn.close()
    prefix = ''.join([w[0] for w in co['name'].split()[:3]]).upper() if co else 'CMC'
    return f"{prefix}/{str(count).zfill(5)}"

def migrate_db():
    """Add new columns to existing database without losing data"""
    conn = get_db()
    c = conn.cursor()
    new_cols = [
        ("companies", "contract_expiry", "TEXT"),
        ("companies", "nature_of_work", "TEXT"),
        ("companies", "parta_checklist", "TEXT"),
        ("companies", "partb_checklist", "TEXT"),
        ("companies", "hr_status", "TEXT DEFAULT 'Pending'"),
        ("companies", "hr_approved_by", "TEXT"),
        ("companies", "hr_approved_at", "TEXT"),
        ("companies", "hr_remarks", "TEXT"),
        ("companies", "safety_status", "TEXT DEFAULT 'Pending'"),
        ("companies", "safety_approved_by", "TEXT"),
        ("companies", "safety_approved_at", "TEXT"),
        ("companies", "safety_remarks", "TEXT"),
        ("companies", "cluster_approved_by", "TEXT"),
        ("companies", "cluster_approved_at", "TEXT"),
        ("companies", "cluster_remarks", "TEXT"),
        ("companies", "is_active", "INTEGER DEFAULT 1"),
        ("companies", "draft_step", "INTEGER DEFAULT 1"),
        ("companies", "draft_data", "TEXT"),
        ("company_docs", "part", "TEXT DEFAULT 'A'"),
        ("worker_cards", "photo_path", "TEXT"),
        ("worker_cards", "medical_valid", "TEXT"),
        ("worker_cards", "induction_valid", "TEXT"),
        ("worker_cards", "induction_modules", "TEXT"),
        ("worker_cards", "safety_approved_by", "TEXT"),
        ("worker_cards", "safety_approved_at", "TEXT"),
        ("worker_cards", "reinduction_date", "TEXT"),
        ("worker_cards", "reinduction_by", "TEXT"),
        ("worker_cards", "reinduction_at", "TEXT"),
        ("supervisors", "nature_of_work", "TEXT"),
        ("supervisors", "draft_step", "INTEGER DEFAULT 1"),
        ("supervisors", "draft_data", "TEXT"),
        ("supervisors", "status", "TEXT DEFAULT 'Pending'"),
        ("companies", "vendor_code", "TEXT"),
    ]
    for table, col, col_type in new_cols:
        try:
            c.execute(f"ALTER TABLE {table} ADD COLUMN {col} {col_type}")
        except:
            pass  # Column already exists
    conn.commit()
    conn.close()

if __name__ == '__main__':
    init_db()
    migrate_db()
    print("Database initialized at:", DB_PATH)
