import os
import sqlite3
from datetime import datetime, timedelta
import pandas as pd

# إعداد المسارات
script_dir = os.path.dirname(os.path.abspath(__file__))
DB_NAME = os.path.join(script_dir, 'document_management.db')
ATTACHMENTS_DIR = os.path.join(script_dir, 'attachments')

if not os.path.exists(ATTACHMENTS_DIR):
    os.makedirs(ATTACHMENTS_DIR)

# --- دوال تحويل التاريخ ---
def convert_date_to_db_format(date_str_ddmmyyyy):
    """
    يحول تنسيق التاريخ من DD-MM-YYYY إلى YYYY-MM-DD للتخزين في قاعدة البيانات.
    يرفع ValueError إذا كان التنسيق غير صحيح.
    """
    if not date_str_ddmmyyyy:
        return None
    try:
        return datetime.strptime(date_str_ddmmyyyy, "%d-%m-%Y").strftime("%Y-%m-%d")
    except ValueError:
        raise ValueError("تنسيق تاريخ غير صحيح. يرجى استخدام DD-MM-YYYY.")

def convert_date_from_db_format(date_str_yyyymmdd):
    """
    يحول تنسيق التاريخ من YYYY-MM-DD (من قاعدة البيانات) إلى DD-MM-YYYY للعرض.
    يعيد سلسلة فارغة إذا كان التاريخ فارغًا أو غير صحيح.
    """
    if not date_str_yyyymmdd:
        return ""
    try:
        return datetime.strptime(date_str_yyyymmdd, "%Y-%m-%d").strftime("%d-%m-%Y")
    except ValueError:
        return ""

# --- إنشاء قاعدة البيانات ---
def create_database():
    """
    يقوم بإنشاء جداول قاعدة البيانات إذا لم تكن موجودة، ويضيف الفهارس لتحسين الأداء.
    """
    with sqlite3.connect(DB_NAME) as conn:
        cursor = conn.cursor()

        # جدول المستندات
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS documents (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                number TEXT NOT NULL UNIQUE,
                type TEXT NOT NULL,
                category TEXT NOT NULL,
                issue_date TEXT,
                expiry_date TEXT,
                status TEXT,
                employee_id INTEGER,
                notes TEXT,
                FOREIGN KEY(employee_id) REFERENCES employees(id)
            )
        ''')

        # جدول الموظفين
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS employees (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                position TEXT,
                department TEXT,
                start_date TEXT,
                phone TEXT,
                email TEXT UNIQUE,
                address TEXT,
                notes TEXT
            )
        ''')

        # جدول الرواتب
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS salaries (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                employee_id INTEGER NOT NULL,
                basic_salary REAL NOT NULL,
                allowances REAL DEFAULT 0,
                deductions REAL DEFAULT 0,
                net_salary REAL NOT NULL,
                payment_method TEXT,
                payment_date TEXT NOT NULL,
                FOREIGN KEY(employee_id) REFERENCES employees(id)
            )
        ''')

        # جدول سجل التدقيق (Audit Log)
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS audit_log (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                timestamp TEXT NOT NULL,
                action TEXT NOT NULL,
                details TEXT
            )
        ''')

        # إنشاء الفهارس لتحسين أداء الاستعلامات
        # فهرسة على الأعمدة المستخدمة بشكل متكرر في شروط WHERE و JOIN
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_documents_employee_id ON documents (employee_id)")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_documents_number ON documents (number)") # مهم للبحث عن المستندات برقم
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_documents_expiry_date ON documents (expiry_date)") # للبحث عن المستندات التي قاربت على الانتهاء

        cursor.execute("CREATE INDEX IF NOT EXISTS idx_employees_name ON employees (name)") # للبحث عن الموظفين بالاسم
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_employees_department ON employees (department)") # للبحث عن الموظفين بالقسم

        cursor.execute("CREATE INDEX IF NOT EXISTS idx_salaries_employee_id ON salaries (employee_id)")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_salaries_payment_date ON salaries (payment_date)") # مهم للبحث عن الرواتب حسب التاريخ

        conn.commit()

# --- دوال إدارة المستندات ---
def add_document(name, number, doc_type, category, issue_date, expiry_date, status, employee_id, notes):
    """يضيف مستندًا جديدًا إلى قاعدة البيانات."""
    with sqlite3.connect(DB_NAME) as conn:
        cursor = conn.cursor()
        try:
            cursor.execute('''
                INSERT INTO documents (name, number, type, category, issue_date, expiry_date, status, employee_id, notes)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (name, number, doc_type, category, issue_date, expiry_date, status, employee_id, notes))
            conn.commit()
            log_audit_event(f"تمت إضافة مستند جديد: {name} (رقم: {number})")
            return True
        except sqlite3.IntegrityError:
            return False # يشير إلى أن رقم المستند مكرر
        except Exception as e:
            print(f"خطأ عند إضافة مستند: {e}")
            return False

def update_document(doc_id, name, number, doc_type, category, issue_date, expiry_date, status, employee_id, notes):
    """يقوم بتحديث مستند موجود في قاعدة البيانات."""
    with sqlite3.connect(DB_NAME) as conn:
        cursor = conn.cursor()
        try:
            cursor.execute('''
                UPDATE documents
                SET name=?, number=?, type=?, category=?, issue_date=?, expiry_date=?, status=?, employee_id=?, notes=?
                WHERE id=?
            ''', (name, number, doc_type, category, issue_date, expiry_date, status, employee_id, notes, doc_id))
            conn.commit()
            if cursor.rowcount > 0:
                log_audit_event(f"تم تحديث المستند ID: {doc_id} (رقم: {number})")
                return True
            return False
        except sqlite3.IntegrityError:
            return False # يشير إلى أن رقم المستند مكرر
        except Exception as e:
            print(f"خطأ عند تحديث مستند: {e}")
            return False

def delete_document(doc_id):
    """يحذف مستندًا من قاعدة البيانات ويحذف المرفقات المرتبطة به."""
    with sqlite3.connect(DB_NAME) as conn:
        cursor = conn.cursor()
        try:
            doc_info = cursor.execute("SELECT name, number FROM documents WHERE id = ?", (doc_id,)).fetchone()
            if doc_info:
                doc_name, doc_number = doc_info
                # حذف المرفقات أولاً
                attachments = get_attachments_for_document(doc_id)
                for att in attachments:
                    delete_attachment(att[0]) # att[0] هو attachment_id

                cursor.execute("DELETE FROM documents WHERE id=?", (doc_id,))
                conn.commit()
                if cursor.rowcount > 0:
                    log_audit_event(f"تم حذف المستند ID: {doc_id} (الاسم: {doc_name}, الرقم: {doc_number}) وجميع مرفقاته.")
                    return True
            return False
        except Exception as e:
            print(f"خطأ عند حذف مستند: {e}")
            return False

def fetch_all_documents():
    """يجلب جميع المستندات من قاعدة البيانات."""
    with sqlite3.connect(DB_NAME) as conn:
        cursor = conn.cursor()
        cursor.execute('''
            SELECT d.id, d.name, d.number, d.type, d.category, d.issue_date, d.expiry_date, d.status, e.name, d.notes
            FROM documents d
            LEFT JOIN employees e ON d.employee_id = e.id
        ''')
        return cursor.fetchall()

def fetch_all_documents_for_export():
    """يجلب جميع المستندات مع أسماء الموظفين لتصديرها."""
    with sqlite3.connect(DB_NAME) as conn:
        cursor = conn.cursor()
        cursor.execute('''
            SELECT d.name, d.number, d.type, d.category, d.issue_date, d.expiry_date, d.status, e.name AS employee_name, d.notes
            FROM documents d
            LEFT JOIN employees e ON d.employee_id = e.id
            ORDER BY d.id
        ''')
        return cursor.fetchall()

def get_all_categories():
    """يجلب جميع الفئات الفريدة للمستندات."""
    with sqlite3.connect(DB_NAME) as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT DISTINCT category FROM documents WHERE category IS NOT NULL AND category != ''")
        categories = [row[0] for row in cursor.fetchall()]
        return sorted(categories)

def calculate_remaining_time(expiry_date_str_db):
    """
    يحسب الوقت المتبقي (أو المنقضي) من تاريخ انتهاء الصلاحية.
    يفترض أن تاريخ انتهاء الصلاحية بتنسيق YYYY-MM-DD.
    """
    if not expiry_date_str_db:
        return "غير محدد"
    try:
        expiry_date = datetime.strptime(expiry_date_str_db, "%Y-%m-%d")
        today = datetime.now()
        remaining = expiry_date - today
        
        if remaining.days > 0:
            return f"متبقي {remaining.days} يوم"
        elif remaining.days == 0:
            return "ينتهي اليوم"
        else:
            return f"منقضي {abs(remaining.days)} يوم"
    except ValueError:
        return "تاريخ غير صالح"

# --- دوال إدارة المرفقات ---
def add_attachment(document_id, file_path, file_name):
    """
    يضيف مرفقًا لمستند معين. يتم حفظ الملف فعليًا في مجلد المرفقات.
    """
    try:
        # التأكد من وجود مجلد المرفقات
        if not os.path.exists(ATTACHMENTS_DIR):
            os.makedirs(ATTACHMENTS_DIR)

        # إنشاء مسار فريد للملف داخل مجلد المرفقات
        unique_file_name = f"{document_id}_{datetime.now().strftime('%Y%m%d%H%M%S')}_{file_name}"
        destination_path = os.path.join(ATTACHMENTS_DIR, unique_file_name)

        # نسخ الملف إلى مجلد المرفقات
        import shutil
        shutil.copy(file_path, destination_path)

        with sqlite3.connect(DB_NAME) as conn:
            cursor = conn.cursor()
            cursor.execute('''
                INSERT INTO attachments (document_id, file_name, file_path)
                VALUES (?, ?, ?)
            ''', (document_id, file_name, destination_path))
            conn.commit()
            log_audit_event(f"تمت إضافة مرفق '{file_name}' للمستند ID: {document_id}")
            return True
    except Exception as e:
        print(f"خطأ عند إضافة مرفق: {e}")
        return False

def get_attachments_for_document(document_id):
    """يجلب جميع المرفقات لمستند معين."""
    with sqlite3.connect(DB_NAME) as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT id, file_name, file_path FROM attachments WHERE document_id=?", (document_id,))
        return cursor.fetchall()

def delete_attachment(attachment_id):
    """يحذف مرفقًا من قاعدة البيانات ومن نظام الملفات."""
    with sqlite3.connect(DB_NAME) as conn:
        cursor = conn.cursor()
        file_info = cursor.execute("SELECT file_name, file_path FROM attachments WHERE id=?", (attachment_id,)).fetchone()
        if file_info:
            file_name, file_path = file_info
            try:
                if os.path.exists(file_path):
                    os.remove(file_path)
                cursor.execute("DELETE FROM attachments WHERE id=?", (attachment_id,))
                conn.commit()
                log_audit_event(f"تم حذف مرفق ID: {attachment_id} (الاسم: {file_name})")
                return True
            except Exception as e:
                print(f"خطأ عند حذف المرفق: {e}")
                return False
        return False

# --- دوال إدارة الموظفين ---
def add_employee(name, position, department, start_date, phone, email, address, notes):
    """يضيف موظفًا جديدًا إلى قاعدة البيانات."""
    with sqlite3.connect(DB_NAME) as conn:
        cursor = conn.cursor()
        try:
            cursor.execute('''
                INSERT INTO employees (name, position, department, start_date, phone, email, address, notes)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            ''', (name, position, department, start_date, phone, email, address, notes))
            conn.commit()
            log_audit_event(f"تمت إضافة موظف جديد: {name}")
            return True
        except sqlite3.IntegrityError:
            return False # يشير إلى أن البريد الإلكتروني مكرر
        except Exception as e:
            print(f"خطأ عند إضافة موظف: {e}")
            return False

def update_employee(emp_id, name, position, department, start_date, phone, email, address, notes):
    """يقوم بتحديث بيانات موظف موجود."""
    with sqlite3.connect(DB_NAME) as conn:
        cursor = conn.cursor()
        try:
            cursor.execute('''
                UPDATE employees
                SET name=?, position=?, department=?, start_date=?, phone=?, email=?, address=?, notes=?
                WHERE id=?
            ''', (name, position, department, start_date, phone, email, address, notes, emp_id))
            conn.commit()
            if cursor.rowcount > 0:
                log_audit_event(f"تم تحديث بيانات الموظف ID: {emp_id} (الاسم: {name})")
                return True
            return False
        except sqlite3.IntegrityError:
            return False # يشير إلى أن البريد الإلكتروني مكرر
        except Exception as e:
            print(f"خطأ عند تحديث موظف: {e}")
            return False

def delete_employee(emp_id):
    """يحذف موظفًا من قاعدة البيانات ويزيل ارتباط المستندات والرواتب به."""
    with sqlite3.connect(DB_NAME) as conn:
        cursor = conn.cursor()
        try:
            emp_name = cursor.execute("SELECT name FROM employees WHERE id = ?", (emp_id,)).fetchone()
            if emp_name:
                # تحديث المستندات المرتبطة بهذا الموظف لتعيين employee_id إلى NULL
                cursor.execute("UPDATE documents SET employee_id = NULL WHERE employee_id = ?", (emp_id,))
                # حذف جميع سجلات الرواتب المرتبطة بهذا الموظف
                cursor.execute("DELETE FROM salaries WHERE employee_id = ?", (emp_id,))
                # حذف الموظف نفسه
                cursor.execute("DELETE FROM employees WHERE id=?", (emp_id,))
                conn.commit()
                if cursor.rowcount > 0:
                    log_audit_event(f"تم حذف الموظف ID: {emp_id} (الاسم: {emp_name[0]}) وحذف رواتبه وتعديل مستنداته.")
                    return True
            return False
        except Exception as e:
            print(f"خطأ عند حذف موظف: {e}")
            return False

def fetch_all_employees():
    """يجلب جميع الموظفين من قاعدة البيانات."""
    with sqlite3.connect(DB_NAME) as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM employees")
        return cursor.fetchall()

def fetch_employee_id_name():
    """يجلب معرفات وأسماء جميع الموظفين."""
    with sqlite3.connect(DB_NAME) as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT id, name FROM employees")
        return cursor.fetchall()

def get_all_departments():
    """يجلب جميع الأقسام الفريدة للموظفين."""
    with sqlite3.connect(DB_NAME) as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT DISTINCT department FROM employees WHERE department IS NOT NULL AND department != ''")
        departments = [row[0] for row in cursor.fetchall()]
        return sorted(departments)

# --- دوال إدارة الرواتب ---
def calculate_net_salary(basic_salary, allowances, deductions):
    """يحسب صافي الراتب بناءً على الراتب الأساسي والبدلات والخصومات."""
    try:
        return float(basic_salary) + float(allowances) - float(deductions)
    except ValueError:
        return 0.0 # أو يمكنك رفع استثناء

def add_salary(employee_id, basic_salary, allowances, deductions, net_salary, payment_method, payment_date):
    """يضيف سجل راتب جديد إلى قاعدة البيانات."""
    with sqlite3.connect(DB_NAME) as conn:
        cursor = conn.cursor()
        try:
            cursor.execute('''
                INSERT INTO salaries (employee_id, basic_salary, allowances, deductions, net_salary, payment_method, payment_date)
                VALUES (?, ?, ?, ?, ?, ?, ?)
            ''', (employee_id, basic_salary, allowances, deductions, net_salary, payment_method, payment_date))
            conn.commit()
            log_audit_event(f"تمت إضافة راتب للموظف ID: {employee_id} بتاريخ {payment_date}")
            return True
        except Exception as e:
            print(f"خطأ عند إضافة راتب: {e}")
            return False

def update_salary(salary_id, employee_id, basic_salary, allowances, deductions, net_salary, payment_method, payment_date):
    """يقوم بتحديث سجل راتب موجود."""
    with sqlite3.connect(DB_NAME) as conn:
        cursor = conn.cursor()
        try:
            cursor.execute('''
                UPDATE salaries
                SET employee_id=?, basic_salary=?, allowances=?, deductions=?, net_salary=?, payment_method=?, payment_date=?
                WHERE id=?
            ''', (employee_id, basic_salary, allowances, deductions, net_salary, payment_method, payment_date, salary_id))
            conn.commit()
            if cursor.rowcount > 0:
                log_audit_event(f"تم تحديث راتب ID: {salary_id} للموظف ID: {employee_id}")
                return True
            return False
        except Exception as e:
            print(f"خطأ عند تحديث راتب: {e}")
            return False

def delete_salary(salary_id):
    """يحذف سجل راتب من قاعدة البيانات."""
    with sqlite3.connect(DB_NAME) as conn:
        cursor = conn.cursor()
        try:
            salary_info = cursor.execute("SELECT employee_id, payment_date FROM salaries WHERE id = ?", (salary_id,)).fetchone()
            if salary_info:
                emp_id, pay_date = salary_info
                cursor.execute("DELETE FROM salaries WHERE id=?", (salary_id,))
                conn.commit()
                if cursor.rowcount > 0:
                    log_audit_event(f"تم حذف راتب ID: {salary_id} للموظف ID: {emp_id} بتاريخ {pay_date}")
                    return True
            return False
        except Exception as e:
            print(f"خطأ عند حذف راتب: {e}")
            return False

def fetch_all_salaries():
    """يجلب جميع سجلات الرواتب مع أسماء الموظفين والأقسام."""
    with sqlite3.connect(DB_NAME) as conn:
        cursor = conn.cursor()
        cursor.execute('''
            SELECT s.id, e.name, e.department, s.basic_salary, s.allowances, s.deductions, s.net_salary, s.payment_method, s.payment_date
            FROM salaries s
            JOIN employees e ON s.employee_id = e.id
            ORDER BY s.payment_date DESC
        ''')
        rows = cursor.fetchall()
    
    # حساب الراتب السنوي الأساسي وتنسيق التاريخ للعرض
    processed_rows = []
    for row in rows:
        salary_id, emp_name, department, basic_salary, allowances, deductions, net_salary, payment_method, payment_date_db = row
        annual_basic_salary = basic_salary * 12
        payment_date_ddmmyyyy = convert_date_from_db_format(payment_date_db)
        processed_rows.append((salary_id, emp_name, department, basic_salary, annual_basic_salary, allowances, deductions, net_salary, payment_method, payment_date_ddmmyyyy))
    return processed_rows

def fetch_all_salaries_for_export():
    """يجلب جميع سجلات الرواتب مع أسماء الموظفين والأقسام لتصديرها."""
    with sqlite3.connect(DB_NAME) as conn:
        cursor = conn.cursor()
        cursor.execute('''
            SELECT e.name, e.department, s.basic_salary, s.allowances, s.deductions, s.net_salary, s.payment_method, s.payment_date
            FROM salaries s
            JOIN employees e ON s.employee_id = e.id
            ORDER BY s.payment_date DESC
        ''')
        rows = cursor.fetchall()
    
    processed_rows = []
    for row in rows:
        emp_name, department, basic_salary, allowances, deductions, net_salary, payment_method, payment_date_db = row
        annual_basic_salary = basic_salary * 12
        payment_date_ddmmyyyy = convert_date_from_db_format(payment_date_db)
        processed_rows.append((emp_name, department, basic_salary, annual_basic_salary, allowances, deductions, net_salary, payment_method, payment_date_ddmmyyyy))
    return processed_rows

def get_last_employee_salary(employee_id):
    """
    يجلب آخر راتب تم دفعه لموظف معين.
    """
    with sqlite3.connect(DB_NAME) as conn:
        cursor = conn.cursor()
        cursor.execute('''
            SELECT basic_salary, allowances, deductions, payment_method
            FROM salaries
            WHERE employee_id = ?
            ORDER BY payment_date DESC
            LIMIT 1
        ''', (employee_id,))
        return cursor.fetchone()

def salary_exists_for_month(employee_id, year, month):
    """
    يتحقق مما إذا كان هناك سجل راتب لموظف معين في شهر وسنة محددين.
    """
    with sqlite3.connect(DB_NAME) as conn:
        cursor = conn.cursor()
        # نستخدم دالة STRFTIME للمقارنة بالعام والشهر بشكل مباشر
        # هذا أكثر كفاءة مع الفهارس من LIKE إذا كانت الأعمدة من نوع TEXT وتخزن بتنسيق YYYY-MM-DD
        cursor.execute(f"""
            SELECT 1 FROM salaries
            WHERE employee_id = ? AND STRFTIME('%Y-%m', payment_date) = ?
            LIMIT 1
        """, (employee_id, f"{year}-{month:02d}")) # تأكد من أن الشهر بتنسيق 01, 02...12
        return cursor.fetchone() is not None

def fetch_employee_salary_history(employee_id):
    """
    يجلب جميع سجلات الرواتب لموظف معين، مرتبة تنازليًا حسب تاريخ الدفع.
    """
    with sqlite3.connect(DB_NAME) as conn:
        cursor = conn.cursor()
        cursor.execute("""
            SELECT id, basic_salary, allowances, deductions, net_salary, payment_method, payment_date
            FROM salaries
            WHERE employee_id = ?
            ORDER BY payment_date DESC
        """, (employee_id,))
        rows = cursor.fetchall()
    
    history_data = []
    for row in rows:
        salary_id, basic_salary, allowances, deductions, net_salary, payment_method, payment_date_db = row
        payment_date_ddmmyyyy = convert_date_from_db_format(payment_date_db)
        annual_basic_salary = basic_salary * 12
        history_data.append((salary_id, basic_salary, annual_basic_salary, allowances, deductions, net_salary, payment_method, payment_date_ddmmyyyy))
    return history_data

# --- دوال سجل التدقيق (Audit Log) ---
def log_audit_event(action, details=""):
    """يسجل حدثًا في سجل التدقيق."""
    with sqlite3.connect(DB_NAME) as conn:
        cursor = conn.cursor()
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        cursor.execute("INSERT INTO audit_log (timestamp, action, details) VALUES (?, ?, ?)", (timestamp, action, details))
        conn.commit()

def fetch_audit_log():
    """يجلب جميع أحداث سجل التدقيق، مرتبة تنازليًا حسب الوقت."""
    with sqlite3.connect(DB_NAME) as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT id, timestamp, action, details FROM audit_log ORDER BY timestamp DESC")
        return cursor.fetchall()