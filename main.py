from backend import (
    create_database,
    add_document,
    update_document,
    delete_document,
    add_employee,
    update_employee,
    delete_employee,
    fetch_employee_id_name,
    fetch_all_employees,
    fetch_audit_log,
    convert_date_from_db_format,
    convert_date_to_db_format,
    add_attachment,
    get_attachments_for_document,
    delete_attachment,
    get_all_categories,
    calculate_remaining_time,
    fetch_all_documents_for_export,
    log_audit_event,
    calculate_net_salary,
    add_salary,
    update_salary,
    delete_salary,
    fetch_all_salaries,
    get_all_departments,
    fetch_all_salaries_for_export,
    get_last_employee_salary, # New import
    salary_exists_for_month   # New import
)

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkcalendar import DateEntry
import sqlite3
from datetime import datetime, timedelta
import os
import subprocess
import pandas as pd

# إنشاء قاعدة البيانات
create_database()

root = tk.Tk()
root.title("نظام إدارة المستندات")
root.geometry("1400x850")

# تطبيق ثيم
style = ttk.Style()
style.theme_use('clam')

# تخصيص Treeview (أعمدة قابلة لتغيير الحجم والترتيب)
style.configure("Treeview.Heading", font=('Arial', 10, 'bold'))
style.configure("Treeview", rowheight=25)

# شريط الحالة
status_bar = ttk.Label(root, text="جاهز.", relief=tk.SUNKEN, anchor=tk.W)
status_bar.pack(side=tk.BOTTOM, fill=tk.X)

def set_status(message):
    """تحديث رسالة شريط الحالة."""
    status_bar.config(text=message)
    root.update_idletasks()

# --- نافذة تأكيد مخصصة ---
class CustomConfirmDialog(tk.Toplevel):
    """نافذة منبثقة مخصصة للتأكيد بدلاً من messagebox."""
    def __init__(self, parent, title, message):
        super().__init__(parent)
        self.title(title)
        self.transient(parent)
        self.grab_set()
        self.result = False

        parent_x = parent.winfo_x()
        parent_y = parent.winfo_y()
        parent_width = parent.winfo_width()
        parent_height = parent.winfo_height()

        dialog_width = 300
        dialog_height = 150
        x = parent_x + (parent_width // 2) - (dialog_width // 2)
        y = parent_y + (parent_height // 2) - (dialog_height // 2)
        self.geometry(f"{dialog_width}x{dialog_height}+{x}+{y}")
        self.resizable(False, False)

        ttk.Label(self, text=message, wraplength=250, font=('Arial', 10)).pack(pady=20)

        button_frame = ttk.Frame(self)
        button_frame.pack(pady=10)

        ttk.Button(button_frame, text="نعم", command=self._on_yes).pack(side=tk.LEFT, padx=10)
        ttk.Button(button_frame, text="لا", command=self._on_no).pack(side=tk.RIGHT, padx=10)

        self.wait_window(self)

    def _on_yes(self):
        self.result = True
        self.destroy()

    def _on_no(self):
        self.result = False
        self.destroy()

# --- دالة الفرز العامة للجداول ---
def treeview_sort_column(tv, col, reverse):
    """
    دالة لفرز بيانات Treeview بناءً على عمود معين.
    tv: كائن Treeview
    col: معرف العمود (اسم العمود)
    reverse: True للفرز التنازلي، False للفرز التصاعدي
    """
    l = [(tv.set(k, col), k) for k in tv.get_children('')]

    try:
        if col in ["id", "الرقم", "الرقم الوظيفي", "الراتب الأساسي (شهري)", "الراتب الأساسي (سنوي)", "البدلات", "الخصومات", "صافي الراتب"]:
            # تحويل إلى عدد حقيقي، مع التعامل مع القيم غير الرقمية بوضعها في النهاية
            l.sort(key=lambda t: float(t[0]) if str(t[0]).replace('.', '', 1).isdigit() else float('inf'), reverse=reverse)
        elif col in ["تاريخ الإصدار", "تاريخ الانتهاء", "تاريخ التعيين", "الوقت", "تاريخ الدفع"]:
            # تحويل إلى كائن تاريخ، مع التعامل مع القيم الفارغة بوضعها في البداية/النهاية
            l.sort(key=lambda t: datetime.strptime(t[0], "%d-%m-%Y") if t[0] and t[0] != "N/A" else (datetime.min if not reverse else datetime.max), reverse=reverse)
        else:
            l.sort(key=lambda t: t[0], reverse=reverse)
    except Exception as e:
        l.sort(key=lambda t: t[0], reverse=reverse)
        print(f"Warning: Could not sort column '{col}' numerically/date-wise. Sorting as text. Error: {e}")


    for index, (val, k) in enumerate(l):
        tv.move(k, '', index)

    tv.heading(col, command=lambda: treeview_sort_column(tv, col, not reverse))

# --- دالة اللصق لحقول الإدخال ---
def paste_event_handler(event):
    """
    يعالج حدث اللصق (Ctrl+V أو Command+V) لحقول الإدخال.
    يستخدم event_generate لمحاكاة حدث اللصق الافتراضي.
    """
    widget = root.focus_get()
    # تأكد من أن الودجت الحالي هو حقل إدخال
    if isinstance(widget, (ttk.Entry, tk.Entry)):
        widget.event_generate("<<Paste>>")
        return "break" # يوقف انتشار الحدث لمنع أي سلوك افتراضي آخر
    return None # يسمح بالحدث الافتراضي إذا لم يكن حقل إدخال

# ربط دالة اللصق بحدث Ctrl+V و Command+V على مستوى النافذة الرئيسية
root.bind_all("<Control-v>", paste_event_handler)
root.bind_all("<Command-v>", paste_event_handler) # لدعم macOS

# --- دوال مساعدة للمستندات ---
def get_row_color(expiry_date_str):
    """تحديد لون الصف بناءً على تاريخ الانتهاء."""
    try:
        if not expiry_date_str:
            return "valid"
        expiry = datetime.strptime(expiry_date_str, "%Y-%m-%d").date()
        today = datetime.today().date()
        if expiry < today:
            return "expired"
        elif expiry <= today + timedelta(days=90):
            return "near"
        else:
            return "valid"
    except:
        return "valid"

def clear_fields():
    """مسح جميع حقول إدخال المستندات."""
    for entry in entries:
        if isinstance(entry, DateEntry):
            entry.set_date(datetime.now().date())
        else:
            entry.delete(0, tk.END)
    attachments_table.delete(*attachments_table.get_children())
    set_status("تم مسح حقول المستندات.")

def save_document():
    """حفظ مستند جديد في قاعدة البيانات."""
    set_status("جاري حفظ المستند...")
    try:
        name = entry_name.get().strip()
        number = entry_number.get().strip()
        date = entry_date.get_date().strftime("%d-%m-%Y")
        expiry = entry_expiry.get_date().strftime("%d-%m-%Y") if entry_expiry.get_date() else ""
        issuer = entry_issuer.get().strip()
        category = entry_category.get().strip()
        tags = entry_tags.get().strip()

        doc_id = add_document(name, number, date, expiry, issuer, None, category, tags)
        messagebox.showinfo("نجاح", "تمت إضافة المستند بنجاح.")
        clear_fields()
        load_documents()
        update_category_filter_options()
        set_status("تم حفظ المستند بنجاح.")
    except ValueError as e:
        messagebox.showerror("خطأ في الإدخال", str(e))
        set_status(f"خطأ في الإدخال: {e}")
    except Exception as e:
        messagebox.showerror("خطأ", f"حدث خطأ غير متوقع: {e}")
        set_status(f"خطأ: {e}")

def populate_form_from_selection():
    """ملء حقول الإدخال بمعلومات المستند المحدد في الجدول."""
    selected = doc_table.selection()
    if not selected:
        clear_fields()
        return
    values = doc_table.item(selected[0])['values']
    clear_fields()
    
    entry_name.insert(0, values[1])
    entry_number.insert(0, values[2])
    
    date_ddmmyyyy = convert_date_from_db_format(values[3])
    entry_date.set_date(datetime.strptime(date_ddmmyyyy, "%d-%m-%Y").date())
    
    expiry_ddmmyyyy = convert_date_from_db_format(values[4])
    if expiry_ddmmyyyy:
        entry_expiry.set_date(datetime.strptime(expiry_ddmmyyyy, "%d-%m-%Y").date())
    else:
        entry_expiry.set_date("")

    entry_issuer.insert(0, values[5])
    entry_category.insert(0, values[6] if values[6] else "")
    entry_tags.insert(0, values[7] if values[7] else "")

    load_attachments(values[0])
    set_status(f"تم تحديد المستند: {values[1]}.")

def update_selected_document():
    """تحديث معلومات المستند المحدد في قاعدة البيانات."""
    selected = doc_table.selection()
    if not selected:
        messagebox.showwarning("تحذير", "يرجى تحديد مستند للتعديل.")
        return
    
    set_status("جاري تعديل المستند...")
    name = entry_name.get().strip()
    number = entry_number.get().strip()
    date = entry_date.get_date().strftime("%d-%m-%Y")
    expiry = entry_expiry.get_date().strftime("%d-%m-%Y") if entry_expiry.get_date() else ""
    issuer = entry_issuer.get().strip()
    category = entry_category.get().strip()
    tags = entry_tags.get().strip()

    if not name or not number or not date or not issuer:
        messagebox.showwarning("حقول مطلوبة", "يرجى تعبئة جميع الحقول: الاسم، الرقم، تاريخ الإصدار، الجهة المصدرة.")
        set_status("فشل التعديل: حقول مطلوبة مفقودة.")
        return

    doc_id = doc_table.item(selected[0])['values'][0]
    try:
        update_document(doc_id, name, number, date, expiry, issuer, None, category, tags)
        messagebox.showinfo("نجاح", "تم تعديل المستند بنجاح.")
        clear_fields()
        load_documents()
        update_category_filter_options()
        set_status("تم تعديل المستند بنجاح.")
    except ValueError as e:
        messagebox.showerror("خطأ في الإدخال", str(e))
        set_status(f"خطأ في الإدخال: {e}")
    except Exception as e:
        messagebox.showerror("خطأ", f"حدث خطأ غير متوقع: {e}")
        set_status(f"خطأ: {e}")

def delete_selected_document():
    """حذف المستند المحدد من قاعدة البيانات."""
    selected = doc_table.selection()
    if not selected:
        messagebox.showwarning("تحذير", "يرجى تحديد مستند للحذف.")
        return
    
    dialog = CustomConfirmDialog(root, "تأكيد الحذف", "هل أنت متأكد أنك تريد حذف هذا المستند وكل مرفقاته؟")
    if dialog.result:
        doc_id = doc_table.item(selected[0])['values'][0]
        set_status(f"جاري حذف المستند ID: {doc_id}...")
        try:
            delete_document(doc_id)
            load_documents()
            update_category_filter_options()
            messagebox.showinfo("نجاح", "تم حذف المستند بنجاح.")
            clear_fields()
            set_status("تم حذف المستند بنجاح.")
        except Exception as e:
            messagebox.showerror("خطأ", f"حدث خطأ أثناء حذف المستند: {e}")
            set_status(f"خطأ في الحذف: {e}")

def search_documents():
    """البحث عن المستندات وتصفيتها وعرضها في الجدول."""
    keyword = search_var.get()
    filter_status = filter_var.get()
    selected_category = category_filter_var.get()

    doc_table.delete(*doc_table.get_children())
    attachments_table.delete(*attachments_table.get_children())
    set_status("جاري البحث عن المستندات...")

    try:
        with sqlite3.connect("document_management.db") as conn:
            cursor = conn.cursor()
            
            query = """
                SELECT id, name, number, date, expiry_date, issuer, category, tags
                FROM documents
                WHERE (name LIKE ? OR number LIKE ? OR issuer LIKE ? OR category LIKE ? OR tags LIKE ?)
            """
            params = (f"%{keyword}%", f"%{keyword}%", f"%{keyword}%", f"%{keyword}%", f"%{keyword}%")

            if selected_category != "الكل":
                query += " AND category = ?"
                params += (selected_category,)

            cursor.execute(query, params)

            results_count = 0
            for row in cursor.fetchall():
                color = get_row_color(row[4])
                if filter_status == "الكل" or \
                   (filter_status == "صالحة" and color == "valid") or \
                   (filter_status == "قرب الانتهاء" and color == "near") or \
                   (filter_status == "منتهية" and color == "expired"):
                    doc_table.insert("", "end", values=row, tags=(color,))
                    results_count += 1

            set_status(f"تم العثور على {results_count} مستند/ات.")

    except Exception as e:
        messagebox.showerror("خطأ في البحث", f"حدث خطأ أثناء البحث عن المستندات: {e}")
        set_status(f"خطأ في البحث: {e}")

def load_documents():
    """تحميل جميع المستندات أو المستندات بناءً على البحث/التصفية."""
    search_documents()

def update_category_filter_options():
    """تحديث خيارات تصفية الفئات."""
    categories = get_all_categories()
    category_filter_menu['values'] = ["الكل"] + categories

# --- دوال إدارة المرفقات في الواجهة ---
def load_attachments(document_id):
    """تحميل وعرض المرفقات لمستند معين."""
    attachments_table.delete(*attachments_table.get_children())
    try:
        attachments = get_attachments_for_document(document_id)
        for att in attachments:
            attachments_table.insert("", "end", values=att)
    except Exception as e:
        messagebox.showerror("خطأ", f"فشل تحميل المرفقات: {e}")
        set_status(f"خطأ في تحميل المرفقات: {e}")

def add_attachment_to_selected():
    """إرفاق ملف بالمستند المحدد."""
    selected = doc_table.selection()
    if not selected:
        messagebox.showwarning("تحذير", "يرجى تحديد مستند أولاً لإرفاق ملف به.")
        return
    
    doc_id = doc_table.item(selected[0])['values'][0]
    file_path = filedialog.askopenfilename(
        title="اختر ملفاً لإرفاقه",
        filetypes=(("جميع الملفات", "*.*"), ("ملفات PDF", "*.pdf"), ("مستندات Word", "*.doc *.docx"), ("صور", "*.png *.jpg *.jpeg"))
    )
    if file_path:
        set_status(f"جاري إرفاق الملف: {os.path.basename(file_path)}...")
        try:
            add_attachment(doc_id, file_path)
            load_attachments(doc_id)
            messagebox.showinfo("نجاح", "تم إرفاق الملف بنجاح.")
            set_status("تم إرفاق الملف بنجاح.")
        except Exception as e:
            messagebox.showerror("خطأ", f"فشل إرفاق الملف: {e}")
            set_status(f"فشل إرفاق الملف: {e}")

def open_selected_attachment():
    """فتح المرفق المحدد."""
    selected = attachments_table.selection()
    if not selected:
        messagebox.showwarning("تحذير", "يرجى تحديد مرفق لفتحه.")
        return
    
    values = attachments_table.item(selected[0])['values']
    filepath = values[2]

    if os.path.exists(filepath):
        try:
            if os.name == 'nt':
                os.startfile(filepath)
            elif os.uname().sysname == 'Darwin':
                subprocess.call(('open', filepath))
            else:
                subprocess.call(('xdg-open', filepath))
            set_status(f"تم فتح الملف: {os.path.basename(filepath)}.")
        except Exception as e:
            messagebox.showerror("خطأ", f"فشل فتح الملف: {e}")
            set_status(f"فشل فتح الملف: {e}")
    else:
        messagebox.showerror("خطأ", "الملف غير موجود في المسار المحدد.")
        set_status("الملف غير موجود.")

def delete_selected_attachment():
    """حذف المرفق المحدد."""
    selected = attachments_table.selection()
    if not selected:
        messagebox.showwarning("تحذير", "يرجى تحديد مرفق لحذفه.")
        return
    
    dialog = CustomConfirmDialog(root, "تأكيد الحذف", "هل أنت متأكد أنك تريد حذف هذا المرفق؟")
    if dialog.result:
        values = attachments_table.item(selected[0])['values']
        attachment_id = values[0]
        filepath = values[2]
        selected_doc_item = doc_table.selection()
        if selected_doc_item:
            doc_id = doc_table.item(selected_doc_item[0])['values'][0]
        else:
            doc_id = None

        set_status(f"جاري حذف المرفق: {values[1]}...")
        try:
            delete_attachment(attachment_id, filepath)
            if doc_id:
                load_attachments(doc_id)
            else:
                attachments_table.delete(*attachments_table.get_children())
            messagebox.showinfo("نجاح", "تم حذف المرفق بنجاح.")
            set_status("تم حذف المرفق بنجاح.")
        except Exception as e:
            messagebox.showerror("خطأ", f"فشل حذف المرفق: {e}")
            set_status(f"فشل حذف المرفق: {e}")

def delete_all_attachments_for_document():
    """حذف جميع المرفقات للمستند المحدد."""
    selected_doc_item = doc_table.selection()
    if not selected_doc_item:
        messagebox.showwarning("تحذير", "يرجى تحديد مستند أولاً لحذف مرفقاته.")
        return
    
    doc_id = doc_table.item(selected_doc_item[0])['values'][0]
    doc_name = doc_table.item(selected_doc_item[0])['values'][1]

    dialog = CustomConfirmDialog(root, "تأكيد الحذف", f"هل أنت متأكد أنك تريد حذف جميع المرفقات للمستند: {doc_name}؟")
    if dialog.result:
        set_status(f"جاري حذف جميع المرفقات للمستند ID: {doc_id}...")
        try:
            attachments = get_attachments_for_document(doc_id)
            if not attachments:
                messagebox.showinfo("معلومات", "لا توجد مرفقات لحذفها لهذا المستند.")
                set_status("لا توجد مرفقات لحذفها.")
                return

            for att_id, att_filename, att_filepath, _ in attachments:
                delete_attachment(att_id, att_filepath)
            
            load_attachments(doc_id)
            messagebox.showinfo("نجاح", f"تم حذف جميع المرفقات للمستند: {doc_name} بنجاح.")
            log_audit_event("حذف جميع المرفقات", f"تم حذف جميع المرفقات للمستند ID: {doc_id} ({doc_name})")
            set_status("تم حذف جميع المرفقات بنجاح.")
        except Exception as e:
            messagebox.showerror("خطأ", f"فشل حذف جميع المرفقات: {e}")
            set_status(f"فشل حذف جميع المرفقات: {e}")

# --- دوال الموظفين ---
def clear_employee_fields():
    """مسح جميع حقول إدخال الموظف."""
    for entry in emp_entries:
        if isinstance(entry, DateEntry):
            entry.set_date(datetime.now().date())
        else:
            entry.delete(0, tk.END)
    set_status("تم مسح حقول الموظف.")

def load_employees():
    """تحميل وعرض بيانات الموظفين في الجدول."""
    emp_table.delete(*emp_table.get_children())
    set_status("جاري تحميل بيانات الموظفين...")
    try:
        employees = fetch_all_employees()
        for emp in employees:
            formatted_hire_date = convert_date_from_db_format(emp[5])
            emp_table.insert("", "end", values=(emp[0], emp[1], emp[2], emp[3], emp[4], formatted_hire_date))
        set_status(f"تم تحميل {len(employees)} موظف/موظفين.")
    except Exception as e:
        messagebox.showerror("خطأ", f"حدث خطأ أثناء تحميل بيانات الموظفين: {e}")
        set_status(f"خطأ في تحميل الموظفين: {e}")

def save_employee():
    """حفظ بيانات موظف جديد في قاعدة البيانات."""
    set_status("جاري حفظ بيانات الموظف...")
    try:
        name = emp_entry_name.get().strip()
        number = emp_entry_number.get().strip()
        department = emp_entry_department.get().strip()
        contact = emp_entry_contact.get().strip()
        hire_date = emp_entry_hire_date.get_date().strftime("%d-%m-%Y")

        add_employee(name, number, department, contact, hire_date)
        messagebox.showinfo("نجاح", "تمت إضافة الموظف بنجاح.")
        clear_employee_fields()
        load_employees()
        set_status("تم حفظ الموظف بنجاح.")
    except ValueError as e:
        messagebox.showerror("خطأ في الإدخال", str(e))
        set_status(f"خطأ في الإدخال: {e}")
    except Exception as e:
        messagebox.showerror("خطأ", f"حدث خطأ غير متوقع: {e}")
        set_status(f"خطأ: {e}")

def populate_employee_form_from_selection():
    """ملء حقول إدخال الموظف بمعلومات الموظف المحدد في الجدول."""
    selected = emp_table.selection()
    if not selected:
        clear_employee_fields()
        return
    values = emp_table.item(selected[0])['values']
    clear_employee_fields()

    emp_entry_name.insert(0, values[1])
    emp_entry_number.insert(0, values[2])
    emp_entry_department.insert(0, values[3])
    emp_entry_contact.insert(0, values[4])
    
    hire_date_ddmmyyyy = values[5]
    if hire_date_ddmmyyyy:
        emp_entry_hire_date.set_date(datetime.strptime(hire_date_ddmmyyyy, "%d-%m-%Y").date())
    else:
        emp_entry_hire_date.set_date("")
    set_status(f"تم تحديد الموظف: {values[1]}.")

def update_selected_employee():
    """تحديث بيانات الموظف المحدد في قاعدة البيانات."""
    selected = emp_table.selection()
    if not selected:
        messagebox.showwarning("تحذير", "يرجى تحديد موظف للتعديل.")
        return
    
    set_status("جاري تعديل بيانات الموظف...")
    emp_id = emp_table.item(selected[0])['values'][0]
    name = emp_entry_name.get().strip()
    number = emp_entry_number.get().strip()
    department = emp_entry_department.get().strip()
    contact = emp_entry_contact.get().strip()
    hire_date = emp_entry_hire_date.get_date().strftime("%d-%m-%Y")

    try:
        update_employee(emp_id, name, number, department, contact, hire_date)
        messagebox.showinfo("نجاح", "تم تعديل بيانات الموظف بنجاح.")
        clear_employee_fields()
        load_employees()
        set_status("تم تعديل الموظف بنجاح.")
    except ValueError as e:
        messagebox.showerror("خطأ في الإدخال", str(e))
        set_status(f"خطأ في الإدخال: {e}")
    except Exception as e:
        messagebox.showerror("خط", f"حدث خطأ غير متوقع: {e}")
        set_status(f"خطأ: {e}")

def delete_selected_employee():
    """حذف الموظف المحدد من قاعدة البيانات."""
    selected = emp_table.selection()
    if not selected:
        messagebox.showwarning("تحذير", "يرجى تحديد موظف للحذف.")
        return
    
    dialog = CustomConfirmDialog(root, "تأكيد الحذف", "هل أنت متأكد أنك تريد حذف هذا الموظف؟\n(ملاحظة: سيتم إزالة ربط هذا الموظف بأي مستندات.)")
    if dialog.result:
        emp_id = emp_table.item(selected[0])['values'][0]
        set_status(f"جاري حذف الموظف ID: {emp_id}...")
        try:
            delete_employee(emp_id)
            load_employees()
            messagebox.showinfo("نجاح", "تم حذف الموظف بنجاح.")
            clear_employee_fields()
            set_status("تم حذف الموظف بنجاح.")
        except Exception as e:
            messagebox.showerror("خطأ", f"حدث خطأ أثناء حذف الموظف: {e}")
            set_status(f"خطأ في الحذف: {e}")

# --- دوال سجل التدقيق ---
def load_audit_log():
    """تحميل وعرض سجل التدقيق في الجدول."""
    audit_table.delete(*audit_table.get_children())
    set_status("جاري تحميل سجل التدقيق...")
    try:
        logs = fetch_audit_log()
        for log in logs:
            audit_table.insert("", "end", values=log)
        set_status(f"تم تحميل {len(logs)} سجل/سجلات تدقيق.")
    except Exception as e:
        messagebox.showerror("خطأ", f"حدث خطأ أثناء تحميل سجل التدقيق: {e}")
        set_status(f"خطأ في تحميل سجل التدقيق: {e}")

# --- دوال المدة المتبقية ---
def load_remaining_time_documents():
    """تحميل وعرض معلومات المدة المتبقية للمستندات في الجدول."""
    remaining_time_table.delete(*remaining_time_table.get_children())
    set_status("جاري تحميل معلومات المدة المتبقية للمستندات...")
    try:
        with sqlite3.connect("document_management.db") as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT id, name, number, expiry_date FROM documents WHERE expiry_date IS NOT NULL")
            
            for row in cursor.fetchall():
                doc_id, name, number, expiry_date_db = row
                
                remaining_time_str = calculate_remaining_time(expiry_date_db)
                
                display_expiry_date = convert_date_from_db_format(expiry_date_db)

                remaining_time_table.insert("", "end", values=(doc_id, name, number, display_expiry_date, remaining_time_str))
            
            set_status(f"تم تحميل معلومات المدة المتبقية لـ {len(remaining_time_table.get_children())} مستند/ات.")
    except Exception as e:
        messagebox.showerror("خطأ", f"حدث خطأ أثناء تحميل المدة المتبقية للمستندات: {e}")
        set_status(f"خطأ في تحميل المدة المتبقية: {e}")

# --- دالة التصدير للمستندات ---
def export_documents_to_excel():
    """تصدير جميع بيانات المستندات إلى ملف Excel."""
    filepath = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        title="حفظ المستندات كملف Excel"
    )
    if not filepath:
        set_status("تم إلغاء عملية التصدير.")
        return

    set_status("جاري تصدير المستندات إلى Excel...")
    try:
        documents_data = fetch_all_documents_for_export()
        
        formatted_documents_data = []
        for doc in documents_data:
            doc_list = list(doc)
            if doc_list[4]:
                doc_list[4] = convert_date_from_db_format(doc_list[4])
            formatted_documents_data.append(doc_list)

        df = pd.DataFrame(formatted_documents_data, columns=[
            "ID", "الاسم", "الرقم", "تاريخ الإصدار", "تاريخ الانتهاء", "الجهة المصدرة", "الفئة", "العلامات"
        ])
        df.to_excel(filepath, index=False)
        messagebox.showinfo("نجاح", f"تم تصدير المستندات بنجاح إلى:\n{filepath}")
        log_audit_event("تصدير بيانات", f"تم تصدير جميع المستندات إلى ملف Excel: {filepath}")
        set_status(f"تم تصدير المستندات بنجاح إلى: {filepath}")
    except Exception as e:
        messagebox.showerror("خطأ في التصدير", f"حدث خطأ أثناء تصدير المستندات: {e}")
        set_status(f"خطأ في التصدير: {e}")

# --- دوال الرواتب ---
def update_employee_salary_options():
    """تحديث خيارات الموظفين في قائمة الرواتب المنسدلة."""
    employees = fetch_employee_id_name()
    employee_options = [f"{emp[1]} (ID: {emp[0]})" for emp in employees]
    emp_id_salary_combobox['values'] = employee_options

def update_department_salary_filter_options():
    """تحديث خيارات تصفية الأقسام في قائمة الرواتب المنسدلة."""
    departments = get_all_departments()
    department_salary_filter_combobox['values'] = ["الكل"] + departments

def update_salary_display_fields(event=None):
    """
    يقوم بتحديث حقول الراتب الشهري والسنوي وصافي الراتب
    بناءً على الحقل الذي تم تعديله.
    """
    focused_widget = root.focus_get()
    
    try:
        # تحديد الحقل الذي تم تعديله
        if focused_widget == entry_monthly_basic_salary:
            monthly_val = float(monthly_basic_salary_var.get() or 0)
            annual_val = monthly_val * 12
            annual_basic_salary_var.set(f"{annual_val:.2f}")
        elif focused_widget == entry_annual_basic_salary:
            annual_val = float(annual_basic_salary_var.get() or 0)
            monthly_val = annual_val / 12
            monthly_basic_salary_var.set(f"{monthly_val:.2f}")
        else:
            # إذا لم يكن أي من حقول الراتب الأساسي، فقط احسب الصافي
            monthly_val = float(monthly_basic_salary_var.get() or 0)

        # حساب وعرض صافي الراتب
        allow = float(entry_allowances.get() or 0)
        deduct = float(entry_deductions.get() or 0)
        net = calculate_net_salary(monthly_val, allow, deduct)
        label_net_salary_value.config(text=f"{net:.2f}")
    except ValueError:
        # إذا كان الإدخال غير رقمي، قم بمسح الحقول ذات الصلة وعرض خطأ
        if focused_widget == entry_monthly_basic_salary:
            annual_basic_salary_var.set("")
        elif focused_widget == entry_annual_basic_salary:
            monthly_basic_salary_var.set("")
        label_net_salary_value.config(text="خطأ")
        set_status("خطأ: يرجى إدخال قيم رقمية للراتب والبدلات والخصومات.")
    except Exception as e:
        print(f"Error in update_salary_display_fields: {e}")
        label_net_salary_value.config(text="خطأ")
        set_status(f"خطأ غير متوقع في حساب الراتب: {e}")

def clear_salary_fields():
    """مسح جميع حقول إدخال الرواتب."""
    emp_id_salary_combobox.set("")
    monthly_basic_salary_var.set("")
    annual_basic_salary_var.set("")
    entry_allowances.delete(0, tk.END)
    entry_deductions.delete(0, tk.END)
    label_net_salary_value.config(text="0.00")
    payment_method_var.set("تحويل بنكي")
    entry_payment_date.set_date(datetime.now().date())
    set_status("تم مسح حقول الرواتب.")

def save_salary():
    """حفظ سجل راتب جديد."""
    set_status("جاري حفظ الراتب...")
    try:
        selected_emp_str = emp_id_salary_combobox.get()
        if not selected_emp_str:
            messagebox.showwarning("تحذير", "يرجى تحديد موظف.")
            return
        emp_id = int(selected_emp_str.split('(ID: ')[1][:-1])

        # نأخذ الراتب الشهري من الحقل المخصص له
        basic_salary_monthly = float(monthly_basic_salary_var.get() or 0)
        allowances = float(entry_allowances.get() or 0)
        deductions = float(entry_deductions.get() or 0)
        payment_method = payment_method_var.get()
        payment_date = entry_payment_date.get_date().strftime("%d-%m-%Y")

        add_salary(emp_id, basic_salary_monthly, allowances, deductions, payment_method, payment_date)
        messagebox.showinfo("نجاح", "تم حفظ الراتب بنجاح.")
        clear_salary_fields()
        load_salaries()
        set_status("تم حفظ الراتب بنجاح.")
    except ValueError as e:
        messagebox.showerror("خطأ في الإدخال", str(e))
        set_status(f"خطأ في الإدخال: {e}")
    except Exception as e:
        messagebox.showerror("خطأ", f"حدث خطأ غير متوقع أثناء حفظ الراتب: {e}")
        set_status(f"خطأ: {e}")

def populate_salary_form_from_selection():
    """ملء حقول إدخال الرواتب بمعلومات الراتب المحدد في الجدول."""
    selected = salary_table.selection()
    if not selected:
        clear_salary_fields()
        return
    values = salary_table.item(selected[0])['values']
    # القيم هي: id, employee_name, department, basic_salary_monthly, basic_salary_annual, allowances, deductions, net_salary, payment_method, payment_date, employee_id
    salary_id = values[0]
    employee_name = values[1]
    department = values[2] # القيمة الجديدة
    monthly_basic_salary = values[3]
    annual_basic_salary = values[4]
    allowances = values[5]
    deductions = values[6]
    net_salary = values[7]
    payment_method = values[8]
    payment_date_ddmmyyyy = values[9] # تاريخ الدفع هو الآن العنصر العاشر (الفهرس 9)
    employee_id_from_db = values[10] # employee_id هو الآن العنصر الحادي عشر (الفهرس 10)

    clear_salary_fields()
    emp_id_salary_combobox.set(f"{employee_name} (ID: {employee_id_from_db})")
    monthly_basic_salary_var.set(str(monthly_basic_salary))
    annual_basic_salary_var.set(str(annual_basic_salary))
    entry_allowances.insert(0, str(allowances))
    entry_deductions.insert(0, str(deductions))
    label_net_salary_value.config(text=f"{net_salary:.2f}")
    payment_method_var.set(payment_method)
    entry_payment_date.set_date(datetime.strptime(payment_date_ddmmyyyy, "%d-%m-%Y").date())
    salary_table.current_salary_id = salary_id
    set_status(f"تم تحديد سجل الراتب للموظف: {employee_name}.")

def update_selected_salary():
    """تحديث سجل الراتب المحدد."""
    selected = salary_table.selection()
    if not selected:
        messagebox.showwarning("تحذير", "يرجى تحديد سجل راتب للتعديل.")
        return
    
    salary_id = salary_table.current_salary_id
    if not salary_id:
        messagebox.showwarning("خطأ", "لم يتم تحديد راتب صالح للتعديل.")
        return

    set_status("جاري تعديل سجل الراتب...")
    try:
        selected_emp_str = emp_id_salary_combobox.get()
        if not selected_emp_str:
            messagebox.showwarning("تحذير", "يرجى تحديد موظف.")
            return
        emp_id = int(selected_emp_str.split('(ID: ')[1][:-1])
        
        basic_salary_monthly = float(monthly_basic_salary_var.get() or 0)
        allowances = float(entry_allowances.get() or 0)
        deductions = float(entry_deductions.get() or 0)
        payment_method = payment_method_var.get()
        payment_date = entry_payment_date.get_date().strftime("%d-%m-%Y")

        update_salary(salary_id, emp_id, basic_salary_monthly, allowances, deductions, payment_method, payment_date)
        messagebox.showinfo("نجاح", "تم تعديل سجل الراتب بنجاح.")
        clear_salary_fields()
        load_salaries()
        set_status("تم تعديل سجل الراتب بنجاح.")
    except ValueError as e:
        messagebox.showerror("خطأ في الإدخال", str(e))
        set_status(f"خطأ في الإدخال: {e}")
    except Exception as e:
        messagebox.showerror("خطأ", f"حدث خطأ غير متوقع أثناء تعديل الراتب: {e}")
        set_status(f"خطأ: {e}")

def delete_selected_salary():
    """حذف سجل الراتب المحدد."""
    selected = salary_table.selection()
    if not selected:
        messagebox.showwarning("تحذير", "يرجى تحديد سجل راتب للحذف.")
        return
    
    dialog = CustomConfirmDialog(root, "تأكيد الحذف", "هل أنت متأكد أنك تريد حذف سجل الراتب هذا؟")
    if dialog.result:
        salary_id = salary_table.item(selected[0])['values'][0]
        set_status(f"جاري حذف سجل الراتب ID: {salary_id}...")
        try:
            delete_salary(salary_id)
            load_salaries()
            messagebox.showinfo("نجاح", "تم حذف سجل الراتب بنجاح.")
            clear_salary_fields()
            set_status("تم حذف سجل الراتب بنجاح.")
        except Exception as e:
            messagebox.showerror("خطأ", f"حدث خطأ أثناء حذف سجل الراتب: {e}")
            set_status(f"خطأ في الحذف: {e}")

def search_salaries():
    """البحث وتصفية سجلات الرواتب وعرضها في الجدول."""
    keyword = salary_search_var.get()
    selected_department = department_salary_filter_var.get()
    selected_month = salary_month_filter_var.get()
    selected_year = salary_year_filter_var.get()

    salary_table.delete(*salary_table.get_children())
    set_status("جاري البحث عن سجلات الرواتب...")

    try:
        salaries = fetch_all_salaries()
        results_count = 0
        for s in salaries:
            salary_id, employee_name, department, basic_salary_monthly, allowances, deductions, net_salary, payment_method, payment_date_db, employee_id = s
            
            # Convert date for display and filtering
            payment_date_ddmmyyyy = convert_date_from_db_format(payment_date_db)
            payment_date_obj = datetime.strptime(payment_date_ddmmyyyy, "%d-%m-%Y").date()

            # Calculate annual basic salary for display
            basic_salary_annual = basic_salary_monthly * 12

            # Apply filters
            keyword_match = (keyword.lower() in employee_name.lower() or
                             keyword.lower() in department.lower() or
                             keyword.lower() in str(basic_salary_monthly) or
                             keyword.lower() in str(net_salary))

            department_match = (selected_department == "الكل" or department == selected_department)
            month_match = (selected_month == "الكل" or payment_date_obj.month == int(selected_month))
            year_match = (selected_year == "الكل" or payment_date_obj.year == int(selected_year))

            if keyword_match and department_match and month_match and year_match:
                salary_table.insert("", "end", values=(
                    salary_id, employee_name, department, 
                    f"{basic_salary_monthly:.2f}", f"{basic_salary_annual:.2f}", 
                    f"{allowances:.2f}", f"{deductions:.2f}", f"{net_salary:.2f}", 
                    payment_method, payment_date_ddmmyyyy, employee_id
                ))
                results_count += 1
        set_status(f"تم العثور على {results_count} سجل/سجلات راتب.")

    except Exception as e:
        messagebox.showerror("خطأ في البحث", f"حدث خطأ أثناء البحث عن سجلات الرواتب: {e}")
        set_status(f"خطأ في البحث: {e}")

def load_salaries():
    """تحميل جميع سجلات الرواتب أو السجلات بناءً على البحث/التصفية."""
    search_salaries()

def export_salaries_to_excel():
    """تصدير جميع بيانات الرواتب إلى ملف Excel."""
    filepath = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        title="حفظ الرواتب كملف Excel"
    )
    if not filepath:
        set_status("تم إلغاء عملية التصدير.")
        return

    set_status("جاري تصدير الرواتب إلى Excel...")
    try:
        salaries_data = fetch_all_salaries_for_export()
        
        formatted_salaries_data = []
        for sal in salaries_data:
            sal_list = list(sal)
            sal_list[8] = convert_date_from_db_format(sal_list[8]) # Format payment_date
            sal_list.insert(4, f"{sal_list[3] * 12:.2f}") # Insert annual salary after monthly
            formatted_salaries_data.append(sal_list)

        df = pd.DataFrame(formatted_salaries_data, columns=[
            "ID", "اسم الموظف", "القسم", "الراتب الأساسي (شهري)", "الراتب الأساسي (سنوي)",
            "البدلات", "الخصومات", "صافي الراتب", "طريقة الدفع", "تاريخ الدفع", "معرف الموظف"
        ])
        df.to_excel(filepath, index=False)
        messagebox.showinfo("نجاح", f"تم تصدير الرواتب بنجاح إلى:\n{filepath}")
        log_audit_event("تصدير بيانات", f"تم تصدير جميع الرواتب إلى ملف Excel: {filepath}")
        set_status(f"تم تصدير الرواتب بنجاح إلى: {filepath}")
    except Exception as e:
        messagebox.showerror("خطأ في التصدير", f"حدث خطأ أثناء تصدير الرواتب: {e}")
        set_status(f"خطأ في التصدير: {e}")

# --- الألسنة (Notebook) ---
notebook = ttk.Notebook(root)
notebook.pack(pady=10, expand=True, fill="both")

# --- تبويب المستندات ---
documents_tab = ttk.Frame(notebook)
notebook.add(documents_tab, text="إدارة المستندات")

# إطارات المستندات
doc_inputs_frame = ttk.LabelFrame(documents_tab, text="إدخال بيانات المستند")
doc_inputs_frame.pack(pady=10, padx=10, fill="x")

doc_table_frame = ttk.LabelFrame(documents_tab, text="المستندات")
doc_table_frame.pack(pady=10, padx=10, expand=True, fill="both")

# حقول إدخال المستندات
labels_docs = ["اسم المستند:", "رقم المستند:", "تاريخ الإصدار:", "تاريخ الانتهاء:", "الجهة المصدرة:", "الفئة:", "العلامات:"]
entries = []

tk.Label(doc_inputs_frame, text=labels_docs[0]).grid(row=0, column=0, padx=5, pady=5, sticky="w")
entry_name = ttk.Entry(doc_inputs_frame, width=30)
entry_name.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
entries.append(entry_name)

tk.Label(doc_inputs_frame, text=labels_docs[1]).grid(row=0, column=2, padx=5, pady=5, sticky="w")
entry_number = ttk.Entry(doc_inputs_frame, width=30)
entry_number.grid(row=0, column=3, padx=5, pady=5, sticky="ew")
entries.append(entry_number)

tk.Label(doc_inputs_frame, text=labels_docs[2]).grid(row=1, column=0, padx=5, pady=5, sticky="w")
entry_date = DateEntry(doc_inputs_frame, width=12, background='darkblue', foreground='white', borderwidth=2, date_pattern='dd-mm-yyyy', locale='ar')
entry_date.grid(row=1, column=1, padx=5, pady=5, sticky="ew")
entries.append(entry_date)

tk.Label(doc_inputs_frame, text=labels_docs[3]).grid(row=1, column=2, padx=5, pady=5, sticky="w")
entry_expiry = DateEntry(doc_inputs_frame, width=12, background='darkblue', foreground='white', borderwidth=2, date_pattern='dd-mm-yyyy', locale='ar')
entry_expiry.grid(row=1, column=3, padx=5, pady=5, sticky="ew")
entries.append(entry_expiry)

tk.Label(doc_inputs_frame, text=labels_docs[4]).grid(row=2, column=0, padx=5, pady=5, sticky="w")
entry_issuer = ttk.Entry(doc_inputs_frame, width=30)
entry_issuer.grid(row=2, column=1, padx=5, pady=5, sticky="ew")
entries.append(entry_issuer)

tk.Label(doc_inputs_frame, text=labels_docs[5]).grid(row=2, column=2, padx=5, pady=5, sticky="w")
entry_category = ttk.Entry(doc_inputs_frame, width=30)
entry_category.grid(row=2, column=3, padx=5, pady=5, sticky="ew")
entries.append(entry_category)

tk.Label(doc_inputs_frame, text=labels_docs[6]).grid(row=3, column=0, padx=5, pady=5, sticky="w")
entry_tags = ttk.Entry(doc_inputs_frame, width=30)
entry_tags.grid(row=3, column=1, padx=5, pady=5, sticky="ew")
entries.append(entry_tags)

doc_inputs_frame.grid_columnconfigure((1, 3), weight=1)

# أزرار الإجراءات للمستندات
doc_buttons_frame = ttk.Frame(doc_inputs_frame)
doc_buttons_frame.grid(row=4, column=0, columnspan=4, pady=10)

ttk.Button(doc_buttons_frame, text="حفظ مستند", command=save_document).pack(side=tk.LEFT, padx=5)
ttk.Button(doc_buttons_frame, text="مسح الحقول", command=clear_fields).pack(side=tk.LEFT, padx=5)
ttk.Button(doc_buttons_frame, text="تعديل المستند", command=update_selected_document).pack(side=tk.LEFT, padx=5)
ttk.Button(doc_buttons_frame, text="حذف المستند", command=delete_selected_document).pack(side=tk.LEFT, padx=5)
ttk.Button(doc_buttons_frame, text="تصدير إلى Excel", command=export_documents_to_excel).pack(side=tk.LEFT, padx=5)


# شريط البحث والتصفية للمستندات
search_frame = ttk.Frame(doc_table_frame)
search_frame.pack(pady=5, fill="x")

search_var = tk.StringVar()
ttk.Entry(search_frame, textvariable=search_var, width=40).pack(side=tk.LEFT, padx=5, expand=True, fill="x")
ttk.Button(search_frame, text="بحث", command=load_documents).pack(side=tk.LEFT, padx=5)

filter_var = tk.StringVar(value="الكل")
ttk.OptionMenu(search_frame, filter_var, "الكل", "الكل", "صالحة", "قرب الانتهاء", "منتهية", command=lambda x: load_documents()).pack(side=tk.LEFT, padx=5)

category_filter_var = tk.StringVar(value="الكل")
category_filter_menu = ttk.OptionMenu(search_frame, category_filter_var, "الكل", *["الكل"], command=lambda x: load_documents())
category_filter_menu.pack(side=tk.LEFT, padx=5)


# جدول عرض المستندات
doc_table = ttk.Treeview(doc_table_frame, columns=("id", "name", "number", "date", "expiry", "issuer", "category", "tags"), show="headings")
doc_table.pack(fill="both", expand=True)

# أعمدة جدول المستندات
doc_table.heading("id", text="ID", command=lambda: treeview_sort_column(doc_table, "id", False))
doc_table.heading("name", text="اسم المستند", command=lambda: treeview_sort_column(doc_table, "name", False))
doc_table.heading("number", text="رقم المستند", command=lambda: treeview_sort_column(doc_table, "number", False))
doc_table.heading("date", text="تاريخ الإصدار", command=lambda: treeview_sort_column(doc_table, "date", False))
doc_table.heading("expiry", text="تاريخ الانتهاء", command=lambda: treeview_sort_column(doc_table, "expiry", False))
doc_table.heading("issuer", text="الجهة المصدرة", command=lambda: treeview_sort_column(doc_table, "issuer", False))
doc_table.heading("category", text="الفئة", command=lambda: treeview_sort_column(doc_table, "category", False))
doc_table.heading("tags", text="العلامات", command=lambda: treeview_sort_column(doc_table, "tags", False))

doc_table.column("id", width=30, stretch=tk.NO)
doc_table.column("name", width=150)
doc_table.column("number", width=100)
doc_table.column("date", width=100)
doc_table.column("expiry", width=100)
doc_table.column("issuer", width=100)
doc_table.column("category", width=80)
doc_table.column("tags", width=120)

doc_table.bind("<<TreeviewSelect>>", lambda event: populate_form_from_selection())

# تحديد ألوان الصفوف بناءً على حالة الانتهاء
doc_table.tag_configure("expired", background="#FFCCCC") # أحمر فاتح للمنتهي
doc_table.tag_configure("near", background="#FFFFCC")    # أصفر فاتح لقرب الانتهاء
doc_table.tag_configure("valid", background="#CCFFCC")   # أخضر فاتح للصالح

# شريط التمرير لجدول المستندات
doc_table_scrollbar_y = ttk.Scrollbar(doc_table_frame, orient="vertical", command=doc_table.yview)
doc_table_scrollbar_y.pack(side="right", fill="y")
doc_table.configure(yscrollcommand=doc_table_scrollbar_y.set)

doc_table_scrollbar_x = ttk.Scrollbar(doc_table_frame, orient="horizontal", command=doc_table.xview)
doc_table_scrollbar_x.pack(side="bottom", fill="x")
doc_table.configure(xscrollcommand=doc_table_scrollbar_x.set)


# --- إطار المرفقات ---
attachments_frame = ttk.LabelFrame(documents_tab, text="المرفقات")
attachments_frame.pack(pady=10, padx=10, fill="both", expand=True)

attachments_table = ttk.Treeview(attachments_frame, columns=("id", "filename", "filepath", "upload_time"), show="headings")
attachments_table.pack(fill="both", expand=True)

attachments_table.heading("id", text="ID")
attachments_table.heading("filename", text="اسم الملف")
attachments_table.heading("filepath", text="المسار")
attachments_table.heading("upload_time", text="وقت الرفع")

attachments_table.column("id", width=30)
attachments_table.column("filename", width=200)
attachments_table.column("filepath", width=300)
attachments_table.column("upload_time", width=150)

# أزرار المرفقات
attachment_buttons_frame = ttk.Frame(attachments_frame)
attachment_buttons_frame.pack(pady=5)

ttk.Button(attachment_buttons_frame, text="إضافة مرفق", command=add_attachment_to_selected).pack(side=tk.LEFT, padx=5)
ttk.Button(attachment_buttons_frame, text="فتح المرفق", command=open_selected_attachment).pack(side=tk.LEFT, padx=5)
ttk.Button(attachment_buttons_frame, text="حذف المرفق", command=delete_selected_attachment).pack(side=tk.LEFT, padx=5)
ttk.Button(attachment_buttons_frame, text="حذف كل المرفقات", command=delete_all_attachments_for_document).pack(side=tk.LEFT, padx=5)

# شريط التمرير لجدول المرفقات
attachments_table_scrollbar_y = ttk.Scrollbar(attachments_frame, orient="vertical", command=attachments_table.yview)
attachments_table_scrollbar_y.pack(side="right", fill="y")
attachments_table.configure(yscrollcommand=attachments_table_scrollbar_y.set)

attachments_table_scrollbar_x = ttk.Scrollbar(attachments_frame, orient="horizontal", command=attachments_table.xview)
attachments_table_scrollbar_x.pack(side="bottom", fill="x")
attachments_table.configure(xscrollcommand=attachments_table_scrollbar_x.set)


# --- تبويب الموظفين ---
employees_tab = ttk.Frame(notebook)
notebook.add(employees_tab, text="إدارة الموظفين")

# إطارات الموظفين
emp_inputs_frame = ttk.LabelFrame(employees_tab, text="إدخال بيانات الموظف")
emp_inputs_frame.pack(pady=10, padx=10, fill="x")

emp_table_frame = ttk.LabelFrame(employees_tab, text="الموظفون")
emp_table_frame.pack(pady=10, padx=10, expand=True, fill="both")

# حقول إدخال الموظفين
labels_emp = ["اسم الموظف:", "الرقم الوظيفي:", "القسم:", "معلومات الاتصال:", "تاريخ التعيين:"]
emp_entries = []

tk.Label(emp_inputs_frame, text=labels_emp[0]).grid(row=0, column=0, padx=5, pady=5, sticky="w")
emp_entry_name = ttk.Entry(emp_inputs_frame, width=30)
emp_entry_name.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
emp_entries.append(emp_entry_name)

tk.Label(emp_inputs_frame, text=labels_emp[1]).grid(row=0, column=2, padx=5, pady=5, sticky="w")
emp_entry_number = ttk.Entry(emp_inputs_frame, width=30)
emp_entry_number.grid(row=0, column=3, padx=5, pady=5, sticky="ew")
emp_entries.append(emp_entry_number)

tk.Label(emp_inputs_frame, text=labels_emp[2]).grid(row=1, column=0, padx=5, pady=5, sticky="w")
emp_entry_department = ttk.Entry(emp_inputs_frame, width=30)
emp_entry_department.grid(row=1, column=1, padx=5, pady=5, sticky="ew")
emp_entries.append(emp_entry_department)

tk.Label(emp_inputs_frame, text=labels_emp[3]).grid(row=1, column=2, padx=5, pady=5, sticky="w")
emp_entry_contact = ttk.Entry(emp_inputs_frame, width=30)
emp_entry_contact.grid(row=1, column=3, padx=5, pady=5, sticky="ew")
emp_entries.append(emp_entry_contact)

tk.Label(emp_inputs_frame, text=labels_emp[4]).grid(row=2, column=0, padx=5, pady=5, sticky="w")
emp_entry_hire_date = DateEntry(emp_inputs_frame, width=12, background='darkblue', foreground='white', borderwidth=2, date_pattern='dd-mm-yyyy', locale='ar')
emp_entry_hire_date.grid(row=2, column=1, padx=5, pady=5, sticky="ew")
emp_entries.append(emp_entry_hire_date)

emp_inputs_frame.grid_columnconfigure((1, 3), weight=1)

# أزرار الإجراءات للموظفين
emp_buttons_frame = ttk.Frame(emp_inputs_frame)
emp_buttons_frame.grid(row=3, column=0, columnspan=4, pady=10)

ttk.Button(emp_buttons_frame, text="حفظ موظف", command=save_employee).pack(side=tk.LEFT, padx=5)
ttk.Button(emp_buttons_frame, text="مسح الحقول", command=clear_employee_fields).pack(side=tk.LEFT, padx=5)
ttk.Button(emp_buttons_frame, text="تعديل موظف", command=update_selected_employee).pack(side=tk.LEFT, padx=5)
ttk.Button(emp_buttons_frame, text="حذف موظف", command=delete_selected_employee).pack(side=tk.LEFT, padx=5)

# جدول عرض الموظفين
emp_table = ttk.Treeview(emp_table_frame, columns=("id", "name", "number", "department", "contact", "hire_date"), show="headings")
emp_table.pack(fill="both", expand=True)

# أعمدة جدول الموظفين
emp_table.heading("id", text="ID", command=lambda: treeview_sort_column(emp_table, "id", False))
emp_table.heading("name", text="اسم الموظف", command=lambda: treeview_sort_column(emp_table, "name", False))
emp_table.heading("number", text="الرقم الوظيفي", command=lambda: treeview_sort_column(emp_table, "number", False))
emp_table.heading("department", text="القسم", command=lambda: treeview_sort_column(emp_table, "department", False))
emp_table.heading("contact", text="معلومات الاتصال", command=lambda: treeview_sort_column(emp_table, "contact", False))
emp_table.heading("hire_date", text="تاريخ التعيين", command=lambda: treeview_sort_column(emp_table, "hire_date", False))

emp_table.column("id", width=50)
emp_table.column("name", width=150)
emp_table.column("number", width=100)
emp_table.column("department", width=100)
emp_table.column("contact", width=150)
emp_table.column("hire_date", width=100)

emp_table.bind("<<TreeviewSelect>>", lambda event: populate_employee_form_from_selection())

# شريط التمرير لجدول الموظفين
emp_table_scrollbar_y = ttk.Scrollbar(emp_table_frame, orient="vertical", command=emp_table.yview)
emp_table_scrollbar_y.pack(side="right", fill="y")
emp_table.configure(yscrollcommand=emp_table_scrollbar_y.set)

emp_table_scrollbar_x = ttk.Scrollbar(emp_table_frame, orient="horizontal", command=emp_table.xview)
emp_table_scrollbar_x.pack(side="bottom", fill="x")
emp_table.configure(xscrollcommand=emp_table_scrollbar_x.set)

# --- تبويب سجل التدقيق ---
audit_tab = ttk.Frame(notebook)
notebook.add(audit_tab, text="سجل التدقيق")

audit_table_frame = ttk.LabelFrame(audit_tab, text="سجل جميع الإجراءات")
audit_table_frame.pack(pady=10, padx=10, expand=True, fill="both")

audit_table = ttk.Treeview(audit_table_frame, columns=("id", "timestamp", "event_type", "description"), show="headings")
audit_table.pack(fill="both", expand=True)

audit_table.heading("id", text="ID", command=lambda: treeview_sort_column(audit_table, "id", False))
audit_table.heading("timestamp", text="الوقت", command=lambda: treeview_sort_column(audit_table, "timestamp", False))
audit_table.heading("event_type", text="نوع الحدث", command=lambda: treeview_sort_column(audit_table, "event_type", False))
audit_table.heading("description", text="الوصف", command=lambda: treeview_sort_column(audit_table, "description", False))

audit_table.column("id", width=50)
audit_table.column("timestamp", width=150)
audit_table.column("event_type", width=100)
audit_table.column("description", width=400)

# شريط التمرير لجدول التدقيق
audit_table_scrollbar_y = ttk.Scrollbar(audit_table_frame, orient="vertical", command=audit_table.yview)
audit_table_scrollbar_y.pack(side="right", fill="y")
audit_table.configure(yscrollcommand=audit_table_scrollbar_y.set)

audit_table_scrollbar_x = ttk.Scrollbar(audit_table_frame, orient="horizontal", command=audit_table.xview)
audit_table_scrollbar_x.pack(side="bottom", fill="x")
audit_table.configure(xscrollcommand=audit_table_scrollbar_x.set)


# --- تبويب المدة المتبقية للمستندات ---
remaining_time_tab = ttk.Frame(notebook)
notebook.add(remaining_time_tab, text="مدة صلاحية المستندات")

remaining_time_frame = ttk.LabelFrame(remaining_time_tab, text="المستندات حسب المدة المتبقية")
remaining_time_frame.pack(pady=10, padx=10, expand=True, fill="both")

remaining_time_table = ttk.Treeview(remaining_time_frame, columns=("id", "name", "number", "expiry_date", "remaining_time"), show="headings")
remaining_time_table.pack(fill="both", expand=True)

remaining_time_table.heading("id", text="ID", command=lambda: treeview_sort_column(remaining_time_table, "id", False))
remaining_time_table.heading("name", text="اسم المستند", command=lambda: treeview_sort_column(remaining_time_table, "name", False))
remaining_time_table.heading("number", text="رقم المستند", command=lambda: treeview_sort_column(remaining_time_table, "number", False))
remaining_time_table.heading("expiry_date", text="تاريخ الانتهاء", command=lambda: treeview_sort_column(remaining_time_table, "expiry_date", False))
remaining_time_table.heading("remaining_time", text="المدة المتبقية", command=lambda: treeview_sort_column(remaining_time_table, "remaining_time", False))

remaining_time_table.column("id", width=50)
remaining_time_table.column("name", width=150)
remaining_time_table.column("number", width=100)
remaining_time_table.column("expiry_date", width=100)
remaining_time_table.column("remaining_time", width=150)

# شريط التمرير لجدول المدة المتبقية
remaining_time_table_scrollbar_y = ttk.Scrollbar(remaining_time_frame, orient="vertical", command=remaining_time_table.yview)
remaining_time_table_scrollbar_y.pack(side="right", fill="y")
remaining_time_table.configure(yscrollcommand=remaining_time_table_scrollbar_y.set)

remaining_time_table_scrollbar_x = ttk.Scrollbar(remaining_time_frame, orient="horizontal", command=remaining_time_table.xview)
remaining_time_table_scrollbar_x.pack(side="bottom", fill="x")
remaining_time_table.configure(xscrollcommand=remaining_time_table_scrollbar_x.set)

# --- تبويب إدارة الرواتب ---
salaries_tab = ttk.Frame(notebook)
notebook.add(salaries_tab, text="إدارة الرواتب")

# إطارات الرواتب
salary_inputs_frame = ttk.LabelFrame(salaries_tab, text="إدخال بيانات الراتب")
salary_inputs_frame.pack(pady=10, padx=10, fill="x")

salary_table_frame = ttk.LabelFrame(salaries_tab, text="سجلات الرواتب")
salary_table_frame.pack(pady=10, padx=10, expand=True, fill="both")

# حقول إدخال الرواتب
tk.Label(salary_inputs_frame, text="اسم الموظف:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
emp_id_salary_combobox = ttk.Combobox(salary_inputs_frame, width=40, state="readonly")
emp_id_salary_combobox.grid(row=0, column=1, padx=5, pady=5, sticky="ew")

tk.Label(salary_inputs_frame, text="الراتب الأساسي (شهري):").grid(row=1, column=0, padx=5, pady=5, sticky="w")
monthly_basic_salary_var = tk.StringVar()
entry_monthly_basic_salary = ttk.Entry(salary_inputs_frame, width=20, textvariable=monthly_basic_salary_var)
entry_monthly_basic_salary.grid(row=1, column=1, padx=5, pady=5, sticky="ew")
entry_monthly_basic_salary.bind("<FocusOut>", update_salary_display_fields)
entry_monthly_basic_salary.bind("<Return>", update_salary_display_fields)

tk.Label(salary_inputs_frame, text="الراتب الأساسي (سنوي):").grid(row=1, column=2, padx=5, pady=5, sticky="w")
annual_basic_salary_var = tk.StringVar()
entry_annual_basic_salary = ttk.Entry(salary_inputs_frame, width=20, textvariable=annual_basic_salary_var)
entry_annual_basic_salary.grid(row=1, column=3, padx=5, pady=5, sticky="ew")
entry_annual_basic_salary.bind("<FocusOut>", update_salary_display_fields)
entry_annual_basic_salary.bind("<Return>", update_salary_display_fields)

tk.Label(salary_inputs_frame, text="البدلات:").grid(row=2, column=0, padx=5, pady=5, sticky="w")
entry_allowances = ttk.Entry(salary_inputs_frame, width=20)
entry_allowances.grid(row=2, column=1, padx=5, pady=5, sticky="ew")
entry_allowances.bind("<FocusOut>", update_salary_display_fields)
entry_allowances.bind("<Return>", update_salary_display_fields)

tk.Label(salary_inputs_frame, text="الخصومات:").grid(row=2, column=2, padx=5, pady=5, sticky="w")
entry_deductions = ttk.Entry(salary_inputs_frame, width=20)
entry_deductions.grid(row=2, column=3, padx=5, pady=5, sticky="ew")
entry_deductions.bind("<FocusOut>", update_salary_display_fields)
entry_deductions.bind("<Return>", update_salary_display_fields)

tk.Label(salary_inputs_frame, text="صافي الراتب:").grid(row=3, column=0, padx=5, pady=5, sticky="w")
label_net_salary_value = ttk.Label(salary_inputs_frame, text="0.00", font=('Arial', 10, 'bold'))
label_net_salary_value.grid(row=3, column=1, padx=5, pady=5, sticky="w")

tk.Label(salary_inputs_frame, text="طريقة الدفع:").grid(row=3, column=2, padx=5, pady=5, sticky="w")
payment_method_var = tk.StringVar(value="تحويل بنكي")
ttk.OptionMenu(salary_inputs_frame, payment_method_var, "تحويل بنكي", "تحويل بنكي", "نقداً", "شيك").grid(row=3, column=3, padx=5, pady=5, sticky="ew")

tk.Label(salary_inputs_frame, text="تاريخ الدفع:").grid(row=4, column=0, padx=5, pady=5, sticky="w")
entry_payment_date = DateEntry(salary_inputs_frame, width=12, background='darkblue', foreground='white', borderwidth=2, date_pattern='dd-mm-yyyy', locale='ar')
entry_payment_date.grid(row=4, column=1, padx=5, pady=5, sticky="ew")

salary_inputs_frame.grid_columnconfigure((1, 3), weight=1)

# أزرار الإجراءات للرواتب
salary_buttons_frame = ttk.Frame(salary_inputs_frame)
salary_buttons_frame.grid(row=5, column=0, columnspan=4, pady=10)

ttk.Button(salary_buttons_frame, text="حفظ الراتب", command=save_salary).pack(side=tk.LEFT, padx=5)
ttk.Button(salary_buttons_frame, text="مسح الحقول", command=clear_salary_fields).pack(side=tk.LEFT, padx=5)
ttk.Button(salary_buttons_frame, text="تعديل الراتب", command=update_selected_salary).pack(side=tk.LEFT, padx=5)
ttk.Button(salary_buttons_frame, text="حذف الراتب", command=delete_selected_salary).pack(side=tk.LEFT, padx=5)
ttk.Button(salary_buttons_frame, text="تصدير إلى Excel", command=export_salaries_to_excel).pack(side=tk.LEFT, padx=5)

# شريط البحث والتصفية للرواتب
salary_search_frame = ttk.Frame(salary_table_frame)
salary_search_frame.pack(pady=5, fill="x")

salary_search_var = tk.StringVar()
ttk.Entry(salary_search_frame, textvariable=salary_search_var, width=40).pack(side=tk.LEFT, padx=5, expand=True, fill="x")
ttk.Button(salary_search_frame, text="بحث", command=load_salaries).pack(side=tk.LEFT, padx=5)

department_salary_filter_var = tk.StringVar(value="الكل")
department_salary_filter_combobox = ttk.OptionMenu(salary_search_frame, department_salary_filter_var, "الكل", *["الكل"], command=lambda x: load_salaries())
department_salary_filter_combobox.pack(side=tk.LEFT, padx=5)

salary_month_filter_var = tk.StringVar(value="الكل")
ttk.OptionMenu(salary_search_frame, salary_month_filter_var, "الكل", *["الكل"] + [str(i) for i in range(1, 13)], command=lambda x: load_salaries()).pack(side=tk.LEFT, padx=5)

salary_year_filter_var = tk.StringVar(value="الكل")
# Adjust years as needed, e.g., current year +/- a range
current_year = datetime.now().year
years_options = ["الكل"] + [str(i) for i in range(current_year - 5, current_year + 2)]
ttk.OptionMenu(salary_search_frame, salary_year_filter_var, "الكل", *years_options, command=lambda x: load_salaries()).pack(side=tk.LEFT, padx=5)


# جدول عرض الرواتب
salary_table = ttk.Treeview(salary_table_frame, columns=(
    "id", "اسم الموظف", "القسم", "الراتب الأساسي (شهري)", "الراتب الأساسي (سنوي)",
    "البدلات", "الخصومات", "صافي الراتب", "طريقة الدفع", "تاريخ الدفع", "employee_id"
), show="headings")

# إخفاء عمود employee_id
salary_table.column("employee_id", width=0, stretch=tk.NO)

salary_table.pack(fill="both", expand=True)

# أعمدة جدول الرواتب
for col in ["id", "اسم الموظف", "القسم", "الراتب الأساسي (شهري)", "الراتب الأساسي (سنوي)", "البدلات", "الخصومات", "صافي الراتب", "طريقة الدفع", "تاريخ الدفع"]:
    salary_table.heading(col, text=col, command=lambda _col=col: treeview_sort_column(salary_table, _col, False))
    salary_table.column(col, anchor="center")

salary_table.column("id", width=30)
salary_table.column("اسم الموظف", width=100)
salary_table.column("القسم", width=80)
salary_table.column("الراتب الأساسي (شهري)", width=100)
salary_table.column("الراتب الأساسي (سنوي)", width=100)
salary_table.column("البدلات", width=70)
salary_table.column("الخصومات", width=70)
salary_table.column("صافي الراتب", width=90)
salary_table.column("طريقة الدفع", width=90)
salary_table.column("تاريخ الدفع", width=90)

salary_table.pack(fill="both", expand=True)

# شريط التمرير لجدول الرواتب
salary_table_scrollbar_y = ttk.Scrollbar(salary_table_frame, orient="vertical", command=salary_table.yview)
salary_table_scrollbar_y.pack(side="right", fill="y")
salary_table.configure(yscrollcommand=salary_table_scrollbar_y.set)

salary_table_scrollbar_x = ttk.Scrollbar(salary_table_frame, orient="horizontal", command=salary_table.xview)
salary_table_scrollbar_x.pack(side="bottom", fill="x")
salary_table.configure(xscrollcommand=salary_table_scrollbar_x.set)

# ربط تحديد الصف في جدول الرواتب
salary_table.bind("<<TreeviewSelect>>", lambda event: populate_salary_form_from_selection())


# Initial loads
update_category_filter_options()
load_documents()
load_employees()
load_audit_log()
load_remaining_time_documents()
update_employee_salary_options()
update_department_salary_filter_options()
load_salaries()

root.mainloop()