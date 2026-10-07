import sqlite3
import tkinter as tk
from tkinter import ttk, messagebox

# مسار قاعدة البيانات المشتركة على الشبكة المحلية
DB_PATH = r"\\DESKTOP-CK5V925\HelpdeskShare\tickets.db"

# قاموس اللغات (Localization Dictionary)
TRANSLATIONS = {
    "ar": {
        "title": "📩 تقديم طلب دعم فني جديد",
        "emp_label": "اسم الموظف / القسم",
        "issue_label": "وصف المشكلة",
        "priority_label": "درجة الأولوية",
        "high": "عالية",
        "medium": "متوسطة",
        "low": "منخفضة",
        "submit_btn": "🚀 إرسال التذكرة",
        "lang_btn": "English 🌐",
        "warn_fill": "يرجى تعبئة جميع الحقول المطلوبة!",
        "success": "تم إرسال التذكرة إلى قسم الدعم الفني بنجاح!",
        "conn_error": "تعذر الاتصال بقاعدة البيانات على الشبكة:\n",
        "db_error": "حدث خطأ أثناء إرسال التذكرة:\n"
    },
    "en": {
        "title": "📩 Submit New Helpdesk Ticket",
        "emp_label": "Employee / Department Name:",
        "issue_label": "Issue Description:",
        "priority_label": "Priority Level:",
        "high": "High",
        "medium": "Medium",
        "low": "Low",
        "submit_btn": "🚀 Submit Ticket",
        "lang_btn": "العربية 🌐",
        "warn_fill": "Please fill in all required fields!",
        "success": "Ticket submitted successfully to IT Support!",
        "conn_error": "Could not connect to the network database:\n",
        "db_error": "An error occurred while submitting the ticket:\n"
    }
}

current_lang = "ar"

def init_db():
    try:
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS tickets (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                employee TEXT NOT NULL,
                issue TEXT NOT NULL,
                priority TEXT NOT NULL,
                status TEXT DEFAULT 'مفتوحة'
            )
        ''')
        conn.commit()
        conn.close()
    except Exception as e:
        messagebox.showerror("Network Error", f"{TRANSLATIONS[current_lang]['conn_error']}{e}")

def add_ticket():
    employee = entry_emp.get().strip()
    issue = entry_issue.get().strip()
    priority_val = combo_priority.get()

    if not employee or not issue:
        messagebox.showwarning("Warning", TRANSLATIONS[current_lang]["warn_fill"])
        return

    try:
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        cursor.execute(
            "INSERT INTO tickets (employee, issue, priority) VALUES (?, ?, ?)",
            (employee, issue, priority_val)
        )
        conn.commit()
        conn.close()

        messagebox.showinfo("Success", TRANSLATIONS[current_lang]["success"])
        entry_emp.delete(0, tk.END)
        entry_issue.delete(0, tk.END)
        combo_priority.current(1)
    except Exception as e:
        messagebox.showerror("Error", f"{TRANSLATIONS[current_lang]['db_error']}{e}")

def toggle_language():
    global current_lang
    current_lang = "en" if current_lang == "ar" else "ar"
    update_ui()

def update_ui():
    lang = TRANSLATIONS[current_lang]
    lbl_title.config(text=lang["title"])
    lbl_emp.config(text=lang["emp_label"])
    lbl_issue.config(text=lang["issue_label"])
    lbl_priority.config(text=lang["priority_label"])
    btn_add.config(text=lang["submit_btn"])
    btn_lang.config(text=lang["lang_btn"])
    
    combo_priority['values'] = [lang["high"], lang["medium"], lang["low"]]
    combo_priority.current(1)

# بناء الواجهة
init_db()

root = tk.Tk()
root.title("Helpdesk System - Submit Ticket")
root.geometry("500x480")
root.configure(bg="#F5F7FA")

FONT_FAMILY = "Segoe UI"

# الهيدر العلوي
header_frame = tk.Frame(root, bg="#1A365D", height=70)
header_frame.pack(fill="x", side="top")

lbl_title = tk.Label(header_frame, font=(FONT_FAMILY, 14, "bold"), bg="#1A365D", fg="#FFFFFF")
lbl_title.pack(pady=10)

btn_lang = tk.Button(header_frame, command=toggle_language, bg="#2B6CB0", fg="white", 
                     font=(FONT_FAMILY, 9, "bold"), relief="flat", padx=10, cursor="hand2")
btn_lang.pack(pady=(0, 10))

# بطاقة الإدخال بالتوسيط التام
card_input = tk.Frame(root, bg="#FFFFFF", relief="flat", bd=1)
card_input.pack(pady=20, padx=30, fill="both", expand=True)

lbl_emp = tk.Label(card_input, font=(FONT_FAMILY, 10, "bold"), bg="#FFFFFF", fg="#4A5568")
lbl_emp.pack(pady=(15, 2))
entry_emp = tk.Entry(card_input, font=(FONT_FAMILY, 10), bg="#F7FAFC", relief="solid", bd=1, justify="center")
entry_emp.pack(fill="x", padx=30, pady=5)

lbl_issue = tk.Label(card_input, font=(FONT_FAMILY, 10, "bold"), bg="#FFFFFF", fg="#4A5568")
lbl_issue.pack(pady=(10, 2))
entry_issue = tk.Entry(card_input, font=(FONT_FAMILY, 10), bg="#F7FAFC", relief="solid", bd=1, justify="center")
entry_issue.pack(fill="x", padx=30, pady=5)

lbl_priority = tk.Label(card_input, font=(FONT_FAMILY, 10, "bold"), bg="#FFFFFF", fg="#4A5568")
lbl_priority.pack(pady=(10, 2))
combo_priority = ttk.Combobox(card_input, state="readonly", font=(FONT_FAMILY, 10), justify="center")
combo_priority.pack(fill="x", padx=30, pady=5)

btn_add = tk.Button(card_input, command=add_ticket, bg="#3182CE", fg="white", 
                    font=(FONT_FAMILY, 11, "bold"), relief="flat", pady=8, cursor="hand2")
btn_add.pack(fill="x", padx=30, pady=25)

update_ui()
root.mainloop()