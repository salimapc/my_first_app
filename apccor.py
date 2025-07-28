import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from PIL import Image, ImageTk
import customtkinter as ctk
import datetime
import os
import pandas as pd
import shutil
import tempfile
import sqlite3
from tkinter import scrolledtext

# إعداد المظهر العام للتطبيق
ctk.set_appearance_mode("dark")  # يمكن تغييرها إلى "light" أو "system"
ctk.set_default_color_theme("blue")

class CorrespondenceArchiveApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        # إعداد النافذة الرئيسية
        self.title("نظام أرشفة المراسلات")
        self.geometry("1200x700")
        self.minsize(1000, 600)
        
        # إنشاء قاعدة بيانات
        self.create_database()
        
        # إعداد واجهة المستخدم
        self.setup_ui()
        
        # تحميل البيانات
        self.load_data()
        
    def create_database(self):
        self.conn = sqlite3.connect('correspondence.db')
        self.cursor = self.conn.cursor()
        
        # إنشاء جدول المراسلات
        self.cursor.execute('''
        CREATE TABLE IF NOT EXISTS correspondence (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            type TEXT NOT NULL,
            reference_number TEXT NOT NULL,
            date TEXT NOT NULL,
            sender TEXT,
            receiver TEXT,
            subject TEXT NOT NULL,
            department TEXT,
            priority TEXT,
            status TEXT,
            notes TEXT,
            attachment_path TEXT
        )
        ''')
        
        self.conn.commit()
    
    def setup_ui(self):
        # إنشاء تخطيط الشبكة
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)
        
        # إنشاء الشريط الجانبي
        self.sidebar_frame = ctk.CTkFrame(self, width=140, corner_radius=0)
        self.sidebar_frame.grid(row=0, column=0, rowspan=4, sticky="nsew")
        self.sidebar_frame.grid_rowconfigure(5, weight=1)
        
        # شعار التطبيق
        self.logo_label = ctk.CTkLabel(
            self.sidebar_frame, 
            text="أرشيف المراسلات",
            font=ctk.CTkFont(size=20, weight="bold")
        )
        self.logo_label.grid(row=0, column=0, padx=20, pady=(20, 10))
        
        # أزرار التنقل
        self.add_btn = ctk.CTkButton(
            self.sidebar_frame, 
            text="إضافة مراسلة جديدة", 
            command=self.show_add_correspondence
        )
        self.add_btn.grid(row=1, column=0, padx=20, pady=10)
        
        self.view_btn = ctk.CTkButton(
            self.sidebar_frame, 
            text="عرض جميع المراسلات", 
            command=self.show_all_correspondence
        )
        self.view_btn.grid(row=2, column=0, padx=20, pady=10)
        
        self.stats_btn = ctk.CTkButton(
            self.sidebar_frame, 
            text="الإحصائيات والتقارير", 
            command=self.show_statistics
        )
        self.stats_btn.grid(row=3, column=0, padx=20, pady=10)
        
        self.scan_btn = ctk.CTkButton(
            self.sidebar_frame, 
            text="مسح ضوئي", 
            command=self.show_scan_interface,
            fg_color="#D32F2F",
            hover_color="#B71C1C"
        )
        self.scan_btn.grid(row=4, column=0, padx=20, pady=10)
        
        # خيارات المظهر
        self.appearance_label = ctk.CTkLabel(
            self.sidebar_frame, 
            text="مظهر التطبيق:", 
            anchor="w"
        )
        self.appearance_label.grid(row=6, column=0, padx=20, pady=(10, 0))
        
        self.appearance_option = ctk.CTkOptionMenu(
            self.sidebar_frame, 
            values=["فاتح", "داكن", "نظام"],
            command=self.change_appearance
        )
        self.appearance_option.grid(row=7, column=0, padx=20, pady=10)
        
        # إصدار التطبيق
        self.version_label = ctk.CTkLabel(
            self.sidebar_frame, 
            text="الإصدار 1.0.0",
            font=ctk.CTkFont(size=10)
        )
        self.version_label.grid(row=8, column=0, padx=20, pady=(10, 20))
        
        # إنشاء منطقة المحتوى الرئيسية
        self.main_frame = ctk.CTkFrame(self, corner_radius=0)
        self.main_frame.grid(row=0, column=1, sticky="nsew")
        self.main_frame.grid_columnconfigure(0, weight=1)
        self.main_frame.grid_rowconfigure(0, weight=1)
        
        # عرض الشاشة الرئيسية
        self.show_home_screen()
    
    def show_home_screen(self):
        # مسح المحتوى الحالي
        for widget in self.main_frame.winfo_children():
            widget.destroy()
        
        # إنشاء شاشة الترحيب
        welcome_frame = ctk.CTkFrame(self.main_frame, corner_radius=0, fg_color="transparent")
        welcome_frame.grid(row=0, column=0, sticky="nsew", padx=20, pady=20)
        welcome_frame.grid_columnconfigure(0, weight=1)
        welcome_frame.grid_rowconfigure(1, weight=1)
        
        # عنوان الترحيب
        welcome_label = ctk.CTkLabel(
            welcome_frame,
            text="مرحباً بك في نظام أرشفة المراسلات",
            font=ctk.CTkFont(size=24, weight="bold")
        )
        welcome_label.grid(row=0, column=0, pady=(0, 20))
        
        # إطار الإحصائيات
        stats_frame = ctk.CTkFrame(welcome_frame)
        stats_frame.grid(row=1, column=0, sticky="nsew", pady=20)
        stats_frame.grid_columnconfigure(0, weight=1)
        stats_frame.grid_columnconfigure(1, weight=1)
        stats_frame.grid_rowconfigure(0, weight=1)
        
        # إحصائيات المراسلات الواردة
        incoming_frame = ctk.CTkFrame(stats_frame)
        incoming_frame.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")
        
        incoming_label = ctk.CTkLabel(
            incoming_frame,
            text="المراسلات الواردة",
            font=ctk.CTkFont(size=18, weight="bold")
        )
        incoming_label.pack(pady=(10, 5))
        
        self.incoming_count = ctk.CTkLabel(
            incoming_frame,
            text="0",
            font=ctk.CTkFont(size=36, weight="bold"),
            text_color="#4CAF50"
        )
        self.incoming_count.pack(pady=(5, 10))
        
        # إحصائيات المراسلات الصادرة
        outgoing_frame = ctk.CTkFrame(stats_frame)
        outgoing_frame.grid(row=0, column=1, padx=20, pady=20, sticky="nsew")
        
        outgoing_label = ctk.CTkLabel(
            outgoing_frame,
            text="المراسلات الصادرة",
            font=ctk.CTkFont(size=18, weight="bold")
        )
        outgoing_label.pack(pady=(10, 5))
        
        self.outgoing_count = ctk.CTkLabel(
            outgoing_frame,
            text="0",
            font=ctk.CTkFont(size=36, weight="bold"),
            text_color="#2196F3"
        )
        self.outgoing_count.pack(pady=(5, 10))
        
        # آخر 5 مراسلات
        recent_label = ctk.CTkLabel(
            welcome_frame,
            text="آخر المراسلات",
            font=ctk.CTkFont(size=18, weight="bold")
        )
        recent_label.grid(row=2, column=0, pady=(20, 10), sticky="w")
        
        self.recent_table = ttk.Treeview(
            welcome_frame,
            columns=("id", "type", "reference", "date", "subject"),
            show="headings",
            height=5
        )
        
        # تنسيق الجدول
        style = ttk.Style()
        style.configure("Treeview", font=("Arial", 11))
        style.configure("Treeview.Heading", font=("Arial", 12, "bold"))
        
        self.recent_table.heading("id", text="ID")
        self.recent_table.heading("type", text="النوع")
        self.recent_table.heading("reference", text="رقم المرجع")
        self.recent_table.heading("date", text="التاريخ")
        self.recent_table.heading("subject", text="الموضوع")
        
        self.recent_table.column("id", width=50, anchor="center")
        self.recent_table.column("type", width=100, anchor="center")
        self.recent_table.column("reference", width=150, anchor="center")
        self.recent_table.column("date", width=120, anchor="center")
        self.recent_table.column("subject", width=400, anchor="w")
        
        self.recent_table.grid(row=3, column=0, sticky="ew", pady=(0, 20))
        
        # تحديث الإحصائيات والبيانات
        self.update_stats()
    
    def update_stats(self):
        # حساب عدد المراسلات الواردة
        self.cursor.execute("SELECT COUNT(*) FROM correspondence WHERE type='واردة'")
        incoming = self.cursor.fetchone()[0]
        self.incoming_count.configure(text=str(incoming))
        
        # حساب عدد المراسلات الصادرة
        self.cursor.execute("SELECT COUNT(*) FROM correspondence WHERE type='صادرة'")
        outgoing = self.cursor.fetchone()[0]
        self.outgoing_count.configure(text=str(outgoing))
        
        # جلب آخر 5 مراسلات
        self.cursor.execute("SELECT id, type, reference_number, date, subject FROM correspondence ORDER BY date DESC LIMIT 5")
        recent = self.cursor.fetchall()
        
        # تحديث الجدول
        for item in self.recent_table.get_children():
            self.recent_table.delete(item)
            
        for row in recent:
            self.recent_table.insert("", "end", values=row)
    
    def show_add_correspondence(self):
        # مسح المحتوى الحالي
        for widget in self.main_frame.winfo_children():
            widget.destroy()
        
        # إنشاء نموذج إضافة مراسلة
        form_frame = ctk.CTkFrame(self.main_frame, corner_radius=0, fg_color="transparent")
        form_frame.grid(row=0, column=0, sticky="nsew", padx=20, pady=20)
        form_frame.grid_columnconfigure(0, weight=1)
        
        # عنوان النموذج
        title_label = ctk.CTkLabel(
            form_frame,
            text="إضافة مراسلة جديدة",
            font=ctk.CTkFont(size=24, weight="bold")
        )
        title_label.grid(row=0, column=0, pady=(0, 20), sticky="w")
        
        # إنشاء إطار النموذج
        form_container = ctk.CTkFrame(form_frame)
        form_container.grid(row=1, column=0, sticky="nsew")
        form_container.grid_columnconfigure(1, weight=1)
        
        # نوع المراسلة
        ctk.CTkLabel(form_container, text="نوع المراسلة:", font=ctk.CTkFont(weight="bold")).grid(
            row=0, column=0, padx=20, pady=10, sticky="w")
        
        self.corr_type = ctk.CTkOptionMenu(
            form_container, 
            values=["واردة", "صادرة"],
            dynamic_resizing=False
        )
        self.corr_type.grid(row=0, column=1, padx=20, pady=10, sticky="ew")
        
        # رقم المرجع
        ctk.CTkLabel(form_container, text="رقم المرجع:", font=ctk.CTkFont(weight="bold")).grid(
            row=1, column=0, padx=20, pady=10, sticky="w")
        
        self.ref_number = ctk.CTkEntry(form_container)
        self.ref_number.grid(row=1, column=1, padx=20, pady=10, sticky="ew")
        
        # التاريخ
        ctk.CTkLabel(form_container, text="التاريخ:", font=ctk.CTkFont(weight="bold")).grid(
            row=2, column=0, padx=20, pady=10, sticky="w")
        
        today = datetime.date.today()
        self.date_entry = ctk.CTkEntry(form_container)
        self.date_entry.insert(0, today.strftime("%Y-%m-%d"))
        self.date_entry.grid(row=2, column=1, padx=20, pady=10, sticky="ew")
        
        # المرسل/المستلم
        self.sender_receiver_label = ctk.CTkLabel(
            form_container, 
            text="المرسل:", 
            font=ctk.CTkFont(weight="bold")
        )
        self.sender_receiver_label.grid(row=3, column=0, padx=20, pady=10, sticky="w")
        
        self.sender_receiver_entry = ctk.CTkEntry(form_container)
        self.sender_receiver_entry.grid(row=3, column=1, padx=20, pady=10, sticky="ew")
        
        # تحديث التسمية عند تغيير النوع
        self.corr_type.configure(command=self.update_sender_receiver_label)
        
        # الموضوع
        ctk.CTkLabel(form_container, text="الموضوع:", font=ctk.CTkFont(weight="bold")).grid(
            row=4, column=0, padx=20, pady=10, sticky="w")
        
        self.subject_entry = ctk.CTkEntry(form_container)
        self.subject_entry.grid(row=4, column=1, padx=20, pady=10, sticky="ew")
        
        # الإدارة
        ctk.CTkLabel(form_container, text="الإدارة:", font=ctk.CTkFont(weight="bold")).grid(
            row=5, column=0, padx=20, pady=10, sticky="w")
        
        self.department_entry = ctk.CTkEntry(form_container)
        self.department_entry.grid(row=5, column=1, padx=20, pady=10, sticky="ew")
        
        # الأولوية
        ctk.CTkLabel(form_container, text="الأولوية:", font=ctk.CTkFont(weight="bold")).grid(
            row=6, column=0, padx=20, pady=10, sticky="w")
        
        self.priority = ctk.CTkOptionMenu(
            form_container, 
            values=["عادية", "متوسطة", "عاجلة"],
            dynamic_resizing=False
        )
        self.priority.grid(row=6, column=1, padx=20, pady=10, sticky="ew")
        
        # الحالة
        ctk.CTkLabel(form_container, text="الحالة:", font=ctk.CTkFont(weight="bold")).grid(
            row=7, column=0, padx=20, pady=10, sticky="w")
        
        self.status = ctk.CTkOptionMenu(
            form_container, 
            values=["مستلمة", "قيد المعالجة", "مكتملة", "ملغاة"],
            dynamic_resizing=False
        )
        self.status.grid(row=7, column=1, padx=20, pady=10, sticky="ew")
        
        # الملاحظات
        ctk.CTkLabel(form_container, text="ملاحظات:", font=ctk.CTkFont(weight="bold")).grid(
            row=8, column=0, padx=20, pady=10, sticky="w")
        
        self.notes_text = scrolledtext.ScrolledText(
            form_container, 
            width=40, 
            height=5,
            font=("Arial", 12)
        )
        self.notes_text.grid(row=8, column=1, padx=20, pady=10, sticky="ew")
        
        # المرفقات
        ctk.CTkLabel(form_container, text="المرفقات:", font=ctk.CTkFont(weight="bold")).grid(
            row=9, column=0, padx=20, pady=10, sticky="w")
        
        attachment_frame = ctk.CTkFrame(form_container, fg_color="transparent")
        attachment_frame.grid(row=9, column=1, padx=20, pady=10, sticky="ew")
        attachment_frame.grid_columnconfigure(0, weight=1)
        
        self.attachment_path = ctk.CTkLabel(
            attachment_frame, 
            text="لا يوجد ملف مرفق",
            anchor="w"
        )
        self.attachment_path.grid(row=0, column=0, sticky="ew")
        
        self.attachment_btn = ctk.CTkButton(
            attachment_frame, 
            text="اختيار ملف",
            width=100,
            command=self.attach_file
        )
        self.attachment_btn.grid(row=0, column=1, padx=(10, 0))
        
        # أزرار الحفظ والإلغاء
        button_frame = ctk.CTkFrame(form_frame, fg_color="transparent")
        button_frame.grid(row=2, column=0, pady=20)
        
        save_btn = ctk.CTkButton(
            button_frame, 
            text="حفظ المراسلة",
            command=self.save_correspondence,
            fg_color="#4CAF50",
            hover_color="#388E3C"
        )
        save_btn.pack(side="left", padx=10)
        
        cancel_btn = ctk.CTkButton(
            button_frame, 
            text="إلغاء",
            command=self.show_home_screen,
            fg_color="#F44336",
            hover_color="#D32F2F"
        )
        cancel_btn.pack(side="left", padx=10)
    
    def update_sender_receiver_label(self, value):
        if value == "واردة":
            self.sender_receiver_label.configure(text="المرسل:")
        else:
            self.sender_receiver_label.configure(text="المرسل إليه:")
    
    def attach_file(self):
        file_path = filedialog.askopenfilename(
            title="اختر ملف مرفق",
            filetypes=[("صورة", "*.jpg *.jpeg *.png"), ("PDF", "*.pdf"), ("كل الملفات", "*.*")]
        )
        
        if file_path:
            self.attachment_path.configure(text=os.path.basename(file_path))
            self.attachment_path._text = file_path
    
    def save_correspondence(self):
        # جمع بيانات النموذج
        corr_type = self.corr_type.get()
        ref_number = self.ref_number.get()
        date = self.date_entry.get()
        sender_receiver = self.sender_receiver_entry.get()
        subject = self.subject_entry.get()
        department = self.department_entry.get()
        priority = self.priority.get()
        status = self.status.get()
        notes = self.notes_text.get("1.0", tk.END).strip()
        attachment = getattr(self.attachment_path, "_text", "")
        
        # التحقق من البيانات المطلوبة
        if not ref_number or not date or not sender_receiver or not subject:
            messagebox.showerror("خطأ", "يرجى ملء الحقول المطلوبة")
            return
        
        try:
            # حفظ البيانات في قاعدة البيانات
            self.cursor.execute('''
            INSERT INTO correspondence (
                type, reference_number, date, sender, receiver, subject, 
                department, priority, status, notes, attachment_path
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                corr_type, ref_number, date, 
                sender_receiver if corr_type == "واردة" else None,
                sender_receiver if corr_type == "صادرة" else None,
                subject, department, priority, status, notes, attachment
            ))
            
            self.conn.commit()
            
            # نسخ الملف المرفق إلى مجلد المرفقات
            if attachment:
                attachments_dir = "attachments"
                if not os.path.exists(attachments_dir):
                    os.makedirs(attachments_dir)
                
                filename = os.path.basename(attachment)
                dest_path = os.path.join(attachments_dir, f"{self.cursor.lastrowid}_{filename}")
                shutil.copyfile(attachment, dest_path)
                
                # تحديث مسار المرفق في قاعدة البيانات
                self.cursor.execute('''
                UPDATE correspondence SET attachment_path=? WHERE id=?
                ''', (dest_path, self.cursor.lastrowid))
                self.conn.commit()
            
            messagebox.showinfo("نجاح", "تم حفظ المراسلة بنجاح")
            self.show_home_screen()
            
        except Exception as e:
            messagebox.showerror("خطأ", f"حدث خطأ أثناء الحفظ: {str(e)}")
    
    def show_all_correspondence(self):
        # مسح المحتوى الحالي
        for widget in self.main_frame.winfo_children():
            widget.destroy()
        
        # إنشاء إطار عرض المراسلات
        view_frame = ctk.CTkFrame(self.main_frame, corner_radius=0, fg_color="transparent")
        view_frame.grid(row=0, column=0, sticky="nsew", padx=20, pady=20)
        view_frame.grid_columnconfigure(0, weight=1)
        view_frame.grid_rowconfigure(1, weight=1)
        
        # عنوان الصفحة
        title_label = ctk.CTkLabel(
            view_frame,
            text="جميع المراسلات",
            font=ctk.CTkFont(size=24, weight="bold")
        )
        title_label.grid(row=0, column=0, pady=(0, 20), sticky="w")
        
        # أدوات البحث والتصفية
        filter_frame = ctk.CTkFrame(view_frame, fg_color="transparent")
        filter_frame.grid(row=1, column=0, sticky="nsew", pady=(0, 10))
        filter_frame.grid_columnconfigure(0, weight=1)
        
        # حقل البحث
        search_frame = ctk.CTkFrame(filter_frame, fg_color="transparent")
        search_frame.grid(row=0, column=0, sticky="ew", pady=5)
        search_frame.grid_columnconfigure(0, weight=1)
        
        self.search_entry = ctk.CTkEntry(
            search_frame, 
            placeholder_text="ابحث بالموضوع أو الرقم المرجعي..."
        )
        self.search_entry.grid(row=0, column=0, padx=(0, 10), sticky="ew")
        
        search_btn = ctk.CTkButton(
            search_frame, 
            text="بحث",
            width=80,
            command=self.search_correspondence
        )
        search_btn.grid(row=0, column=1)
        
        # أزرار التصفية
        filter_buttons_frame = ctk.CTkFrame(filter_frame, fg_color="transparent")
        filter_buttons_frame.grid(row=1, column=0, sticky="w", pady=5)
        
        ctk.CTkButton(
            filter_buttons_frame, 
            text="الكل",
            width=100,
            command=self.load_data
        ).pack(side="left", padx=5)
        
        ctk.CTkButton(
            filter_buttons_frame, 
            text="الواردة",
            width=100,
            command=lambda: self.filter_correspondence("واردة")
        ).pack(side="left", padx=5)
        
        ctk.CTkButton(
            filter_buttons_frame, 
            text="الصادرة",
            width=100,
            command=lambda: self.filter_correspondence("صادرة")
        ).pack(side="left", padx=5)
        
        # جدول المراسلات
        table_frame = ctk.CTkFrame(view_frame)
        table_frame.grid(row=2, column=0, sticky="nsew")
        table_frame.grid_columnconfigure(0, weight=1)
        table_frame.grid_rowconfigure(0, weight=1)
        
        # إنشاء شريط تمرير
        scrollbar = ttk.Scrollbar(table_frame)
        scrollbar.pack(side="right", fill="y")
        
        # إنشاء الجدول
        self.correspondence_table = ttk.Treeview(
            table_frame,
            columns=("id", "type", "reference", "date", "subject", "status"),
            show="headings",
            yscrollcommand=scrollbar.set
        )
        
        # تنسيق الجدول
        style = ttk.Style()
        style.configure("Treeview", font=("Arial", 11), rowheight=25)
        style.configure("Treeview.Heading", font=("Arial", 12, "bold"))
        
        # تعريف الأعمدة
        self.correspondence_table.heading("id", text="ID")
        self.correspondence_table.heading("type", text="النوع")
        self.correspondence_table.heading("reference", text="رقم المرجع")
        self.correspondence_table.heading("date", text="التاريخ")
        self.correspondence_table.heading("subject", text="الموضوع")
        self.correspondence_table.heading("status", text="الحالة")
        
        self.correspondence_table.column("id", width=50, anchor="center")
        self.correspondence_table.column("type", width=100, anchor="center")
        self.correspondence_table.column("reference", width=150, anchor="center")
        self.correspondence_table.column("date", width=120, anchor="center")
        self.correspondence_table.column("subject", width=400, anchor="w")
        self.correspondence_table.column("status", width=120, anchor="center")
        
        self.correspondence_table.pack(fill="both", expand=True)
        scrollbar.config(command=self.correspondence_table.yview)
        
        # ربط حدث النقر المزدوج لعرض التفاصيل
        self.correspondence_table.bind("<Double-1>", self.show_correspondence_details)
        
        # أزرار التصدير
        export_frame = ctk.CTkFrame(view_frame, fg_color="transparent")
        export_frame.grid(row=3, column=0, sticky="e", pady=10)
        
        ctk.CTkButton(
            export_frame, 
            text="تصدير إلى Excel",
            command=self.export_to_excel
        ).pack(side="left", padx=5)
        
        # تحميل البيانات في الجدول
        self.load_data()
    
    def load_data(self, search_query=None, filter_type=None):
        # جلب البيانات من قاعدة البيانات
        query = "SELECT id, type, reference_number, date, subject, status FROM correspondence"
        params = []
        
        if search_query:
            query += " WHERE subject LIKE ? OR reference_number LIKE ?"
            params = [f"%{search_query}%", f"%{search_query}%"]
        elif filter_type:
            query += " WHERE type = ?"
            params = [filter_type]
        
        query += " ORDER BY date DESC"
        
        self.cursor.execute(query, params)
        data = self.cursor.fetchall()
        
        # تحديث الجدول
        for item in self.correspondence_table.get_children():
            self.correspondence_table.delete(item)
            
        for row in data:
            self.correspondence_table.insert("", "end", values=row)
    
    def search_correspondence(self):
        query = self.search_entry.get().strip()
        if query:
            self.load_data(search_query=query)
        else:
            self.load_data()
    
    def filter_correspondence(self, corr_type):
        self.load_data(filter_type=corr_type)
    
    def show_correspondence_details(self, event):
        # الحصول على العنصر المحدد
        selected_item = self.correspondence_table.selection()
        if not selected_item:
            return
        
        item_id = self.correspondence_table.item(selected_item)["values"][0]
        
        # جلب بيانات المراسلة من قاعدة البيانات
        self.cursor.execute("SELECT * FROM correspondence WHERE id=?", (item_id,))
        correspondence = self.cursor.fetchone()
        
        if not correspondence:
            return
        
        # إنشاء نافذة التفاصيل
        details_window = ctk.CTkToplevel(self)
        details_window.title("تفاصيل المراسلة")
        details_window.geometry("800x600")
        details_window.grab_set()
        
        # إطار المحتوى
        content_frame = ctk.CTkFrame(details_window)
        content_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # عرض بيانات المراسلة
        labels = [
            ("النوع", correspondence[1]),
            ("رقم المرجع", correspondence[2]),
            ("التاريخ", correspondence[3]),
            ("المرسل" if correspondence[1] == "واردة" else "المرسل إليه", 
             correspondence[4] if correspondence[1] == "واردة" else correspondence[5]),
            ("الموضوع", correspondence[6]),
            ("الإدارة", correspondence[7]),
            ("الأولوية", correspondence[8]),
            ("الحالة", correspondence[9]),
            ("الملاحظات", correspondence[10])
        ]
        
        for i, (label, value) in enumerate(labels):
            row_frame = ctk.CTkFrame(content_frame, fg_color="transparent")
            row_frame.grid(row=i, column=0, sticky="ew", padx=10, pady=5)
            row_frame.grid_columnconfigure(0, weight=1)
            row_frame.grid_columnconfigure(1, weight=3)
            
            ctk.CTkLabel(
                row_frame, 
                text=f"{label}:",
                font=ctk.CTkFont(weight="bold")
            ).grid(row=0, column=0, padx=(0, 10), sticky="w")
            
            ctk.CTkLabel(
                row_frame, 
                text=value or "---",
                anchor="w"
            ).grid(row=0, column=1, sticky="w")
        
        # عرض المرفق إذا وجد
        if correspondence[11]:
            attachment_frame = ctk.CTkFrame(content_frame, fg_color="transparent")
            attachment_frame.grid(row=len(labels), column=0, sticky="ew", padx=10, pady=5)
            
            ctk.CTkLabel(
                attachment_frame, 
                text="المرفق:",
                font=ctk.CTkFont(weight="bold")
            ).grid(row=0, column=0, padx=(0, 10), sticky="w")
            
            # عرض صورة مصغرة للمرفق إذا كان صورة
            attachment_path = correspondence[11]
            filename = os.path.basename(attachment_path)
            
            if attachment_path.lower().endswith((".jpg", ".jpeg", ".png")):
                try:
                    img = Image.open(attachment_path)
                    img.thumbnail((200, 200))
                    photo = ImageTk.PhotoImage(img)
                    
                    img_label = ctk.CTkLabel(attachment_frame, text="", image=photo)
                    img_label.image = photo
                    img_label.grid(row=0, column=1, padx=10, pady=10)
                    
                except Exception:
                    ctk.CTkLabel(
                        attachment_frame, 
                        text=f"ملف صورة: {filename}"
                    ).grid(row=0, column=1, sticky="w")
            else:
                ctk.CTkLabel(
                    attachment_frame, 
                    text=f"ملف مرفق: {filename}"
                ).grid(row=0, column=1, sticky="w")
            
            # زر فتح المرفق
            ctk.CTkButton(
                attachment_frame,
                text="فتح المرفق",
                command=lambda p=attachment_path: os.startfile(p)
            ).grid(row=0, column=2, padx=10)
        
        # أزرار الحذف والإغلاق
        button_frame = ctk.CTkFrame(content_frame, fg_color="transparent")
        button_frame.grid(row=len(labels)+1, column=0, pady=20)
        
        delete_btn = ctk.CTkButton(
            button_frame, 
            text="حذف المراسلة",
            command=lambda: self.delete_correspondence(item_id, details_window),
            fg_color="#F44336",
            hover_color="#D32F2F"
        )
        delete_btn.pack(side="left", padx=10)
        
        close_btn = ctk.CTkButton(
            button_frame, 
            text="إغلاق",
            command=details_window.destroy
        )
        close_btn.pack(side="left", padx=10)
    
    def delete_correspondence(self, item_id, window):
        if messagebox.askyesno("تأكيد", "هل أنت متأكد من حذف هذه المراسلة؟"):
            try:
                # حذف المرفق إذا وجد
                self.cursor.execute("SELECT attachment_path FROM correspondence WHERE id=?", (item_id,))
                attachment = self.cursor.fetchone()[0]
                
                if attachment and os.path.exists(attachment):
                    os.remove(attachment)
                
                # حذف من قاعدة البيانات
                self.cursor.execute("DELETE FROM correspondence WHERE id=?", (item_id,))
                self.conn.commit()
                
                messagebox.showinfo("نجاح", "تم حذف المراسلة بنجاح")
                window.destroy()
                self.load_data()
                
            except Exception as e:
                messagebox.showerror("خطأ", f"حدث خطأ أثناء الحذف: {str(e)}")
    
    def show_scan_interface(self):
        # مسح المحتوى الحالي
        for widget in self.main_frame.winfo_children():
            widget.destroy()
        
        # إنشاء واجهة المسح الضوئي
        scan_frame = ctk.CTkFrame(self.main_frame, corner_radius=0, fg_color="transparent")
        scan_frame.grid(row=0, column=0, sticky="nsew", padx=20, pady=20)
        scan_frame.grid_columnconfigure(0, weight=1)
        scan_frame.grid_rowconfigure(1, weight=1)
        
        # عنوان الصفحة
        title_label = ctk.CTkLabel(
            scan_frame,
            text="مسح ضوئي للمراسلات",
            font=ctk.CTkFont(size=24, weight="bold")
        )
        title_label.grid(row=0, column=0, pady=(0, 20), sticky="w")
        
        # إطار المسح الضوئي
        scanner_frame = ctk.CTkFrame(scan_frame)
        scanner_frame.grid(row=1, column=0, sticky="nsew")
        scanner_frame.grid_columnconfigure(0, weight=1)
        scanner_frame.grid_rowconfigure(1, weight=1)
        
        # تعليمات المسح الضوئي
        instructions = ctk.CTkLabel(
            scanner_frame,
            text="1. ضع المستند في الماسح الضوئي\n2. اضغط على زر المسح الضوئي\n3. راجع الصورة الممسوحة",
            font=ctk.CTkFont(size=16),
            justify="left"
        )
        instructions.grid(row=0, column=0, padx=20, pady=20, sticky="w")
        
        # منطقة معاينة الصورة
        preview_frame = ctk.CTkFrame(scanner_frame)
        preview_frame.grid(row=1, column=0, padx=20, pady=(0, 20), sticky="nsew")
        preview_frame.grid_columnconfigure(0, weight=1)
        preview_frame.grid_rowconfigure(0, weight=1)
        
        self.scan_preview = ctk.CTkLabel(
            preview_frame, 
            text="لا توجد صورة ممسوحة",
            fg_color="#2B2B2B",
            corner_radius=5
        )
        self.scan_preview.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        
        # أزرار التحكم
        button_frame = ctk.CTkFrame(scanner_frame, fg_color="transparent")
        button_frame.grid(row=2, column=0, pady=(0, 20))
        
        scan_btn = ctk.CTkButton(
            button_frame, 
            text="مسح ضوئي",
            command=self.scan_document,
            fg_color="#4CAF50",
            hover_color="#388E3C"
        )
        scan_btn.pack(side="left", padx=10)
        
        save_btn = ctk.CTkButton(
            button_frame, 
            text="حفظ الصورة",
            command=self.save_scanned_image,
            state="disabled"
        )
        save_btn.pack(side="left", padx=10)
        self.save_scan_btn = save_btn
        
        # ملفات المسح الضوئي الحديثة
        recent_label = ctk.CTkLabel(
            scan_frame,
            text="آخر الملفات الممسوحة",
            font=ctk.CTkFont(size=18, weight="bold")
        )
        recent_label.grid(row=2, column=0, pady=(20, 10), sticky="w")
        
        self.recent_scans_list = ctk.CTkScrollableFrame(
            scan_frame,
            height=100
        )
        self.recent_scans_list.grid(row=3, column=0, sticky="ew", pady=(0, 20))
        
        # تحديث قائمة الملفات الحديثة
        self.update_recent_scans()
    
    def scan_document(self):
        # في بيئة حقيقية، سيتم استدعاء واجهة برمجة تطبيقات الماسح الضوئي هنا
        # سنقوم بمحاكاة المسح الضوئي بفتح صورة عينة
        
        file_path = filedialog.askopenfilename(
            title="اختر صورة للمسح الضوئي",
            filetypes=[("صورة", "*.jpg *.jpeg *.png")]
        )
        
        if file_path:
            try:
                # حفظ الصورة مؤقتاً للمعاينة
                self.temp_scan_path = file_path
                
                # عرض معاينة الصورة
                img = Image.open(file_path)
                img.thumbnail((600, 600))
                photo = ImageTk.PhotoImage(img)
                
                self.scan_preview.configure(image=photo, text="")
                self.scan_preview.image = photo
                
                # تمكين زر الحفظ
                self.save_scan_btn.configure(state="normal")
                
            except Exception as e:
                messagebox.showerror("خطأ", f"تعذر تحميل الصورة: {str(e)}")
    
    def save_scanned_image(self):
        if hasattr(self, "temp_scan_path"):
            # حفظ الصورة في مجلد المرفقات
            attachments_dir = "attachments"
            if not os.path.exists(attachments_dir):
                os.makedirs(attachments_dir)
            
            filename = f"scan_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.jpg"
            dest_path = os.path.join(attachments_dir, filename)
            shutil.copyfile(self.temp_scan_path, dest_path)
            
            messagebox.showinfo("نجاح", f"تم حفظ الصورة الممسوحة في:\n{dest_path}")
            self.update_recent_scans()
    
    def update_recent_scans(self):
        # مسح القائمة الحالية
        for widget in self.recent_scans_list.winfo_children():
            widget.destroy()
        
        # جلب آخر 5 ملفات ممسوحة
        if os.path.exists("attachments"):
            files = []
            for f in os.listdir("attachments"):
                if f.startswith("scan_"):
                    files.append(os.path.join("attachments", f))
            
            files.sort(key=os.path.getmtime, reverse=True)
            recent_files = files[:5]
            
            # عرض الملفات
            for file_path in recent_files:
                file_name = os.path.basename(file_path)
                file_frame = ctk.CTkFrame(self.recent_scans_list, height=30)
                file_frame.pack(fill="x", pady=5, padx=5)
                
                ctk.CTkLabel(
                    file_frame, 
                    text=file_name,
                    anchor="w"
                ).pack(side="left", padx=10, fill="x", expand=True)
                
                ctk.CTkButton(
                    file_frame, 
                    text="فتح",
                    width=60,
                    command=lambda p=file_path: os.startfile(p)
                ).pack(side="right", padx=5)
    
    def show_statistics(self):
        # مسح المحتوى الحالي
        for widget in self.main_frame.winfo_children():
            widget.destroy()
        
        # إنشاء واجهة الإحصائيات
        stats_frame = ctk.CTkFrame(self.main_frame, corner_radius=0, fg_color="transparent")
        stats_frame.grid(row=0, column=0, sticky="nsew", padx=20, pady=20)
        stats_frame.grid_columnconfigure(0, weight=1)
        stats_frame.grid_rowconfigure(1, weight=1)
        
        # عنوان الصفحة
        title_label = ctk.CTkLabel(
            stats_frame,
            text="الإحصائيات والتقارير",
            font=ctk.CTkFont(size=24, weight="bold")
        )
        title_label.grid(row=0, column=0, pady=(0, 20), sticky="w")
        
        # إطار الإحصائيات
        charts_frame = ctk.CTkFrame(stats_frame)
        charts_frame.grid(row=1, column=0, sticky="nsew")
        charts_frame.grid_columnconfigure(0, weight=1)
        charts_frame.grid_columnconfigure(1, weight=1)
        charts_frame.grid_rowconfigure(0, weight=1)
        
        # مخطط المراسلات حسب النوع
        type_frame = ctk.CTkFrame(charts_frame)
        type_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")
        
        ctk.CTkLabel(
            type_frame,
            text="المراسلات حسب النوع",
            font=ctk.CTkFont(size=16, weight="bold")
        ).pack(pady=10)
        
        # جلب البيانات
        self.cursor.execute("SELECT type, COUNT(*) FROM correspondence GROUP BY type")
        type_data = self.cursor.fetchall()
        
        # إنشاء مخطط دائري (مبسط)
        canvas = ctk.CTkCanvas(type_frame, bg="#2B2B2B", highlightthickness=0)
        canvas.pack(fill="both", expand=True, padx=20, pady=20)
        
        width = canvas.winfo_reqwidth()
        height = canvas.winfo_reqheight()
        
        colors = ["#4CAF50", "#2196F3"]
        total = sum(count for _, count in type_data)
        start_angle = 0
        
        for i, (corr_type, count) in enumerate(type_data):
            extent = 360 * count / total
            canvas.create_arc(
                10, 10, width-10, height-10,
                start=start_angle, extent=extent,
                fill=colors[i], outline=""
            )
            start_angle += extent
            
            # إضافة وسيلة إيضاح
            canvas.create_rectangle(width-150, 20 + i*30, width-130, 40 + i*30, fill=colors[i], outline="")
            canvas.create_text(
                width-120, 30 + i*30,
                text=f"{corr_type}: {count} ({count/total:.0%})",
                anchor="w",
                fill="white",
                font=("Arial", 12)
            )
        
        # مخطط المراسلات حسب الحالة
        status_frame = ctk.CTkFrame(charts_frame)
        status_frame.grid(row=0, column=1, padx=10, pady=10, sticky="nsew")
        
        ctk.CTkLabel(
            status_frame,
            text="المراسلات حسب الحالة",
            font=ctk.CTkFont(size=16, weight="bold")
        ).pack(pady=10)
        
        # جلب البيانات
        self.cursor.execute("SELECT status, COUNT(*) FROM correspondence GROUP BY status")
        status_data = self.cursor.fetchall()
        
        # إنشاء مخطط أعمدة (مبسط)
        canvas = ctk.CTkCanvas(status_frame, bg="#2B2B2B", highlightthickness=0)
        canvas.pack(fill="both", expand=True, padx=20, pady=20)
        
        width = canvas.winfo_reqwidth()
        height = canvas.winfo_reqheight()
        
        colors = ["#4CAF50", "#2196F3", "#FFC107", "#F44336"]
        max_count = max(count for _, count in status_data) if status_data else 1
        bar_width = 50
        spacing = 30
        start_x = 50
        
        for i, (status, count) in enumerate(status_data):
            bar_height = (count / max_count) * (height - 100)
            canvas.create_rectangle(
                start_x, height - 50 - bar_height,
                start_x + bar_width, height - 50,
                fill=colors[i], outline=""
            )
            canvas.create_text(
                start_x + bar_width/2, height - 30,
                text=status,
                anchor="n",
                fill="white",
                font=("Arial", 10)
            )
            canvas.create_text(
                start_x + bar_width/2, height - 55 - bar_height,
                text=str(count),
                anchor="s",
                fill="white",
                font=("Arial", 10)
            )
            start_x += bar_width + spacing
    
    def export_to_excel(self):
        # جلب جميع البيانات
        self.cursor.execute("SELECT * FROM correspondence")
        data = self.cursor.fetchall()
        
        if not data:
            messagebox.showinfo("معلومات", "لا توجد بيانات للتصدير")
            return
        
        # إنشاء إطار بيانات
        columns = [
            "ID", "النوع", "رقم المرجع", "التاريخ", "المرسل", "المرسل إليه",
            "الموضوع", "الإدارة", "الأولوية", "الحالة", "الملاحظات", "مسار المرفق"
        ]
        
        df = pd.DataFrame(data, columns=columns)
        
        # حفظ في ملف Excel
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )
        
        if file_path:
            try:
                df.to_excel(file_path, index=False)
                messagebox.showinfo("نجاح", f"تم التصدير بنجاح إلى:\n{file_path}")
            except Exception as e:
                messagebox.showerror("خطأ", f"حدث خطأ أثناء التصدير: {str(e)}")
    
    def change_appearance(self, new_appearance):
        if new_appearance == "فاتح":
            ctk.set_appearance_mode("light")
        elif new_appearance == "داكن":
            ctk.set_appearance_mode("dark")
        else:
            ctk.set_appearance_mode("system")

if __name__ == "__main__":
    app = CorrespondenceArchiveApp()
    app.mainloop()
