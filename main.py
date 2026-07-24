import sys
import os
import calendar
import traceback
import threading
from datetime import datetime, date

from kivy.utils import platform
from kivy.core.window import Window
from kivy.clock import Clock

# --- ضبط لون خلفية الشاشة ---
Window.clearcolor = (0.93, 0.94, 0.96, 1)

# --- 1. حارس الإقلاع: فحص المكتبات والخط ---
IMPORT_ERRORS = []

try:
    import openpyxl
except Exception as e:
    IMPORT_ERRORS.append(f"openpyxl: {e}")

try:
    import arabic_reshaper
except Exception as e:
    IMPORT_ERRORS.append(f"arabic_reshaper: {e}")

try:
    from bidi.algorithm import get_display
except Exception as e:
    IMPORT_ERRORS.append(f"python-bidi: {e}")

try:
    from fpdf import FPDF
except Exception as e:
    IMPORT_ERRORS.append(f"fpdf2: {e}")

try:
    from kivy.app import App
    from kivy.uix.boxlayout import BoxLayout
    from kivy.uix.gridlayout import GridLayout
    from kivy.uix.scrollview import ScrollView
    from kivy.uix.label import Label
    from kivy.uix.textinput import TextInput
    from kivy.uix.button import Button
    from kivy.uix.popup import Popup
    from kivy.uix.filechooser import FileChooserListView
    from kivy.core.text import LabelBase
except Exception as e:
    IMPORT_ERRORS.append(f"Kivy Core: {e}")


# --- 2. البحث عن الخط العربي وتسجيله ---
def find_font():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    candidates = [
        os.path.join(base_dir, "arial.ttf"),
        "arial.ttf",
        os.path.abspath("arial.ttf"),
        os.path.join(os.getcwd(), "arial.ttf"),
    ]
    for p in candidates:
        if p and os.path.exists(p):
            return p
    return None

FONT_PATH = find_font()

if FONT_PATH:
    try:
        LabelBase.register(
            name="Roboto",
            fn_regular=FONT_PATH,
            fn_bold=FONT_PATH,
            fn_italic=FONT_PATH,
            fn_bolditalic=FONT_PATH
        )
    except Exception as e:
        IMPORT_ERRORS.append(f"خطأ في تسجيل الخط: {e}")
else:
    IMPORT_ERRORS.append("⚠️ لم يتم العثور على ملف الخط (arial.ttf)!")

def ar(text):
    """ضبط اتجاه النص العربي"""
    if text is None or text == "" or str(text).strip().lower() == "nan":
        return ""
    try:
        reshaped = arabic_reshaper.reshape(str(text))
        return get_display(reshaped)
    except Exception:
        return str(text)

def parse_date(val):
    """تحويل القيم لتاريخ بشكل سريع وآمن"""
    if val is None or val == "" or str(val).strip().lower() == "nan":
        return None
    if isinstance(val, (datetime, date)):
        if isinstance(val, datetime):
            return val
        return datetime(val.year, val.month, val.day)
    try:
        s_val = str(val).strip()
        for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%Y/%m/%d"):
            try:
                return datetime.strptime(s_val, fmt)
            except ValueError:
                pass
    except Exception:
        pass
    return None


# --- 3. التطبيق الرئيسي ---
class CoordinationKivyApp(App):

    def on_start(self):
        """طلب صلاحيات الوصول للملفات في أندرويد"""
        if platform == 'android':
            try:
                from android.permissions import request_permissions, Permission
                from android import api_version
                from jnius import autoclass

                if api_version >= 33:
                    request_permissions([Permission.READ_MEDIA_IMAGES])
                    Environment = autoclass('android.os.Environment')
                    if not Environment.isExternalStorageManager():
                        Intent = autoclass('android.content.Intent')
                        Settings = autoclass('android.provider.Settings')
                        Uri = autoclass('android.net.Uri')
                        activity = autoclass('org.kivy.android.PythonActivity').mActivity
                        
                        intent = Intent(Settings.ACTION_MANAGE_APP_ALL_FILES_ACCESS_PERMISSION)
                        intent.setData(Uri.parse(f"package:{activity.getPackageName()}"))
                        activity.startActivity(intent)
                else:
                    request_permissions([
                        Permission.READ_EXTERNAL_STORAGE,
                        Permission.WRITE_EXTERNAL_STORAGE
                    ])
            except Exception as e:
                print(f"Error requesting permissions: {e}")

    def show_error_popup(self, title_text, err_text):
        content = BoxLayout(orientation="vertical", padding=10, spacing=10)
        lbl = Label(text=ar(err_text), size_hint_y=0.8, color=(1, 0.3, 0.3, 1))
        lbl.bind(size=lbl.setter('text_size'))
        btn = Button(text=ar("إغلاق"), size_hint_y=0.2)
        
        content.add_widget(lbl)
        content.add_widget(btn)
        
        popup = Popup(title=ar(title_text), content=content, size_hint=(0.9, 0.6))
        btn.bind(on_press=popup.dismiss)
        popup.open()

    def build(self):
        if IMPORT_ERRORS:
            err_box = BoxLayout(orientation="vertical", padding=20)
            msg = "⚠️ تعذر تشغيل التطبيق بسبب الأخطاء التالية:\n\n" + "\n".join(IMPORT_ERRORS)
            err_lbl = Label(text=msg, color=(1, 0.2, 0.2, 1), font_size="13sp")
            err_lbl.bind(size=err_lbl.setter('text_size'))
            err_box.add_widget(err_lbl)
            return err_box

        return self.create_main_ui()

    def create_main_ui(self):
        self.title = "منظومة تنسيق رياض الأطفال"
        self.excel_path = ""
        self.logo_path = ""
        self.school_inputs = {}
        self.loading_popup = None

        root = BoxLayout(orientation="vertical", padding=10, spacing=8)

        title_lbl = Label(
            text=ar("نظام التنسيق الإلكتروني المطور (أندرويد)"),
            font_size="15sp",
            bold=True,
            size_hint_y=None,
            height=30,
            color=(0.1, 0.2, 0.35, 1)
        )
        root.add_widget(title_lbl)

        files_box = BoxLayout(orientation="vertical", spacing=5, size_hint_y=None, height=150)

        btn_excel = Button(
            text=ar("📊 اختر ملف الإكسيل الرئيسي (.xlsx)"),
            size_hint_y=None,
            height=50,
            font_size="13sp",
            bold=True,
            background_color=(0.15, 0.55, 0.3, 1)
        )
        btn_excel.bind(on_press=lambda instance: self.open_file_picker("excel"))
        
        self.excel_status = Label(
            text=ar("لم يتم اختيار ملف الإكسيل"),
            color=(0.4, 0.4, 0.4, 1),
            font_size="11sp",
            size_hint_y=None,
            height=18
        )

        btn_logo = Button(
            text=ar("🖼️ اختر صورة الشعار / اللوجو"),
            size_hint_y=None,
            height=50,
            font_size="13sp",
            bold=True,
            background_color=(0.2, 0.45, 0.65, 1)
        )
        btn_logo.bind(on_press=lambda instance: self.open_file_picker("logo"))
        
        self.logo_status = Label(
            text=ar("لم يتم اختيار اللوجو (اختياري)"),
            color=(0.4, 0.4, 0.4, 1),
            font_size="11sp",
            size_hint_y=None,
            height=18
        )

        files_box.add_widget(btn_excel)
        files_box.add_widget(self.excel_status)
        files_box.add_widget(btn_logo)
        files_box.add_widget(self.logo_status)
        root.add_widget(files_box)

        # التاريخ والمرحلة
        date_box = BoxLayout(orientation="vertical", spacing=4, size_hint_y=None, height=75)
        date_box.add_widget(Label(text=ar("إعدادات تاريخ احتساب السن والمرحلة:"), bold=True, color=(0.1, 0.1, 0.1, 1), size_hint_y=None, height=18, font_size="12sp"))

        inputs_grid = GridLayout(cols=4, spacing=5, size_hint_y=None, height=30)
        self.day_tf = TextInput(text="1", multiline=False, input_filter="int")
        self.month_tf = TextInput(text="10", multiline=False, input_filter="int")
        self.year_tf = TextInput(text="2026", multiline=False, input_filter="int")
        self.stage_tf = TextInput(text="1", multiline=False, input_filter="int")

        inputs_grid.add_widget(self.day_tf)
        inputs_grid.add_widget(self.month_tf)
        inputs_grid.add_widget(self.year_tf)
        inputs_grid.add_widget(self.stage_tf)
        date_box.add_widget(inputs_grid)

        labels_grid = GridLayout(cols=4, spacing=5, size_hint_y=None, height=16)
        labels_grid.add_widget(Label(text=ar("اليوم"), font_size="10sp", color=(0.3, 0.3, 0.3, 1)))
        labels_grid.add_widget(Label(text=ar("الشهر"), font_size="10sp", color=(0.3, 0.3, 0.3, 1)))
        labels_grid.add_widget(Label(text=ar("السنة"), font_size="10sp", color=(0.3, 0.3, 0.3, 1)))
        labels_grid.add_widget(Label(text=ar("المرحلة"), font_size="10sp", color=(0.3, 0.3, 0.3, 1)))
        date_box.add_widget(labels_grid)

        root.add_widget(date_box)

        # جدول الكثافات
        root.add_widget(Label(text=ar("الكثافات والحد الأقصى لتاريخ الميلاد المقبول:"), bold=True, color=(0.1, 0.1, 0.1, 1), size_hint_y=None, height=22, font_size="12sp"))

        hdr_box = BoxLayout(orientation="horizontal", size_hint_y=None, height=25, spacing=5)
        hdr_box.add_widget(Label(text=ar("اسم المدرسة"), size_hint_x=0.5, bold=True, font_size="11sp", color=(0.2, 0.2, 0.2, 1)))
        hdr_box.add_widget(Label(text=ar("الكثافة"), size_hint_x=0.2, bold=True, font_size="11sp", color=(0.2, 0.2, 0.2, 1)))
        hdr_box.add_widget(Label(text=ar("أقصى تاريخ"), size_hint_x=0.3, bold=True, font_size="11sp", color=(0.2, 0.2, 0.2, 1)))
        root.add_widget(hdr_box)

        self.schools_layout = GridLayout(cols=1, spacing=6, size_hint_y=None)
        self.schools_layout.bind(minimum_height=self.schools_layout.setter("height"))

        scroll_view = ScrollView(size_hint=(1, 1))
        scroll_view.add_widget(self.schools_layout)
        root.add_widget(scroll_view)

        # الأزرار السفلية
        bottom_box = BoxLayout(orientation="vertical", spacing=5, size_hint_y=None, height=75)
        self.run_btn = Button(
            text=ar("🚀 بدء معالجة التنسيق وتوليد التقارير"),
            background_color=(0.07, 0.3, 0.36, 1),
            disabled=True,
            size_hint_y=None,
            height=45,
            font_size="13sp",
            bold=True
        )
        self.run_btn.bind(on_press=self.start_coordination_thread)

        self.status_txt = Label(
            text=ar("جاهز.. قم باختيار ملف الإكسيل أولاً."),
            color=(0.3, 0.3, 0.3, 1),
            font_size="11sp",
            size_hint_y=None,
            height=25,
        )

        bottom_box.add_widget(self.run_btn)
        bottom_box.add_widget(self.status_txt)
        root.add_widget(bottom_box)

        return root

    def open_file_picker(self, file_type):
        content = BoxLayout(orientation="vertical", spacing=10)
        
        default_path = "/storage/emulated/0/Download" if platform == 'android' else os.path.expanduser("~")
        if platform == 'android' and not os.path.exists(default_path):
            default_path = "/sdcard/Download"

        filechooser = FileChooserListView(path=default_path)

        if file_type == "excel":
            filechooser.filters = ["*.xlsx"]
        elif file_type == "logo":
            filechooser.filters = ["*.png", "*.jpg", "*.jpeg"]

        btn_box = BoxLayout(size_hint_y=None, height=40, spacing=10)
        select_btn = Button(text=ar("اختيار"))
        cancel_btn = Button(text=ar("إلغاء"))

        btn_box.add_widget(cancel_btn)
        btn_box.add_widget(select_btn)

        content.add_widget(filechooser)
        content.add_widget(btn_box)

        popup = Popup(title=ar("اختر الملف المطلوب"), content=content, size_hint=(0.95, 0.95))

        def on_select(instance):
            if filechooser.selection:
                selected = filechooser.selection[0]
                if file_type == "excel":
                    self.excel_path = selected
                    self.excel_status.text = ar(f"تم اختيار: {os.path.basename(selected)}")
                    self.load_schools()
                elif file_type == "logo":
                    self.logo_path = selected
                    self.logo_status.text = ar(f"تم اختيار اللوجو: {os.path.basename(selected)}")
            popup.dismiss()

        select_btn.bind(on_press=on_select)
        cancel_btn.bind(on_press=popup.dismiss)
        popup.open()

    def clean_school_name(self, name):
        if name is None or str(name).strip().lower() == "nan":
            return ""
        return str(name).replace("\t", "").replace("\r", "").replace("\n", "").strip()

    def get_stage_arabic(self, stage_num):
        stages = {"1": "الأولى", "2": "الثانية", "3": "الثالثة", "4": "الرابعة", "5": "الخامسة"}
        return stages.get(str(stage_num), f"الـ {stage_num}")

    def load_schools(self):
        try:
            self.schools_layout.clear_widgets()
            self.school_inputs.clear()

            wb = openpyxl.load_workbook(self.excel_path, read_only=True, data_only=True)
            school_sheets = [s for s in wb.sheetnames if "المدارس" in s]
            if not school_sheets:
                self.show_error_popup("خطأ في الملف", "لم يتم العثور على ورقة عمل تحتوى على كلمة 'المدارس'")
                wb.close()
                return
            
            ws_schools = wb[school_sheets[0]]
            rows = list(ws_schools.iter_rows(values_only=True))
            wb.close()

            if not rows:
                return

            headers = [str(h) if h is not None else "" for h in rows[0]]
            col_school_title_idx = None
            for idx, h in enumerate(headers):
                if "اسم المدرسة" in h:
                    col_school_title_idx = idx
                    break

            unique_schools_set = set()
            for r in rows[1:]:
                if col_school_title_idx is not None and col_school_title_idx < len(r):
                    cleaned = self.clean_school_name(r[col_school_title_idx])
                    if cleaned:
                        unique_schools_set.add(cleaned)

            unique_schools = sorted(list(unique_schools_set))

            for sch_name in unique_schools:
                row_box = BoxLayout(orientation="horizontal", size_hint_y=None, height=45, spacing=5)
                
                name_lbl = Label(
                    text=ar(sch_name),
                    size_hint_x=0.5,
                    halign="right",
                    valign="middle",
                    font_size="11sp",
                    color=(0.1, 0.1, 0.1, 1)
                )
                name_lbl.bind(size=lambda instance, value: setattr(instance, 'text_size', (value[0], None)))

                cap_tf = TextInput(text="45", multiline=False, input_filter="int", size_hint_x=0.2, font_size="11sp")
                age_tf = TextInput(text="2022-10-01", multiline=False, size_hint_x=0.3, font_size="11sp")

                row_box.add_widget(name_lbl)
                row_box.add_widget(cap_tf)
                row_box.add_widget(age_tf)

                self.school_inputs[sch_name] = (cap_tf, age_tf)
                self.schools_layout.add_widget(row_box)

            self.run_btn.disabled = False
            self.status_txt.text = ar(f"تم تحميل عدد ({len(unique_schools)}) مدرسة بنجاح.")

        except Exception as ex:
            self.show_error_popup("خطأ تحميل المدارس", str(ex))

    def calculate_exact_ymd(self, dob, calc_date):
        dob_dt = parse_date(dob)
        if not dob_dt:
            return "", "", ""
        try:
            y = calc_date.year - dob_dt.year
            m = calc_date.month - dob_dt.month
            d = calc_date.day - dob_dt.day
            if d < 0:
                prev_m = calc_date.month - 1 if calc_date.month > 1 else 12
                prev_y = calc_date.year if calc_date.month > 1 else calc_date.year - 1
                days_in_prev = calendar.monthrange(prev_y, prev_m)[1]
                d += days_in_prev
                m -= 1
            if m < 0:
                m += 12
                y -= 1
            return int(y), int(m), int(d)
        except Exception:
            return "", "", ""

    def generate_pdf_report(self, school_name, students_list, pdf_file_path, stage_arabic, calc_date):
        try:
            pdf = FPDF(orientation="P", unit="mm", format="A4")
            pdf.add_page()

            if FONT_PATH:
                pdf.add_font("ArabicFont", "", FONT_PATH)
                pdf.set_font("ArabicFont", size=14)
            else:
                pdf.set_font("Helvetica", size=12)

            pdf.cell(190, 10, txt=ar("مديرية التربية والتعليم بأسوان"), ln=True, align="C")
            pdf.cell(190, 8, txt=ar(f"كشف التنسيق لمدرسة: {school_name} - المرحلة {stage_arabic}"), ln=True, align="C")
            pdf.ln(5)

            pdf.set_font("ArabicFont" if FONT_PATH else "Helvetica", size=10)
            
            # عناوين الجدول من اليمين إلى اليسار
            pdf.cell(40, 8, txt=ar("الملاحظات"), border=1, align="C")
            pdf.cell(40, 8, txt=ar("السن (يوم-شهر-سنة)"), border=1, align="C")
            pdf.cell(30, 8, txt=ar("تاريخ الميلاد"), border=1, align="C")
            pdf.cell(65, 8, txt=ar("اسم الطالب"), border=1, align="C")
            pdf.cell(15, 8, txt=ar("م"), border=1, align="C")
            pdf.ln(8)

            for idx, st in enumerate(students_list, start=1):
                dob_str = st["dob_dt"].strftime("%Y-%m-%d") if st["dob_dt"] else ""
                y, m, d = self.calculate_exact_ymd(st["dob_dt"], calc_date)
                age_str = f"{d}-{m}-{y}" if y != "" else ""

                pdf.cell(40, 7, txt=ar(st["notes"]), border=1, align="R")
                pdf.cell(40, 7, txt=ar(age_str), border=1, align="C")
                pdf.cell(30, 7, txt=dob_str, border=1, align="C")
                pdf.cell(65, 7, txt=ar(st["name"]), border=1, align="R")
                pdf.cell(15, 7, txt=str(idx), border=1, align="C")
                pdf.ln(7)

            pdf.output(pdf_file_path)
        except Exception as e:
            print(f"PDF generation error: {e}")

    def start_coordination_thread(self, instance):
        self.run_btn.disabled = True
        
        content = BoxLayout(orientation="vertical", padding=15, spacing=10)
        self.loading_label = Label(
            text=ar("⏳ جاري بدء معالجة الملف..."),
            halign="center",
            color=(0.1, 0.6, 0.8, 1),
            font_size="13sp"
        )
        self.loading_label.bind(size=self.loading_label.setter('text_size'))
        content.add_widget(self.loading_label)
        
        self.loading_popup = Popup(
            title=ar("جاري معالجة البيانات"),
            content=content,
            size_hint=(0.85, 0.35),
            auto_dismiss=False
        )
        self.loading_popup.open()

        threading.Thread(target=self.run_coordination_process, daemon=True).start()

    def update_loading_status(self, text):
        Clock.schedule_once(lambda dt: setattr(self.loading_label, 'text', ar(text)))

    def run_coordination_process(self):
        try:
            year_str = self.year_tf.text
            stage_num = int(self.stage_tf.text)
            stage_arabic = self.get_stage_arabic(stage_num)

            try:
                calc_date = datetime(int(year_str), int(self.month_tf.text), int(self.day_tf.text))
            except Exception:
                Clock.schedule_once(lambda dt: self.finish_coordination_with_error("تاريخ غير صحيح", "يرجى كتابة التاريخ بشكل صحيح."))
                return

            output_dir = os.path.dirname(self.excel_path)
            output_file = os.path.join(output_dir, f"منظومة_التنسيق_المرحلة_{stage_arabic}.xlsx")
            pdf_folder = os.path.join(output_dir, f"كشوف_المدارس_المرحلة_{stage_arabic}_PDF")
            os.makedirs(pdf_folder, exist_ok=True)

            self.update_loading_status("📖 قراءة الملف من الذاكرة (سريع)...")
            
            wb_fast = openpyxl.load_workbook(self.excel_path, read_only=True, data_only=True)
            student_sheets = [s for s in wb_fast.sheetnames if "الطلاب" in s]
            if not student_sheets:
                wb_fast.close()
                Clock.schedule_once(lambda dt: self.finish_coordination_with_error("خطأ في الملف", "لم يتم العثور على ورقة 'الطلاب'"))
                return
            
            ws_fast = wb_fast[student_sheets[0]]
            all_rows = list(ws_fast.iter_rows(values_only=True))
            wb_fast.close()

            if not all_rows or len(all_rows) < 2:
                Clock.schedule_once(lambda dt: self.finish_coordination_with_error("تنبيه", "ورقة الطلاب فارغة!"))
                return

            headers = [str(h) if h is not None else "" for h in all_rows[0]]

            def find_col(condition_fn, default=None):
                for idx, h in enumerate(headers):
                    if condition_fn(h):
                        return idx
                return default

            col_dob_idx = find_col(lambda h: "تاريخ الميلاد" in h)
            col_student_name_idx = find_col(lambda h: "اسم الطالب" in h or ("الاسم" in h and "المدرسة" not in h), default=1)

            col_r1_name_idx = find_col(lambda h: "رغبة (1)اسم" in h)
            col_r1_code_idx = find_col(lambda h: "رغبة (1)م" in h)
            col_r2_name_idx = find_col(lambda h: "رغبة (2)اسم" in h)
            col_r2_code_idx = find_col(lambda h: "رغبة (2)م" in h)
            col_r3_name_idx = find_col(lambda h: "رغبة (3)اسم" in h)
            col_r3_code_idx = find_col(lambda h: "رغبة (3)م" in h)
            col_r4_name_idx = find_col(lambda h: "المدرسة المتميزة" in h and "كود" not in h)
            col_r4_code_idx = find_col(lambda h: "كود المدرسة المتميزة" in h)

            col_out_name_idx = find_col(lambda h: "اسم" in h and "التسكين" in h)
            col_out_code_idx = find_col(lambda h: "كود" in h and "التسكين" in h)
            col_notes_idx = find_col(lambda h: "الملاحظات" in h)

            students = []
            for r_idx, row_vals in enumerate(all_rows[1:], start=2):
                def get_v(idx):
                    if idx is not None and idx < len(row_vals):
                        return row_vals[idx]
                    return None

                dob_val = get_v(col_dob_idx)
                dob_dt = parse_date(dob_val)

                st = {
                    "row_idx": r_idx,
                    "name": get_v(col_student_name_idx),
                    "dob_val": dob_val,
                    "dob_dt": dob_dt,
                    "r1_name": get_v(col_r1_name_idx),
                    "r1_code": get_v(col_r1_code_idx),
                    "r2_name": get_v(col_r2_name_idx),
                    "r2_code": get_v(col_r2_code_idx),
                    "r3_name": get_v(col_r3_name_idx),
                    "r3_code": get_v(col_r3_code_idx),
                    "r4_name": get_v(col_r4_name_idx),
                    "r4_code": get_v(col_r4_code_idx),
                    "out_name": str(get_v(col_out_name_idx) or "").strip(),
                    "out_code": get_v(col_out_code_idx),
                    "notes": "" if stage_num == 1 else (get_v(col_notes_idx) or ""),
                }
                students.append(st)

            total_st_count = len(students)
            self.update_loading_status(f"🔄 جاري تسكين عدد ({total_st_count}) طالب...")

            school_capacities = {}
            school_min_dobs = {}
            school_last_dob = {}
            school_accepted_count = {}

            for sch_name, (cap_tf, age_tf) in self.school_inputs.items():
                c_name = self.clean_school_name(sch_name)
                try:
                    school_capacities[c_name] = int(cap_tf.text)
                except Exception:
                    school_capacities[c_name] = 45

                min_dob = parse_date(age_tf.text)
                if not min_dob:
                    min_dob = datetime(2022, 10, 1)
                school_min_dobs[c_name] = min_dob

                school_last_dob[c_name] = None
                school_accepted_count[c_name] = 0

                if stage_num > 1:
                    alloc_dobs = [s["dob_dt"] for s in students if self.clean_school_name(s["out_name"]) == c_name and s["dob_dt"] is not None]
                    if alloc_dobs:
                        school_last_dob[c_name] = max(alloc_dobs)

            for st in students:
                curr_alloc = st["out_name"]
                if stage_num > 1 and curr_alloc and curr_alloc not in ["قائمة الانتظار", "nan", ""]:
                    sch_name_str = self.clean_school_name(curr_alloc)
                    if sch_name_str in school_accepted_count:
                        school_accepted_count[sch_name_str] += 1

            students_sorted = sorted(students, key=lambda x: x["dob_dt"] if x["dob_dt"] is not None else datetime.max)

            for idx_st, st in enumerate(students_sorted, start=1):
                if idx_st % 50 == 0:
                    self.update_loading_status(f"⚙️ جاري معالجة الطالب ({idx_st} / {total_st_count})...")

                curr_alloc = st["out_name"]
                if stage_num > 1 and curr_alloc and curr_alloc not in ["قائمة الانتظار", "nan", ""]:
                    continue

                dob_dt = st["dob_dt"]
                allocated = False
                rejected_by_age = False

                choices = [
                    (st["r4_name"], st["r4_code"]),
                    (st["r1_name"], st["r1_code"]),
                    (st["r2_name"], st["r2_code"]),
                    (st["r3_name"], st["r3_code"]),
                ]

                for sch_name, sch_code in choices:
                    sch_name_str = self.clean_school_name(sch_name)
                    if sch_name_str and sch_name_str not in ["لا يوجد", "nan"]:
                        if sch_name_str not in school_capacities:
                            school_capacities[sch_name_str] = 45
                            school_min_dobs[sch_name_str] = datetime(2022, 10, 1)
                            school_last_dob[sch_name_str] = None
                            school_accepted_count[sch_name_str] = 0

                        if dob_dt and dob_dt > school_min_dobs[sch_name_str]:
                            rejected_by_age = True
                            continue

                        if school_capacities[sch_name_str] > 0:
                            st["out_name"] = sch_name_str
                            st["out_code"] = sch_code
                            school_capacities[sch_name_str] -= 1
                            school_accepted_count[sch_name_str] += 1
                            school_last_dob[sch_name_str] = dob_dt
                            allocated = True
                            break
                        elif school_capacities[sch_name_str] == 0 and school_last_dob[sch_name_str] == dob_dt and dob_dt is not None:
                            st["out_name"] = sch_name_str
                            st["out_code"] = sch_code
                            school_accepted_count[sch_name_str] += 1
                            st["notes"] = "مقبول تساوي سن"
                            allocated = True
                            break

                if not allocated:
                    st["out_name"] = "قائمة الانتظار"
                    st["out_code"] = 0
                    st["notes"] = "استنفاذ رغبات اقل من السن المحدد" if rejected_by_age else "استنفاذ رغبات"

            self.update_loading_status("💾 كتابة النتائج داخل الملف...")
            
            wb_write = openpyxl.load_workbook(self.excel_path)
            ws_write = wb_write[student_sheets[0]]

            headers_write = [cell.value for cell in ws_write[1]]
            
            def get_col_1based(name_fn):
                for i, h in enumerate(headers_write):
                    if h and name_fn(str(h)):
                        return i + 1
                return None

            c_out_n = get_col_1based(lambda h: "اسم" in h and "التسكين" in h)
            c_out_c = get_col_1based(lambda h: "كود" in h and "التسكين" in h)
            c_notes = get_col_1based(lambda h: "الملاحظات" in h)
            
            if not c_notes:
                c_notes = len(headers_write) + 1
                ws_write.cell(row=1, column=c_notes, value="الملاحظات")

            for st in students:
                r_idx = st["row_idx"]
                if c_out_n: ws_write.cell(row=r_idx, column=c_out_n, value=st["out_name"])
                if c_out_c: ws_write.cell(row=r_idx, column=c_out_c, value=st["out_code"])
                ws_write.cell(row=r_idx, column=c_notes, value=st["notes"])

            wb_write.save(output_file)
            wb_write.close()

            self.update_loading_status("📄 جاري توليد كشوف الـ PDF...")
            grouped_students = {}
            for st in students:
                alloc = st["out_name"] or "غير مسكن"
                grouped_students.setdefault(alloc, []).append(st)

            for sch_title, st_list in grouped_students.items():
                safe_filename = f"كشف_{sch_title.replace(' ', '_')}_المرحلة_{stage_arabic}.pdf"
                pdf_out_path = os.path.join(pdf_folder, safe_filename)
                self.generate_pdf_report(sch_title, st_list, pdf_out_path, stage_arabic, calc_date)

            Clock.schedule_once(lambda dt: self.finish_coordination_success(pdf_folder))

        except Exception as err:
            err_details = traceback.format_exc()
            Clock.schedule_once(lambda dt: self.finish_coordination_with_error("حدث خطأ أثناء التنسيق", f"تفاصيل الخطأ:\n{str(err)}\n\n{err_details}"))

    def finish_coordination_success(self, pdf_folder):
        if self.loading_popup:
            self.loading_popup.dismiss()
        self.run_btn.disabled = False
        self.status_txt.text = ar(f"✅ تم الانتهاء بنجاح وحفظ ملفات PDF في:\n{pdf_folder}")

    def finish_coordination_with_error(self, title, err_msg):
        if self.loading_popup:
            self.loading_popup.dismiss()
        self.run_btn.disabled = False
        self.show_error_popup(title, err_msg)


if __name__ == "__main__":
    CoordinationKivyApp().run()
