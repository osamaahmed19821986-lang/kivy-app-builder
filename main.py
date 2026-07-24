import sys
import os
import calendar
import traceback
from datetime import datetime

# استيراد أداة معرفة بيئة التشغيل من Kivy
from kivy.utils import platform

# --- 1. حارس الإقلاع: فحص المكتبات والخط قبل تشغيل الواجهة ---
IMPORT_ERRORS = []

try:
    import openpyxl
except Exception as e:
    IMPORT_ERRORS.append(f"openpyxl / et_xmlfile: {e}")

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

# استيراد مكتبات Kivy
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


# --- 2. البحث عن الخط العربي وتثبيته في الأندرويد ---
def find_font():
    """البحث عن ملف الخط العربي في كافة المسارات المحتملة على الموبايل"""
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
        # استبدال كافة أنماط Roboto بملف الخط العربي المعتمد
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
    IMPORT_ERRORS.append("⚠️ لم يتم العثور على ملف الخط (arial.ttf) داخل مجلد التطبيق!")

def ar(text):
    """دالة ضبط اتجاه وتشكيل الحروف العربية"""
    if text is None or text == "" or str(text).strip().lower() == "nan":
        return ""
    try:
        reshaped = arabic_reshaper.reshape(str(text))
        return get_display(reshaped)
    except Exception:
        return str(text)

def parse_date(val):
    if val is None or val == "" or str(val).strip().lower() == "nan":
        return None
    if isinstance(val, (datetime, datetime.date)):
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
        """طلب صلاحيات الوصول للملفات مع دعم أندرويد 14 الكامل"""
        if platform == 'android':
            try:
                from android.permissions import request_permissions, Permission
                from android import api_version
                from jnius import autoclass

                # إذا كان الجهاز يعمل بنظام أندرويد 13 أو 14 (API 33+)
                if api_version >= 33:
                    # طلب صلاحية اختيار الصور والوسائط للوجو
                    request_permissions([Permission.READ_MEDIA_IMAGES])
                    
                    # التحقق من صلاحية إدارة جميع الملفات لقراءة وحفظ الإكسيل والـ PDF
                    Environment = autoclass('android.os.Environment')
                    if not Environment.isExternalStorageManager():
                        Intent = autoclass('android.content.Intent')
                        Settings = autoclass('android.provider.Settings')
                        Uri = autoclass('android.net.Uri')
                        activity = autoclass('org.kivy.android.PythonActivity').mActivity
                        
                        # توجيه المستخدم لشاشة الإعدادات لمنح الصلاحية للتطبيق تلقائياً
                        intent = Intent(Settings.ACTION_MANAGE_APP_ALL_FILES_ACCESS_PERMISSION)
                        intent.setData(Uri.parse(f"package:{activity.getPackageName()}"))
                        activity.startActivity(intent)
                else:
                    # لأجهزة أندرويد 12 وما دون
                    request_permissions([
                        Permission.READ_EXTERNAL_STORAGE,
                        Permission.WRITE_EXTERNAL_STORAGE
                    ])
            except Exception as e:
                print(f"Error requesting permissions: {e}")

    def build(self):
        # عرض الشاشة التوضيحية بالأخطاء في حال نقص الخط أو المكتبات
        if IMPORT_ERRORS:
            err_box = BoxLayout(orientation="vertical", padding=20)
            msg = "⚠️ تعذر تشغيل التطبيق بسبب الأخطاء التالية:\n\n" + "\n".join(IMPORT_ERRORS)
            err_lbl = Label(text=msg, color=(1, 0.2, 0.2, 1), font_size="13sp")
            err_lbl.bind(size=err_lbl.setter('text_size'))
            err_box.add_widget(err_lbl)
            return err_box

        try:
            return self.create_main_ui()
        except Exception as e:
            err_box = BoxLayout(orientation="vertical", padding=20)
            err_lbl = Label(
                text=f"حدث خطأ أثناء بناء الواجهة:\n\n{traceback.format_exc()}",
                color=(1, 0.3, 0.3, 1),
                font_size="11sp"
            )
            err_lbl.bind(size=err_lbl.setter('text_size'))
            err_box.add_widget(err_lbl)
            return err_box

    def create_main_ui(self):
        self.title = "منظومة تنسيق رياض الأطفال"
        self.excel_path = ""
        self.logo_path = ""
        self.school_inputs = {}

        root = BoxLayout(orientation="vertical", padding=15, spacing=10)

        # الترويسة الرئيسية
        title_lbl = Label(
            text=ar("نظام التنسيق الإلكتروني المطور (أندرويد)"),
            font_size="16sp",
            bold=True,
            size_hint_y=None,
            height=35,
            color=(0.12, 0.24, 0.35, 1)
        )
        root.add_widget(title_lbl)

        # اختيار الملفات
        files_box = BoxLayout(orientation="vertical", spacing=5, size_hint_y=None, height=110)

        btn_excel = Button(text=ar("اختر ملف الإكسيل الرئيسي (.xlsx)"), size_hint_y=None, height=35)
        btn_excel.bind(on_press=lambda instance: self.open_file_picker("excel"))
        self.excel_status = Label(text=ar("لم يتم اختيار ملف الإكسيل"), color=(0.5, 0.5, 0.5, 1), font_size="11sp")

        btn_logo = Button(text=ar("اختر صورة الشعار / اللوجو"), size_hint_y=None, height=35)
        btn_logo.bind(on_press=lambda instance: self.open_file_picker("logo"))
        self.logo_status = Label(text=ar("لم يتم اختيار اللوجو (اختياري)"), color=(0.5, 0.5, 0.5, 1), font_size="11sp")

        files_box.add_widget(btn_excel)
        files_box.add_widget(self.excel_status)
        files_box.add_widget(btn_logo)
        files_box.add_widget(self.logo_status)
        root.add_widget(files_box)

        # إعدادات التاريخ والمرحلة
        date_box = BoxLayout(orientation="vertical", spacing=5, size_hint_y=None, height=80)
        date_box.add_widget(Label(text=ar("إعدادات التنسيق والسن المستهدف:"), bold=True, size_hint_y=None, height=20))

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

        labels_grid = GridLayout(cols=4, spacing=5, size_hint_y=None, height=18)
        labels_grid.add_widget(Label(text=ar("اليوم"), font_size="10sp"))
        labels_grid.add_widget(Label(text=ar("الشهر"), font_size="10sp"))
        labels_grid.add_widget(Label(text=ar("السنة"), font_size="10sp"))
        labels_grid.add_widget(Label(text=ar("المرحلة"), font_size="10sp"))
        date_box.add_widget(labels_grid)

        root.add_widget(date_box)

        # قائمة المدارس الديناميكية
        root.add_widget(Label(text=ar("الكثافات والحد الأدنى للسن لكل مدرسة:"), bold=True, size_hint_y=None, height=25))

        self.schools_layout = GridLayout(cols=1, spacing=5, size_hint_y=None)
        self.schools_layout.bind(minimum_height=self.schools_layout.setter("height"))

        scroll_view = ScrollView(size_hint=(1, 1))
        scroll_view.add_widget(self.schools_layout)
        root.add_widget(scroll_view)

        # أزرار التشغيل وحالة المعالجة
        bottom_box = BoxLayout(orientation="vertical", spacing=5, size_hint_y=None, height=80)
        self.run_btn = Button(
            text=ar("بدء معالجة التنسيق وتوليد التقارير"),
            background_color=(0.07, 0.3, 0.36, 1),
            disabled=True,
            size_hint_y=None,
            height=40,
        )
        self.run_btn.bind(on_press=self.start_coordination)

        self.status_txt = Label(
            text=ar("جاهز.. قم باختيار ملف الإكسيل أولاً."),
            color=(0.3, 0.3, 0.3, 1),
            font_size="11sp",
            size_hint_y=None,
            height=30,
        )

        bottom_box.add_widget(self.run_btn)
        bottom_box.add_widget(self.status_txt)
        root.add_widget(bottom_box)

        return root

    def open_file_picker(self, file_type):
        content = BoxLayout(orientation="vertical", spacing=10)
        
        # تحديد المسار الافتراضي للأندرويد
        if platform == 'android':
            default_path = "/storage/emulated/0/Download"
            if not os.path.exists(default_path):
                default_path = "/sdcard/Download"
        else:
            default_path = os.path.expanduser("~")

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

            wb = openpyxl.load_workbook(self.excel_path, data_only=True)
            sheet_schools_name = [s for s in wb.sheetnames if "المدارس" in s][0]
            ws_schools = wb[sheet_schools_name]

            headers = [cell.value for cell in ws_schools[1]]
            col_school_title_idx = None
            for idx, h in enumerate(headers):
                if h and "اسم المدرسة" in str(h):
                    col_school_title_idx = idx + 1
                    break

            unique_schools_set = set()
            for r in range(2, ws_schools.max_row + 1):
                val = ws_schools.cell(row=r, column=col_school_title_idx).value if col_school_title_idx else None
                cleaned = self.clean_school_name(val)
                if cleaned:
                    unique_schools_set.add(cleaned)

            unique_schools = sorted(list(unique_schools_set))

            for sch_name in unique_schools:
                row_box = BoxLayout(orientation="horizontal", size_hint_y=None, height=35, spacing=5)
                name_lbl = Label(text=ar(sch_name), size_hint_x=0.4, halign="right")
                name_lbl.bind(size=name_lbl.setter("text_size"))

                cap_tf = TextInput(text="45", multiline=False, input_filter="int", size_hint_x=0.3)
                age_tf = TextInput(text="4.0", multiline=False, size_hint_x=0.3)

                row_box.add_widget(name_lbl)
                row_box.add_widget(cap_tf)
                row_box.add_widget(age_tf)

                self.school_inputs[sch_name] = (cap_tf, age_tf)
                self.schools_layout.add_widget(row_box)

            self.run_btn.disabled = False
            self.status_txt.text = ar(f"تم تحميل عدد ({len(unique_schools)}) مدرسة بنجاح.")

        except Exception as ex:
            self.status_txt.text = ar(f"خطأ في تحميل المدارس: {str(ex)}")

    def calculate_exact_ymd(self, dob, calc_date):
        """حساب السن بالدقة (يوم-شهر-سنة)"""
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
        """توليد تقارير PDF المنسقة لكل مدرسة"""
        try:
            pdf = FPDF(orientation="P", unit="mm", format="A4")
            pdf.add_page()

            font_to_use = FONT_PATH
            if font_to_use:
                pdf.add_font("ArabicFont", "", font_to_use)
                pdf.set_font("ArabicFont", size=14)
            else:
                pdf.set_font("Helvetica", size=12)

            pdf.cell(190, 10, txt=ar("مديرية التربية و التعليم بأسوان"), ln=True, align="C")
            pdf.cell(190, 8, txt=ar(f"كشف التنسيق لمدرسة: {school_name} - المرحلة {stage_arabic}"), ln=True, align="C")
            pdf.ln(5)

            pdf.set_font("ArabicFont" if font_to_use else "Helvetica", size=10)
            pdf.cell(15, 8, txt=ar("م"), border=1, align="C")
            pdf.cell(65, 8, txt=ar("اسم الطالب"), border=1, align="C")
            pdf.cell(30, 8, txt=ar("تاريخ الميلاد"), border=1, align="C")
            pdf.cell(40, 8, txt=ar("السن (يوم-شهر-سنة)"), border=1, align="C")
            pdf.cell(40, 8, txt=ar("الملاحظات"), border=1, align="C")
            pdf.ln(8)

            for idx, st in enumerate(students_list, start=1):
                dob_str = st["dob_dt"].strftime("%Y-%m-%d") if st["dob_dt"] else ""
                y, m, d = self.calculate_exact_ymd(st["dob_dt"], calc_date)
                age_str = f"{d}-{m}-{y}" if y != "" else ""

                pdf.cell(15, 7, txt=str(idx), border=1, align="C")
                pdf.cell(65, 7, txt=ar(st["name"]), border=1, align="R")
                pdf.cell(30, 7, txt=dob_str, border=1, align="C")
                pdf.cell(40, 7, txt=ar(age_str), border=1, align="C")
                pdf.cell(40, 7, txt=ar(st["notes"]), border=1, align="R")
                pdf.ln(7)

            pdf.output(pdf_file_path)
        except Exception as e:
            print(f"PDF generation error: {e}")

    def start_coordination(self, instance):
        year_str = self.year_tf.text
        stage_num = int(self.stage_tf.text)
        stage_arabic = self.get_stage_arabic(stage_num)

        try:
            calc_date = datetime(int(year_str), int(self.month_tf.text), int(self.day_tf.text))
        except Exception:
            self.status_txt.text = ar("خطأ: يرجى التأكد من تاريخ احتساب السن!")
            return

        self.status_txt.text = ar(f"جاري الفرز والتسكين لطلاب المرحلة {stage_arabic}...")

        output_dir = os.path.dirname(self.excel_path)
        output_file = os.path.join(output_dir, f"منظومة_التنسيق_المرحلة_{stage_arabic}.xlsx")
        pdf_folder = os.path.join(output_dir, f"كشوف_المدارس_المرحلة_{stage_arabic}_PDF")
        os.makedirs(pdf_folder, exist_ok=True)

        wb = openpyxl.load_workbook(self.excel_path)
        sheet_students_name = [s for s in wb.sheetnames if "الطلاب" in s][0]
        ws_students = wb[sheet_students_name]

        headers = [cell.value for cell in ws_students[1]]
        headers_str = [str(h) if h is not None else "" for h in headers]

        def find_col_idx(condition_fn, default=None):
            for idx, h in enumerate(headers_str):
                if condition_fn(h):
                    return idx + 1
            return default

        col_dob_idx = find_col_idx(lambda h: "تاريخ الميلاد" in h)
        col_student_name_idx = find_col_idx(lambda h: "اسم الطالب" in h or ("الاسم" in h and "المدرسة" not in h), default=2)

        col_r1_name_idx = find_col_idx(lambda h: "رغبة (1)اسم" in h)
        col_r1_code_idx = find_col_idx(lambda h: "رغبة (1)م" in h)
        col_r2_name_idx = find_col_idx(lambda h: "رغبة (2)اسم" in h)
        col_r2_code_idx = find_col_idx(lambda h: "رغبة (2)م" in h)
        col_r3_name_idx = find_col_idx(lambda h: "رغبة (3)اسم" in h)
        col_r3_code_idx = find_col_idx(lambda h: "رغبة (3)م" in h)
        col_r4_name_idx = find_col_idx(lambda h: "المدرسة المتميزة" in h and "كود" not in h)
        col_r4_code_idx = find_col_idx(lambda h: "كود المدرسة المتميزة" in h)

        col_out_name_idx = find_col_idx(lambda h: "اسم" in h and "التسكين" in h)
        col_out_code_idx = find_col_idx(lambda h: "كود" in h and "التسكين" in h)

        col_notes_idx = find_col_idx(lambda h: "الملاحظات" in h)
        if not col_notes_idx:
            col_notes_idx = len(headers) + 1
            ws_students.cell(row=1, column=col_notes_idx, value="الملاحظات")

        students = []
        for r_idx in range(2, ws_students.max_row + 1):
            def get_val(col_idx):
                if col_idx and col_idx <= ws_students.max_column:
                    return ws_students.cell(row=r_idx, column=col_idx).value
                return None

            dob_val = get_val(col_dob_idx)
            dob_dt = parse_date(dob_val)

            st = {
                "row_idx": r_idx,
                "name": get_val(col_student_name_idx),
                "dob_val": dob_val,
                "dob_dt": dob_dt,
                "r1_name": get_val(col_r1_name_idx),
                "r1_code": get_val(col_r1_code_idx),
                "r2_name": get_val(col_r2_name_idx),
                "r2_code": get_val(col_r2_code_idx),
                "r3_name": get_val(col_r3_name_idx),
                "r3_code": get_val(col_r3_code_idx),
                "r4_name": get_val(col_r4_name_idx),
                "r4_code": get_val(col_r4_code_idx),
                "out_name": str(get_val(col_out_name_idx) or "").strip(),
                "out_code": get_val(col_out_code_idx),
                "notes": "" if stage_num == 1 else (get_val(col_notes_idx) or ""),
            }
            students.append(st)

        school_capacities = {}
        school_min_ages = {}
        school_last_dob = {}
        school_accepted_count = {}

        for sch_name, (cap_tf, age_tf) in self.school_inputs.items():
            c_name = self.clean_school_name(sch_name)
            try:
                school_capacities[c_name] = int(cap_tf.text)
                school_min_ages[c_name] = float(age_tf.text)
            except Exception:
                school_capacities[c_name] = 45
                school_min_ages[c_name] = 4.0
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

        for st in students_sorted:
            curr_alloc = st["out_name"]
            if stage_num > 1 and curr_alloc and curr_alloc not in ["قائمة الانتظار", "nan", ""]:
                continue

            dob_dt = st["dob_dt"]
            s_age = ((calc_date - dob_dt).days) / 365.25 if dob_dt else 0
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
                        school_min_ages[sch_name_str] = 4.0
                        school_last_dob[sch_name_str] = None
                        school_accepted_count[sch_name_str] = 0

                    if s_age < school_min_ages[sch_name_str]:
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

        for st in students:
            r_idx = st["row_idx"]
            ws_students.cell(row=r_idx, column=col_out_name_idx, value=st["out_name"])
            ws_students.cell(row=r_idx, column=col_out_code_idx, value=st["out_code"])
            ws_students.cell(row=r_idx, column=col_notes_idx, value=st["notes"])

        wb.save(output_file)

        grouped_students = {}
        for st in students:
            alloc = st["out_name"] or "غير مسكن"
            grouped_students.setdefault(alloc, []).append(st)

        for sch_title, st_list in grouped_students.items():
            safe_filename = f"كشف_{sch_title.replace(' ', '_')}_المرحلة_{stage_arabic}.pdf"
            pdf_out_path = os.path.join(pdf_folder, safe_filename)
            self.generate_pdf_report(sch_title, st_list, pdf_out_path, stage_arabic, calc_date)

        self.status_txt.text = ar(f"✅ تم الحفظ وتوليد كشوف PDF بنجاح في:\n{pdf_folder}")


if __name__ == "__main__":
    CoordinationKivyApp().run()
