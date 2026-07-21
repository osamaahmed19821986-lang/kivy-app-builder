import os
import calendar
from datetime import datetime
import openpyxl

# مكتبات Kivy الأساسية
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

# مكتبات الـ PDF واللغة العربية
import arabic_reshaper
from bidi.algorithm import get_display
from reportlab.lib.pagesizes import A4
from reportlab.platypus import BaseDocTemplate, PageTemplate, Frame, Table, TableStyle, Spacer, PageBreak
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont


# --- 1. تحديد مسار الخط العربي المرفق بالمشروع وتصحيحه ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
FONT_PATH = os.path.join(BASE_DIR, "arial.ttf")

def init_arabic_font():
    if os.path.exists(FONT_PATH):
        LabelBase.register(name='Roboto', fn_regular=FONT_PATH)
    else:
        font_paths = [
            r"C:\Windows\Fonts\arial.ttf",
            r"C:\Windows\Fonts\tahoma.ttf",
            "arial.ttf",
            "tahoma.ttf",
            "/system/fonts/Roboto-Regular.ttf"
        ]
        for font in font_paths:
            if os.path.exists(font):
                LabelBase.register(name='Roboto', fn_regular=font)
                break

init_arabic_font()


# دالة مساعدة لضبط اتجاه وتشكيل الحروف العربية
def ar(text):
    if text is None or text == "" or str(text).strip().lower() == "nan":
        return ""
    try:
        reshaped = arabic_reshaper.reshape(str(text))
        return get_display(reshaped)
    except:
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
    except:
        pass
    return None


class CoordinationKivyApp(App):
    def build(self):
        self.title = "منظومة تنسيق رياض الأطفال - أندرويد أوفلاين"
        
        self.excel_path = ""
        self.logo_path = ""
        self.school_inputs = {}

        # الحاوية الرئيسية للتطبيق
        root = BoxLayout(orientation='vertical', padding=15, spacing=10)

        # 1. ترويسة التطبيق
        title_lbl = Label(
            text=ar("نظام التنسيق الإلكتروني المطور (إصدار أندرويد أوفلاين)"),
            font_size='18sp',
            bold=True,
            size_hint_y=None,
            height=40,
            color=(0.12, 0.24, 0.35, 1)
        )
        root.add_widget(title_lbl)

        # 2. قسم اختيار الملفات (Excel واللوجو)
        files_box = BoxLayout(orientation='vertical', spacing=8, size_hint_y=None, height=130)
        
        btn_excel = Button(text=ar("اختر ملف الإكسيل الرئيسي (.xlsx)"), size_hint_y=None, height=38)
        btn_excel.bind(on_press=lambda instance: self.open_file_picker('excel'))
        self.excel_status = Label(text=ar("لم يتم اختيار ملف الإكسيل"), color=(0.5, 0.5, 0.5, 1), font_size='12sp')

        btn_logo = Button(text=ar("اختر صورة الشعار / اللوجو"), size_hint_y=None, height=38)
        btn_logo.bind(on_press=lambda instance: self.open_file_picker('logo'))
        self.logo_status = Label(text=ar("لم يتم اختيار اللوجو (اختياري)"), color=(0.5, 0.5, 0.5, 1), font_size='12sp')

        files_box.add_widget(btn_excel)
        files_box.add_widget(self.excel_status)
        files_box.add_widget(btn_logo)
        files_box.add_widget(self.logo_status)
        root.add_widget(files_box)

        # 3. إعدادات تاريخ الاحتساب والمرحلة
        date_box = BoxLayout(orientation='vertical', spacing=5, size_hint_y=None, height=85)
        date_box.add_widget(Label(text=ar("إعدادات التنسيق والسن المستهدف:"), bold=True, size_hint_y=None, height=25))

        inputs_grid = GridLayout(cols=4, spacing=10, size_hint_y=None, height=35)
        self.day_tf = TextInput(text="1", multiline=False, input_filter='int')
        self.month_tf = TextInput(text="10", multiline=False, input_filter='int')
        self.year_tf = TextInput(text="2026", multiline=False, input_filter='int')
        self.stage_tf = TextInput(text="1", multiline=False, input_filter='int')

        inputs_grid.add_widget(self.day_tf)
        inputs_grid.add_widget(self.month_tf)
        inputs_grid.add_widget(self.year_tf)
        inputs_grid.add_widget(self.stage_tf)
        date_box.add_widget(inputs_grid)

        # عناوين حقول التاريخ
        labels_grid = GridLayout(cols=4, spacing=10, size_hint_y=None, height=20)
        labels_grid.add_widget(Label(text=ar("اليوم"), font_size='11sp'))
        labels_grid.add_widget(Label(text=ar("الشهر"), font_size='11sp'))
        labels_grid.add_widget(Label(text=ar("السنة"), font_size='11sp'))
        labels_grid.add_widget(Label(text=ar("المرحلة"), font_size='11sp'))
        date_box.add_widget(labels_grid)

        root.add_widget(date_box)

        # 4. قائمة المدارس الديناميكية داخل ScrollView
        root.add_widget(Label(text=ar("الكثافات والحد الأدنى للسن لكل مدرسة:"), bold=True, size_hint_y=None, height=30))
        
        self.schools_layout = GridLayout(cols=1, spacing=10, size_hint_y=None)
        self.schools_layout.bind(minimum_height=self.schools_layout.setter('height'))

        scroll_view = ScrollView(size_hint=(1, 1))
        scroll_view.add_widget(self.schools_layout)
        root.add_widget(scroll_view)

        # 5. زر التنسيق وشريط الحالة
        bottom_box = BoxLayout(orientation='vertical', spacing=8, size_hint_y=None, height=90)
        self.run_btn = Button(
            text=ar("بدء معالجة التنسيق وتوليد الكشوف والتقارير"),
            background_color=(0.07, 0.3, 0.36, 1),
            disabled=True,
            size_hint_y=None,
            height=45
        )
        self.run_btn.bind(on_press=self.start_coordination)

        self.status_txt = Label(
            text=ar("جاهز.. برجاء تحميل ملف الإكسيل أولاً."),
            color=(0.3, 0.3, 0.3, 1),
            font_size='13sp',
            size_hint_y=None,
            height=35
        )

        bottom_box.add_widget(self.run_btn)
        bottom_box.add_widget(self.status_txt)
        root.add_widget(bottom_box)

        return root

    # نافذة اختيار الملفات التفاعلية
    def open_file_picker(self, file_type):
        content = BoxLayout(orientation='vertical', spacing=10)
        filechooser = FileChooserListView(path=os.path.expanduser('~'))
        
        if file_type == 'excel':
            filechooser.filters = ['*.xlsx']
        elif file_type == 'logo':
            filechooser.filters = ['*.png', '*.jpg', '*.jpeg']

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
                if file_type == 'excel':
                    self.excel_path = selected
                    self.excel_status.text = ar(f"تم اختيار: {os.path.basename(selected)}")
                    self.load_schools()
                elif file_type == 'logo':
                    self.logo_path = selected
                    self.logo_status.text = ar(f"تم اختيار اللوجو: {os.path.basename(selected)}")
            popup.dismiss()

        select_btn.bind(on_press=on_select)
        cancel_btn.bind(on_press=popup.dismiss)
        popup.open()

    def clean_school_name(self, name):
        if name is None or str(name).strip().lower() == "nan":
            return ""
        return str(name).replace('\t', '').replace('\r', '').replace('\n', '').strip()

    def format_arabic_text(self, text):
        if text is None or text == "" or str(text).strip().lower() == "nan":
            return ""
        try:
            reshaped = arabic_reshaper.reshape(str(text))
            return get_display(reshaped)
        except:
            return str(text)

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
                row_box = BoxLayout(orientation='horizontal', size_hint_y=None, height=40, spacing=5)

                name_lbl = Label(text=ar(sch_name), size_hint_x=0.4, halign='right')
                name_lbl.bind(size=name_lbl.setter('text_size'))

                cap_tf = TextInput(text="45", multiline=False, input_filter='int', size_hint_x=0.3)
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
        except:
            return "", "", ""

    def start_coordination(self, instance):
        year_str = self.year_tf.text
        stage_num = int(self.stage_tf.text)
        stage_arabic = self.get_stage_arabic(stage_num)

        try:
            calc_date = datetime(int(year_str), int(self.month_tf.text), int(self.day_tf.text))
        except:
            self.status_txt.text = ar("خطأ: يرجى التأكد من تاريخ احتساب السن!")
            return

        self.status_txt.text = ar(f"جاري الفرز والتسكين لطلاب المرحلة {stage_arabic}...")

        output_dir = os.path.dirname(self.excel_path)
        output_file = os.path.join(output_dir, f"منظومة_التنسيق_المرحلة_{stage_arabic}.xlsx")
        pdf_folder = os.path.join(output_dir, f"كشوف_المدارس_المرحلة_{stage_arabic}_PDF")
        os.makedirs(pdf_folder, exist_ok=True)

        wb = openpyxl.load_workbook(self.excel_path)
        sheet_students_name = [s for s in wb.sheetnames if "الطلاب" in s][0]
        sheet_schools_name = [s for s in wb.sheetnames if "المدارس" in s][0]

        ws_students = wb[sheet_students_name]
        ws_schools = wb[sheet_schools_name]

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
                'row_idx': r_idx,
                'name': get_val(col_student_name_idx),
                'dob_val': dob_val,
                'dob_dt': dob_dt,
                'r1_name': get_val(col_r1_name_idx),
                'r1_code': get_val(col_r1_code_idx),
                'r2_name': get_val(col_r2_name_idx),
                'r2_code': get_val(col_r2_code_idx),
                'r3_name': get_val(col_r3_name_idx),
                'r3_code': get_val(col_r3_code_idx),
                'r4_name': get_val(col_r4_name_idx),
                'r4_code': get_val(col_r4_code_idx),
                'out_name': str(get_val(col_out_name_idx) or "").strip(),
                'out_code': get_val(col_out_code_idx),
                'notes': "" if stage_num == 1 else (get_val(col_notes_idx) or "")
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
            except:
                school_capacities[c_name] = 45
                school_min_ages[c_name] = 4.0
            school_last_dob[c_name] = None
            school_accepted_count[c_name] = 0

            if stage_num > 1:
                alloc_dobs = [s['dob_dt'] for s in students if self.clean_school_name(s['out_name']) == c_name and s['dob_dt'] is not None]
                if alloc_dobs:
                    school_last_dob[c_name] = max(alloc_dobs)

        for st in students:
            curr_alloc = st['out_name']
            if stage_num > 1 and curr_alloc and curr_alloc not in ["قائمة الانتظار", "nan", ""]:
                sch_name_str = self.clean_school_name(curr_alloc)
                if sch_name_str in school_accepted_count:
                    school_accepted_count[sch_name_str] += 1

        students_sorted = sorted(students, key=lambda x: x['dob_dt'] if x['dob_dt'] is not None else datetime.max)

        for st in students_sorted:
            curr_alloc = st['out_name']
            if stage_num > 1 and curr_alloc and curr_alloc not in ["قائمة الانتظار", "nan", ""]:
                continue

            dob_dt = st['dob_dt']
            if dob_dt:
                s_age = ((calc_date - dob_dt).days) / 365.25
            else:
                s_age = 0

            allocated = False
            rejected_by_age = False

            choices = [
                (st['r4_name'], st['r4_code']),
                (st['r1_name'], st['r1_code']),
                (st['r2_name'], st['r2_code']),
                (st['r3_name'], st['r3_code'])
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
                        st['out_name'] = sch_name_str
                        st['out_code'] = sch_code
                        school_capacities[sch_name_str] -= 1
                        school_accepted_count[sch_name_str] += 1
                        school_last_dob[sch_name_str] = dob_dt
                        allocated = True
                        break
                    elif school_capacities[sch_name_str] == 0 and school_last_dob[sch_name_str] == dob_dt and dob_dt is not None:
                        st['out_name'] = sch_name_str
                        st['out_code'] = sch_code
                        school_accepted_count[sch_name_str] += 1
                        st['notes'] = "مقبول تساوي سن"
                        allocated = True
                        break

            if not allocated:
                st['out_name'] = "قائمة الانتظار"
                st['out_code'] = 0
                st['notes'] = "استنفاذ رغبات اقل من السن المحدد" if rejected_by_age else "استنفاذ رغبات"

        for st in students:
            r_idx = st['row_idx']
            ws_students.cell(row=r_idx, column=col_out_name_idx, value=st['out_name'])
            ws_students.cell(row=r_idx, column=col_out_code_idx, value=st['out_code'])
            ws_students.cell(row=r_idx, column=col_notes_idx, value=st['notes'])

        wb.save(output_file)
        self.status_txt.text = ar(f"✅ تم الحفظ بنجاح داخل المجلد:\n{pdf_folder}")


if __name__ == "__main__":
    CoordinationKivyApp().run()