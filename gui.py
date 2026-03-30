import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
import logging
import sys
from io import StringIO
import os
import pandas as pd
import yaml
import numpy as np
from modules.excel_reader import read_excel
from modules.config_loader import load_sizes, load_layout
from modules.template_matcher import match_template, extract_templates_from_config, get_tolerance
from modules.renderer import DesignRenderer
from modules.utils import cm_to_pixels, ensure_dir, setup_logging, sanitize_filename
import argparse

# ----------------------------------------------------------------------
# إعادة توجيه logging إلى Text widget
# ----------------------------------------------------------------------
class TextHandler(logging.Handler):
    def __init__(self, text_widget):
        super().__init__()
        self.text_widget = text_widget

    def emit(self, record):
        msg = self.format(record)
        def append():
            self.text_widget.insert(tk.END, msg + '\n')
            self.text_widget.see(tk.END)
        self.text_widget.after(0, append)

# ----------------------------------------------------------------------
# دوال منطق التصميم (مأخوذة من main.py)
# ----------------------------------------------------------------------
def run_generation(
    input_file, sheet_name, output_dir, config_layout_path, config_sizes_path,
    dpi, preview, sample_n, test_sample_only, with_shop_flag, without_shop_flag,
    out_format, debug, log_file, log_callback=None
):
    """
    تنفيذ عملية إنشاء التصاميم باستخدام نفس منطق main.py.
    log_callback: دالة تستقبل النص لإضافته إلى سجل الواجهة (اختياري)
    """
    # إعداد logging
    if log_file:
        ensure_dir(os.path.dirname(log_file))
        logging.basicConfig(
            level=logging.DEBUG if debug else logging.INFO,
            format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
            handlers=[logging.FileHandler(log_file, encoding='utf-8')]
        )
    else:
        logging.basicConfig(
            level=logging.DEBUG if debug else logging.INFO,
            format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
        )
    logger = logging.getLogger(__name__)

    try:
        # 1. تحميل التكوينات
        logger.info("Loading configuration files...")
        sizes_config = load_sizes(config_sizes_path)
        layout_config = load_layout(config_layout_path)
        shop_layout_config = load_layout("config_layout_shop.yaml")  # مسار ثابت كما في main
        logger.info("   - Configuration files loaded successfully.")

        # 2. قراءة ملف الإكسل
        logger.info(f"Reading Excel file: {input_file}")
        rows = read_excel(input_file, sheet_name=sheet_name)
        logger.info(f"  - Loaded {len(rows)} rows from Excel.")

        # 3. عينة محدودة
        if sample_n is not None:
            rows = rows[:int(sample_n)]
            logger.info(f"  - Using first {sample_n} rows for testing.")

        # 4. تهيئة المصممين
        renderer = DesignRenderer(layout_config, assets_path="./assets")
        shop_renderer = DesignRenderer(shop_layout_config, assets_path="./assets")

        # 5. التأكد من مجلد الإخراج
        ensure_dir(output_dir)

        # 6. استخراج القوالب
        templates = extract_templates_from_config(sizes_config)
        tolerance = get_tolerance(sizes_config)

        # 7. تنسيق الإخراج
        fmt = out_format.lower()
        if fmt in ("jpg", "jpeg"):
            fmt = "JPEG"
        elif fmt == "tiff":
            fmt = "TIFF"
        elif fmt == "png":
            fmt = "PNG"
        elif fmt == "pdf":
            fmt = "PDF"

        # 8. معالجة الصفوف
        total_designs = 0

        # حالة generate-test-sample الخاصة
        if test_sample_only:
            for template in templates:
                base_filename = template['name']
                shop_name = "Shop Name"
                template_name = template['name']
                width_px = int(template['width'])
                height_px = int(template['height'])
                ratio = template['ratio']

                def save_design(img, suffix):
                    nonlocal total_designs
                    filename = f"{base_filename}{suffix}.{out_format}"
                    filepath = os.path.join(output_dir, filename)
                    if out_format == "png":
                        img.save(filepath)
                    else:
                        img.save(filepath, fmt, resolution=dpi)
                    logger.info(f"Saved: {filepath}")
                    total_designs += 1

                img = renderer.render(
                    width=width_px,
                    height=height_px,
                    template=template_name,
                    shop_name=None,
                    preview=preview,
                )
                save_design(img, " - Ar - 02")

                img = shop_renderer.render(
                    width=width_px,
                    height=height_px,
                    template=template_name,
                    shop_name=str(shop_name).strip(),
                    preview=preview,
                )
                save_design(img, " - Ar - 01")

        else:
            for idx, row in enumerate(rows, 1):
                logger.info(f"Processing row {idx}: {row.get('print_file_name', 'unknown')}")

                try:
                    width_cm = float(row["width"])
                    height_cm = float(row["height"])
                    shop_name = row.get("shop_name", "")
                    base_filename = sanitize_filename(row.get("print_file_name", f"design_{idx}"))
                except (ValueError, KeyError) as e:
                    logger.error(f"Error in row {idx}: {e}. Skipping.")
                    continue

                if height_cm == 0:
                    logger.error(f"Height is zero in row {idx}. Skipping.")
                    continue

                ratio = width_cm / height_cm
                template_name = match_template(ratio, templates, tolerance)

                width_px = cm_to_pixels(width_cm, dpi)
                height_px = cm_to_pixels(height_cm, dpi)

                generate_with_shop = False
                generate_without_shop = False

                if with_shop_flag:
                    generate_with_shop = True
                elif without_shop_flag:
                    generate_without_shop = True
                else:
                    generate_with_shop = True
                    generate_without_shop = True

                def save_design(img, suffix):
                    nonlocal total_designs
                    filename = f"{base_filename}{suffix}.{out_format}"
                    filepath = os.path.join(output_dir, filename)
                    if out_format == "png":
                        img.save(filepath)
                    else:
                        img.save(filepath, fmt, resolution=dpi)
                    logger.debug(f"Saved: {filepath}")
                    total_designs += 1

                # بدون اسم المحل
                if generate_without_shop:
                    logger.debug("Generating version WITHOUT shop name...")
                    img = renderer.render(
                        width=width_px,
                        height=height_px,
                        template=template_name,
                        shop_name=None,
                        preview=preview,
                    )
                    save_design(img, " - Ar - 02")

                # مع اسم المحل
                if generate_with_shop:
                    if shop_name and str(shop_name).strip():
                        logger.debug("Generating version WITH shop name...")
                        img = shop_renderer.render(
                            width=width_px,
                            height=height_px,
                            template=template_name,
                            shop_name=str(shop_name).strip(),
                            preview=preview,
                        )
                        save_design(img, " - Ar - 01")
                    else:
                        logger.warning(f"Row {idx} has no shop name for WITHSHOP version.")

        logger.info(f"Done. Generated {total_designs} designs in: {output_dir}")

    except Exception as e:
        logger.exception(f"Unexpected error occurred: {e}")
        raise

# ----------------------------------------------------------------------
# دوال تحويل Excel إلى YAML (مأخوذة من excel_to_yaml.py)
# ----------------------------------------------------------------------
def convert_excel_to_yaml(excel_path, output_without_shop, output_with_shop, log_callback=None):
    """
    تنفيذ تحويل excel_to_yaml.py مع إمكانية تمرير دالة لتسجيل السجلات.
    """
    try:
        # قراءة ملف Excel
        df = pd.read_excel(excel_path, sheet_name="ورقة1")

        # استبدال القيم الخالية أو "-" بـ None
        df = df.replace({"-": None, np.nan: None})

        # تحويل الأعمدة المئوية إلى أرقام
        pct_cols = ['width_pct', 'height_pct', 'x_offset_pct', 'y_offset_pct', 'font_size_pct']
        for col in pct_cols:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce')

        # دالة لتقريب الأرقام
        def round_number(x):
            if isinstance(x, (float, np.floating)):
                return round(float(x), 2)
            elif isinstance(x, (int, np.integer)):
                return round(float(x), 2)
            return x

        def deep_round(obj):
            if isinstance(obj, dict):
                return {k: deep_round(v) for k, v in obj.items()}
            elif isinstance(obj, list):
                return [deep_round(item) for item in obj]
            elif isinstance(obj, tuple):
                return tuple(deep_round(item) for item in obj)
            else:
                return round_number(obj)

        # تجميع البيانات حسب layout و layout_type
        grouped = df.groupby(['layout', 'layout_type'])
        config_data = {}
        shop_config_data = {}

        for (layout, layout_type), group in grouped:
            if layout_type == "whithout_shop":
                target = config_data
            elif layout_type == "whith_shop":
                target = shop_config_data
            else:
                continue

            layout_dict = {}

            # معالجة background و background2
            for _, row in group[group['element_type'].isin(['background', 'background2'])].iterrows():
                elem_type = row['element_type']
                key = 'background' if elem_type == 'background' else 'background2'
                layout_dict[key] = {
                    'type': 'image',
                    'image': row['image']
                }

            # معالجة mobile و mobile2
            for _, row in group[group['element_type'].isin(['mobile', 'mobile2'])].iterrows():
                elem_type = row['element_type']
                key = 'mobile_image' if elem_type == 'mobile' else 'mobile_image2'
                item = {
                    'anchor': row['anchor'],
                    'width_pct': row['width_pct'],
                    'x_offset_pct': row['x_offset_pct'],
                    'y_offset_pct': row['y_offset_pct']
                }
                item = {k: v for k, v in item.items() if v is not None and not (isinstance(v, float) and np.isnan(v))}
                layout_dict[key] = item

            # معالجة logo و logo2 (فقط في whithout_shop)
            if layout_type == "whithout_shop":
                for _, row in group[group['element_type'].isin(['logo', 'logo2'])].iterrows():
                    elem_type = row['element_type']
                    key = 'logo' if elem_type == 'logo' else 'logo2'
                    item = {
                        'anchor': row['anchor'],
                        'width_pct': row['width_pct'],
                        'x_offset_pct': row['x_offset_pct'],
                        'y_offset_pct': row['y_offset_pct']
                    }
                    item = {k: v for k, v in item.items() if v is not None and not (isinstance(v, float) and np.isnan(v))}
                    layout_dict[key] = item

            # معالجة icons (مرتبة حسب index)
            icons = group[group['element_type'] == 'icon'].sort_values('index')
            if not icons.empty:
                icons_list = []
                for _, row in icons.iterrows():
                    icon = {
                        'anchor': row['anchor'],
                        'image': row['image'],
                        'width_pct': row['width_pct'],
                        'height_pct': row['height_pct'],
                        'x_offset_pct': row['x_offset_pct'],
                        'y_offset_pct': row['y_offset_pct']
                    }
                    icon = {k: v for k, v in icon.items() if v is not None and not (isinstance(v, float) and np.isnan(v))}
                    icons_list.append(icon)
                layout_dict['icons'] = icons_list

            # معالجة banner (فقط في whith_shop)
            if layout_type == "whith_shop":
                banner_row = group[group['element_type'] == 'banner']
                banner_logo_row = group[group['element_type'] == 'banner_logo']
                banner_shop_row = group[group['element_type'] == 'banner_shop_name']

                if not banner_row.empty:
                    b = banner_row.iloc[0]
                    banner_dict = {
                        'width_pct': b['width_pct'],
                        'height_pct': b['height_pct'],
                        'background_color': b['background_color']
                    }
                    banner_dict = {k: v for k, v in banner_dict.items() if v is not None and not (isinstance(v, float) and np.isnan(v))}

                    if not banner_logo_row.empty:
                        bl = banner_logo_row.iloc[0]
                        logo_dict = {
                            'image': bl['image'],
                            'anchor': bl['anchor'],
                            'width_pct': bl['width_pct'],
                            'x_offset_pct': bl['x_offset_pct'],
                            'y_offset_pct': bl['y_offset_pct']
                        }
                        logo_dict = {k: v for k, v in logo_dict.items() if v is not None and not (isinstance(v, float) and np.isnan(v))}
                        banner_dict['logo'] = logo_dict

                    if not banner_shop_row.empty:
                        bs = banner_shop_row.iloc[0]
                        shop_dict = {
                            'font_size_pct': bs['font_size_pct'],
                            'color': bs['background_color'],
                            'x_offset_pct': bs['x_offset_pct'],
                            'y_offset_pct': bs['y_offset_pct']
                        }
                        shop_dict = {k: v for k, v in shop_dict.items() if v is not None and not (isinstance(v, float) and np.isnan(v))}
                        banner_dict['shop_name'] = shop_dict

                    layout_dict['banner'] = banner_dict

            layout_dict = deep_round(layout_dict)
            target[layout] = layout_dict

        # كتابة الملفين
        def write_yaml(file_path, data):
            with open(file_path, 'w', encoding='utf-8') as f:
                yaml.dump(data, f, allow_unicode=True, sort_keys=False, indent=2, default_flow_style=False)

        write_yaml(output_without_shop, config_data)
        write_yaml(output_with_shop, shop_config_data)

        if log_callback:
            log_callback(f"Created successfully: {output_without_shop} and {output_with_shop}")
        return True

    except Exception as e:
        if log_callback:
            log_callback(f"Error during conversion: {e}")
        return False

# ----------------------------------------------------------------------
# الواجهة الرسومية الرئيسية
# ----------------------------------------------------------------------
class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Design Generator & Excel to YAML Converter")
        self.root.geometry("900x700")

        # إنشاء دفتر علامات
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        # تبويب التصميمات
        self.design_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.design_frame, text="Generate Designs")

        # تبويب التحويل
        self.yaml_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.yaml_frame, text="Convert Excel to YAML")

        # بناء كل تبويب
        self.build_design_tab()
        self.build_yaml_tab()

        # منطقة عرض السجلات (مشتركة)
        log_frame = ttk.LabelFrame(root, text="Logs")
        log_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.log_text = scrolledtext.ScrolledText(log_frame, height=15, state='normal')
        self.log_text.pack(fill=tk.BOTH, expand=True)

        # إعداد logging لتوجيه إلى النص
        self.log_handler = TextHandler(self.log_text)
        self.log_handler.setFormatter(logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s'))
        logging.getLogger().addHandler(self.log_handler)
        logging.getLogger().setLevel(logging.INFO)

    def build_design_tab(self):
        """بناء عناصر تبويب Generate Designs"""
        main_frame = ttk.Frame(self.design_frame, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # صفوف الإدخالات
        row = 0
        # Input Excel
        ttk.Label(main_frame, text="Excel File:").grid(row=row, column=0, sticky=tk.W, pady=2)
        self.input_entry = ttk.Entry(main_frame, width=50)
        self.input_entry.grid(row=row, column=1, sticky=tk.W, pady=2)
        ttk.Button(main_frame, text="Browse", command=self.browse_input).grid(row=row, column=2, padx=5)
        row += 1

        # Sheet name
        ttk.Label(main_frame, text="Sheet Name:").grid(row=row, column=0, sticky=tk.W, pady=2)
        self.sheet_entry = ttk.Entry(main_frame, width=30)
        self.sheet_entry.insert(0, "Sheet1")
        self.sheet_entry.grid(row=row, column=1, sticky=tk.W, pady=2)
        row += 1

        # Output directory
        ttk.Label(main_frame, text="Output Folder:").grid(row=row, column=0, sticky=tk.W, pady=2)
        self.output_entry = ttk.Entry(main_frame, width=50)
        self.output_entry.insert(0, "./out")
        self.output_entry.grid(row=row, column=1, sticky=tk.W, pady=2)
        ttk.Button(main_frame, text="Browse", command=self.browse_output).grid(row=row, column=2, padx=5)
        row += 1

        # Config layout
        ttk.Label(main_frame, text="Config Layout:").grid(row=row, column=0, sticky=tk.W, pady=2)
        self.config_layout_entry = ttk.Entry(main_frame, width=50)
        self.config_layout_entry.insert(0, "config_layout.yaml")
        self.config_layout_entry.grid(row=row, column=1, sticky=tk.W, pady=2)
        ttk.Button(main_frame, text="Browse", command=lambda: self.browse_file(self.config_layout_entry)).grid(row=row, column=2, padx=5)
        row += 1

        # Config sizes
        ttk.Label(main_frame, text="Config Sizes:").grid(row=row, column=0, sticky=tk.W, pady=2)
        self.config_sizes_entry = ttk.Entry(main_frame, width=50)
        self.config_sizes_entry.insert(0, "config_sizes.yaml")
        self.config_sizes_entry.grid(row=row, column=1, sticky=tk.W, pady=2)
        ttk.Button(main_frame, text="Browse", command=lambda: self.browse_file(self.config_sizes_entry)).grid(row=row, column=2, padx=5)
        row += 1

        # DPI
        ttk.Label(main_frame, text="DPI:").grid(row=row, column=0, sticky=tk.W, pady=2)
        self.dpi_entry = ttk.Entry(main_frame, width=10)
        self.dpi_entry.insert(0, "72")
        self.dpi_entry.grid(row=row, column=1, sticky=tk.W, pady=2)
        row += 1

        # Format
        ttk.Label(main_frame, text="Output Format:").grid(row=row, column=0, sticky=tk.W, pady=2)
        self.format_combo = ttk.Combobox(main_frame, values=["png", "pdf", "jpg", "tiff"], state="readonly")
        self.format_combo.set("png")
        self.format_combo.grid(row=row, column=1, sticky=tk.W, pady=2)
        row += 1

        # Checkboxes (preview, debug)
        self.preview_var = tk.BooleanVar()
        ttk.Checkbutton(main_frame, text="Preview (low resolution)", variable=self.preview_var).grid(row=row, column=0, columnspan=2, sticky=tk.W, pady=2)
        row += 1

        self.debug_var = tk.BooleanVar()
        ttk.Checkbutton(main_frame, text="Debug logging", variable=self.debug_var).grid(row=row, column=0, columnspan=2, sticky=tk.W, pady=2)
        row += 1

        # Generate sample
        self.sample_var = tk.BooleanVar()
        self.sample_spin = ttk.Spinbox(main_frame, from_=1, to=100, width=5)
        ttk.Checkbutton(main_frame, text="Generate sample (first N rows):", variable=self.sample_var, command=self.toggle_sample).grid(row=row, column=0, sticky=tk.W, pady=2)
        self.sample_spin.grid(row=row, column=1, sticky=tk.W, pady=2)
        self.sample_spin.state(['disabled'])
        row += 1

        # Generate test sample (special mode)
        self.test_sample_var = tk.BooleanVar()
        ttk.Checkbutton(main_frame, text="Generate test sample (only without shop name)", variable=self.test_sample_var).grid(row=row, column=0, columnspan=2, sticky=tk.W, pady=2)
        row += 1

        # With/Without shop
        self.with_shop_var = tk.BooleanVar()
        self.without_shop_var = tk.BooleanVar()
        ttk.Checkbutton(main_frame, text="With shop name", variable=self.with_shop_var, command=self.toggle_shop).grid(row=row, column=0, sticky=tk.W, pady=2)
        ttk.Checkbutton(main_frame, text="Without shop name", variable=self.without_shop_var, command=self.toggle_shop).grid(row=row, column=1, sticky=tk.W, pady=2)
        row += 1

        # Log file
        ttk.Label(main_frame, text="Log file (optional):").grid(row=row, column=0, sticky=tk.W, pady=2)
        self.logfile_entry = ttk.Entry(main_frame, width=40)
        self.logfile_entry.grid(row=row, column=1, sticky=tk.W, pady=2)
        ttk.Button(main_frame, text="Browse", command=lambda: self.browse_file(self.logfile_entry, save=True)).grid(row=row, column=2, padx=5)
        row += 1

        # Run button
        self.run_button = ttk.Button(main_frame, text="Run", command=self.start_generation)
        self.run_button.grid(row=row, column=0, columnspan=3, pady=10)

    def build_yaml_tab(self):
        """بناء عناصر تبويب Convert Excel to YAML"""
        main_frame = ttk.Frame(self.yaml_frame, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        row = 0
        ttk.Label(main_frame, text="Excel File:").grid(row=row, column=0, sticky=tk.W, pady=2)
        self.yaml_excel_entry = ttk.Entry(main_frame, width=50)
        self.yaml_excel_entry.grid(row=row, column=1, sticky=tk.W, pady=2)
        ttk.Button(main_frame, text="Browse", command=self.browse_yaml_excel).grid(row=row, column=2, padx=5)
        row += 1

        ttk.Label(main_frame, text="Output YAML (without shop):").grid(row=row, column=0, sticky=tk.W, pady=2)
        self.yaml_out_without_entry = ttk.Entry(main_frame, width=50)
        self.yaml_out_without_entry.insert(0, "done_config_layout.yaml")
        self.yaml_out_without_entry.grid(row=row, column=1, sticky=tk.W, pady=2)
        ttk.Button(main_frame, text="Browse", command=lambda: self.browse_file(self.yaml_out_without_entry, save=True)).grid(row=row, column=2, padx=5)
        row += 1

        ttk.Label(main_frame, text="Output YAML (with shop):").grid(row=row, column=0, sticky=tk.W, pady=2)
        self.yaml_out_with_entry = ttk.Entry(main_frame, width=50)
        self.yaml_out_with_entry.insert(0, "done_shop_config_layout.yaml")
        self.yaml_out_with_entry.grid(row=row, column=1, sticky=tk.W, pady=2)
        ttk.Button(main_frame, text="Browse", command=lambda: self.browse_file(self.yaml_out_with_entry, save=True)).grid(row=row, column=2, padx=5)
        row += 1

        self.yaml_convert_button = ttk.Button(main_frame, text="Convert", command=self.start_yaml_conversion)
        self.yaml_convert_button.grid(row=row, column=0, columnspan=3, pady=10)

    # دوال مساعدة للتنقل
    def browse_input(self):
        filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if filename:
            self.input_entry.delete(0, tk.END)
            self.input_entry.insert(0, filename)

    def browse_output(self):
        folder = filedialog.askdirectory()
        if folder:
            self.output_entry.delete(0, tk.END)
            self.output_entry.insert(0, folder)

    def browse_file(self, entry, save=False):
        if save:
            filename = filedialog.asksaveasfilename(defaultextension=".yaml", filetypes=[("YAML files", "*.yaml")])
        else:
            filename = filedialog.askopenfilename(filetypes=[("YAML files", "*.yaml"), ("All files", "*.*")])
        if filename:
            entry.delete(0, tk.END)
            entry.insert(0, filename)

    def browse_yaml_excel(self):
        filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if filename:
            self.yaml_excel_entry.delete(0, tk.END)
            self.yaml_excel_entry.insert(0, filename)

    def toggle_sample(self):
        if self.sample_var.get():
            self.sample_spin.state(['!disabled'])
        else:
            self.sample_spin.state(['disabled'])

    def toggle_shop(self):
        # جعل الاختيار حصرياً تقريباً (يمكن أن يكون كلاهما False في البداية)
        if self.with_shop_var.get() and self.without_shop_var.get():
            # إذا حاول المستخدم تفعيلهما معاً، نلغي الآخر
            if self.with_shop_var.get():
                self.without_shop_var.set(False)
            else:
                self.with_shop_var.set(False)

    def start_generation(self):
        """تشغيل عملية التصميم في thread منفصل"""
        # تعطيل الزر مؤقتاً
        self.run_button.config(state=tk.DISABLED)

        # جمع المعاملات
        input_file = self.input_entry.get()
        if not input_file:
            messagebox.showerror("Error", "Please select an Excel file.")
            self.run_button.config(state=tk.NORMAL)
            return
        sheet = self.sheet_entry.get()
        output_dir = self.output_entry.get()
        config_layout = self.config_layout_entry.get()
        config_sizes = self.config_sizes_entry.get()
        try:
            dpi = int(self.dpi_entry.get())
        except ValueError:
            dpi = 72
        preview = self.preview_var.get()
        sample_n = self.sample_spin.get() if self.sample_var.get() else None
        test_sample = self.test_sample_var.get()
        with_shop = self.with_shop_var.get()
        without_shop = self.without_shop_var.get()
        out_format = self.format_combo.get()
        debug = self.debug_var.get()
        log_file = self.logfile_entry.get() if self.logfile_entry.get() else None

        # تشغيل في thread
        def target():
            try:
                run_generation(
                    input_file, sheet, output_dir, config_layout, config_sizes,
                    dpi, preview, sample_n, test_sample, with_shop, without_shop,
                    out_format, debug, log_file
                )
            except Exception as e:
                logging.exception("Generation failed")
            finally:
                self.run_button.config(state=tk.NORMAL)

        threading.Thread(target=target, daemon=True).start()

    def start_yaml_conversion(self):
        """تشغيل تحويل Excel إلى YAML في thread"""
        self.yaml_convert_button.config(state=tk.DISABLED)
        excel_path = self.yaml_excel_entry.get()
        out_without = self.yaml_out_without_entry.get()
        out_with = self.yaml_out_with_entry.get()
        if not excel_path:
            messagebox.showerror("Error", "Please select an Excel file.")
            self.yaml_convert_button.config(state=tk.NORMAL)
            return

        def target():
            def log(msg):
                logging.info(msg)
            success = convert_excel_to_yaml(excel_path, out_without, out_with, log)
            if success:
                logging.info("Conversion completed successfully.")
            else:
                logging.error("Conversion failed.")
            self.yaml_convert_button.config(state=tk.NORMAL)

        threading.Thread(target=target, daemon=True).start()

# ----------------------------------------------------------------------
# بدء التشغيل
# ----------------------------------------------------------------------
if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()