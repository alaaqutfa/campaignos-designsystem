import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
import logging
import os
import pandas as pd
import yaml
import numpy as np
from modules.excel_reader import read_excel
from modules.config_loader import load_sizes, load_layout
from modules.template_matcher import (
    match_template,
    extract_templates_from_config,
    get_tolerance,
)
from modules.renderer import DesignRenderer
from modules.utils import cm_to_pixels, ensure_dir, setup_logging, sanitize_filename


# ----------------------------------------------------------------------
# إعادة توجيه logging إلى Text widget
# ----------------------------------------------------------------------
class TextHandler(logging.Handler):
    """توجيه سجلات logging إلى عنصر ScrolledText"""

    def __init__(self, text_widget):
        super().__init__()
        self.text_widget = text_widget
        self.text_widget.configure(state="normal")
        self.text_widget.delete("1.0", tk.END)
        self.text_widget.configure(state="disabled")

    def emit(self, record):
        msg = self.format(record)

        def append():
            self.text_widget.configure(state="normal")
            self.text_widget.insert(tk.END, msg + "\n")
            self.text_widget.see(tk.END)
            self.text_widget.configure(state="disabled")

        self.text_widget.after(0, append)


# ----------------------------------------------------------------------
# دوال منطق التصميم (مأخوذة من main.py مع إضافة المعاملات الجديدة)
# ----------------------------------------------------------------------
def run_generation(
    input_file,
    sheet_name,
    output_dir,
    config_layout_path,
    config_sizes_path,
    dpi,
    preview,
    sample_n,
    test_sample_only,
    with_shop_flag,
    without_shop_flag,
    out_format,
    debug,
    log_file,
    safety_margin_cm=0.0,
    material="Mesh",
    add_label=False,
    skip_existing=False,  # <-- المعامل الجديد
    log_callback=None,
):
    """
    تنفيذ عملية إنشاء التصاميم.
        - safety_margin_cm: هامش الأمان بالسنتيمتر
        - material: نوع المادة (للتعامل مع هوامش مختلفة مستقبلاً)
        - add_label: إضافة تسمية تحتوي على رقم تسلسلي واسم الملف
        - skip_existing: إذا كان True، يتم تخطي الملفات الموجودة مسبقاً
    """
    # إعداد logging (with default file if not specified)
    if log_file:
        ensure_dir(os.path.dirname(log_file))
        logging.basicConfig(
            level=logging.DEBUG if debug else logging.INFO,
            format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
            handlers=[logging.FileHandler(log_file, encoding="utf-8")],
        )
    else:
        default_log = "error.log"
        logging.basicConfig(
            level=logging.DEBUG if debug else logging.INFO,
            format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
            handlers=[logging.FileHandler(default_log, encoding="utf-8")],
        )
    logger = logging.getLogger(__name__)

    # Add console handler to ensure logs appear in the GUI text widget
    if not any(isinstance(h, logging.StreamHandler) for h in logger.handlers):
        console = logging.StreamHandler()
        console.setLevel(logging.DEBUG if debug else logging.INFO)
        console.setFormatter(
            logging.Formatter("%(asctime)s - %(name)s - %(levelname)s - %(message)s")
        )
        logger.addHandler(console)

    try:
        # 1. تحميل التكوينات
        logger.info("Loading configuration files...")
        sizes_config = load_sizes(config_sizes_path)
        layout_config = load_layout(config_layout_path)
        shop_layout_config = load_layout("config_layout_shop.yaml")
        logger.info("   - Configuration files loaded successfully.")

        # استخراج القوالب
        templates = extract_templates_from_config(sizes_config)
        tolerance = get_tolerance(sizes_config)

        fmt = out_format.lower()
        if fmt in ("jpg", "jpeg"):
            fmt = "JPEG"
        elif fmt == "tiff":
            fmt = "TIFF"
        elif fmt == "png":
            fmt = "PNG"
        elif fmt == "pdf":
            fmt = "PDF"

        # حالة generate-test-sample الخاصة
        if test_sample_only:
            logger.info(
                "Test sample mode: generating designs for all templates (no Excel required)."
            )
            renderer = DesignRenderer(layout_config, assets_path="./assets")
            shop_renderer = DesignRenderer(shop_layout_config, assets_path="./assets")
            ensure_dir(output_dir)
            total_designs = 0
            for template in templates:
                base_filename = template["name"]
                shop_name = "Shop Name"
                template_name = template["name"]
                width_px = int(template["width"])
                height_px = int(template["height"])

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

                # بدون اسم المحل
                # img = renderer.render(
                #     width=width_px,
                #     height=height_px,
                #     template=template_name,
                #     shop_name=None,
                #     preview=preview,
                #     safety_margin=cm_to_pixels(safety_margin_cm, dpi),
                #     add_label=add_label,
                #     label_text=None,
                #     dpi=dpi,
                # )
                # save_design(img, "")

                # مع اسم المحل
                img = shop_renderer.render(
                    width=width_px,
                    height=height_px,
                    template=template_name,
                    shop_name=shop_name,
                    preview=preview,
                    safety_margin=cm_to_pixels(safety_margin_cm, dpi),
                    add_label=add_label,
                    label_text=None,
                    dpi=dpi,
                )
                save_design(img, " - Ar - 02")

            logger.info(
                f"Test sample generation completed. Generated {total_designs} designs in {output_dir}"
            )
            return

        # 2. قراءة ملف الإكسل
        logger.info(f"Reading Excel file: {input_file}")
        rows = read_excel(input_file, sheet_name=sheet_name)
        logger.info(f"  - Loaded {len(rows)} rows from Excel.")

        # 3. التحقق من صحة الصفوف وجمع الأخطاء
        errors = []
        valid_rows = []
        for idx, row in enumerate(rows, start=1):
            try:
                required = ["width", "height", "text"]
                for field in required:
                    if field not in row or row[field] is None:
                        raise ValueError(f"Missing required field '{field}'")
                width_cm = float(row["width"])
                height_cm = float(row["height"])
                if height_cm == 0:
                    raise ValueError("Height cannot be zero")
                if width_cm <= 0 or height_cm <= 0:
                    raise ValueError("Width and height must be positive")
                valid_rows.append((idx, row))
            except (ValueError, KeyError) as e:
                errors.append((idx, f"Row {idx}: {e}"))
                logger.error(f"Row {idx}: {e}")

        if errors:
            logger.error("Excel validation failed. Stopping generation.")
            for err in errors:
                logger.error(err)
            raise Exception("Excel validation failed")

        # 4. عينة محدودة
        if sample_n is not None:
            valid_rows = valid_rows[: int(sample_n)]
            logger.info(f"  - Using first {sample_n} rows for testing.")

        # 5. تهيئة المصممين
        renderer = DesignRenderer(layout_config, assets_path="./assets")
        shop_renderer = DesignRenderer(shop_layout_config, assets_path="./assets")

        # 6. التأكد من مجلد الإخراج
        ensure_dir(output_dir)

        # 7. معالجة الصفوف
        total_designs = 0
        shop_counter = {}

        for idx, row in valid_rows:
            logger.info(
                f"Processing row {idx}: {row.get('print_file_name', 'unknown')}"
            )

            width_cm = float(row["width"])
            height_cm = float(row["height"])
            rowMaterial = row["material"]
            shop_name = row.get("text", "")
            base_filename = sanitize_filename(
                row.get("print_file_name", f"design_{idx}")
            )

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

            # ========== هذا القسم يجب أن يكون داخل حلقة for ==========
            margin_cm = 0
            if dpi > 72:
                if rowMaterial in ["Backlit Lightbox"]:
                    margin_cm = 0.5
                elif rowMaterial in ["Flex Frame"]:
                    margin_cm = 7.5
                elif rowMaterial in ["Flex Lightbox"]:
                    margin_cm = 26.0
                elif rowMaterial in ["Parasol", "Parasol Lightbox", "Parasol lightbox"]:
                    margin_cm = 5.0
                elif rowMaterial in ["Foam Sticker", "Mesh", "Sticker"]:
                    margin_cm = 0.5
                else:
                    margin_cm = 0

            margin_px = cm_to_pixels(margin_cm, dpi)

            # print(f"Material : {rowMaterial}")

            label_text = None
            if add_label:
                if shop_name not in shop_counter:
                    shop_counter[shop_name] = 0
                shop_counter[shop_name] += 1
                label_text = (
                    f"[{shop_counter[shop_name]}] - {shop_name} - {base_filename}"
                )

            def save_design(img, suffix):
                nonlocal total_designs
                filename = f"{base_filename}{suffix}.{out_format}"
                filepath = os.path.join(output_dir, filename)

                # التحقق من وجود الملف إذا كان skip_existing == True
                if skip_existing and os.path.exists(filepath):
                    logger.info(f"Skipped existing file: {filepath}")
                    return

                if out_format == "png":
                    try:
                        img.save(filepath)
                    except Exception as e:
                        logger.error(f"Failed to save {filepath}: {e}")
                        return
                else:
                    try:
                        img.save(filepath, fmt, resolution=dpi)
                    except Exception as e:
                        logger.error(f"Failed to save {filepath}: {e}")
                        return
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
                    safety_margin=margin_px,
                    add_label=add_label,
                    label_text=label_text,
                    dpi=dpi,
                )
                save_design(img, " - 01")

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
                        safety_margin=margin_px,
                        add_label=add_label,
                        label_text=label_text,
                        dpi=dpi,
                    )
                    save_design(img, "")
                else:
                    logger.warning(
                        f"Row {idx} has no shop name for WITHSHOP version."
                    )

        logger.info(f"Done. Generated {total_designs} designs in: {output_dir}")

    except Exception as e:
        logger.exception(f"Unexpected error occurred: {e}")
        raise


# ----------------------------------------------------------------------
# دوال تحويل Excel إلى YAML (مأخوذة من excel_to_yaml.py مع إضافة قائمة الصور الموحدة)
# ----------------------------------------------------------------------
def convert_excel_to_yaml(
    excel_path, output_without_shop, output_with_shop, log_callback=None
):
    """
    تحويل Excel إلى ملفي YAML مع قائمة 'images' موحدة (مرتبة حسب index).
    """
    try:
        df = pd.read_excel(excel_path, sheet_name="artboard_export_need_check")
        df = df.replace({"-": None, np.nan: None})

        pct_cols = [
            "width_pct",
            "height_pct",
            "x_offset_pct",
            "y_offset_pct",
            "font_size_pct",
        ]
        for col in pct_cols:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors="coerce")

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

        grouped = df.groupby(["layout", "layout_type"])
        config_data = {}
        shop_config_data = {}

        for (layout, layout_type), group in grouped:
            # Accept both spellings
            lt = str(layout_type).strip().lower()
            if lt in ("without_shop", "whithout_shop"):
                target = config_data
            elif lt in ("with_shop", "whith_shop"):
                target = shop_config_data
            else:
                continue

            layout_dict = {}
            images = []

            # Collect all image elements
            for _, row in group.iterrows():
                elem_type = str(row["element_type"]).strip().lower()
                if elem_type in ["banner", "banner_logo", "banner_shop_name"]:
                    continue  # handled separately
                if pd.notna(row.get("image")) and row["image"]:
                    img_item = {
                        "image": row["image"],
                        "anchor": row.get("anchor", "top-left"),
                    }
                    if pd.notna(row.get("width_pct")):
                        img_item["width_pct"] = row["width_pct"]
                    if pd.notna(row.get("height_pct")):
                        img_item["height_pct"] = row["height_pct"]
                    if pd.notna(row.get("x_offset_pct")):
                        img_item["x_offset_pct"] = row["x_offset_pct"]
                    if pd.notna(row.get("y_offset_pct")):
                        img_item["y_offset_pct"] = row["y_offset_pct"]
                    # Add temporary index for sorting
                    if "index" in row and pd.notna(row["index"]):
                        img_item["_index"] = row["index"]
                    else:
                        img_item["_index"] = 0
                    images.append(img_item)

            # Sort by index
            images.sort(key=lambda x: x.get("_index", 0))
            for img in images:
                if "_index" in img:
                    del img["_index"]
            layout_dict["images"] = images

            # Handle banner for with_shop layouts
            if lt in ("with_shop", "whith_shop"):
                banner_row = None
                logo_row = None
                name_row = None
                for _, row in group.iterrows():
                    elem_type = str(row["element_type"]).strip().lower()
                    if elem_type == "banner":
                        banner_row = row
                    elif elem_type == "banner_logo":
                        logo_row = row
                    elif elem_type == "banner_shop_name":
                        name_row = row
                if banner_row is not None:
                    banner_dict = {
                        "width_pct": (
                            banner_row["width_pct"]
                            if pd.notna(banner_row["width_pct"])
                            else 100
                        ),
                        "height_pct": (
                            banner_row["height_pct"]
                            if pd.notna(banner_row["height_pct"])
                            else 20
                        ),
                        "background_color": (
                            banner_row["background_color"]
                            if pd.notna(banner_row["background_color"])
                            else "#000000"
                        ),
                    }
                    if logo_row is not None and pd.notna(logo_row.get("image")):
                        banner_dict["logo"] = {
                            "image": logo_row["image"],
                            "anchor": logo_row.get("anchor", "top-left"),
                            "width_pct": (
                                logo_row["width_pct"]
                                if pd.notna(logo_row["width_pct"])
                                else 10
                            ),
                            "x_offset_pct": (
                                logo_row["x_offset_pct"]
                                if pd.notna(logo_row["x_offset_pct"])
                                else 0
                            ),
                            "y_offset_pct": (
                                logo_row["y_offset_pct"]
                                if pd.notna(logo_row["y_offset_pct"])
                                else 0
                            ),
                        }
                    if name_row is not None:
                        banner_dict["shop_name"] = {
                            "font_size_pct": (
                                name_row["font_size_pct"]
                                if pd.notna(name_row["font_size_pct"])
                                else 5
                            ),
                            "color": (
                                name_row["background_color"]
                                if pd.notna(name_row["background_color"])
                                else "#FFFFFF"
                            ),
                            "x_offset_pct": (
                                name_row["x_offset_pct"]
                                if pd.notna(name_row["x_offset_pct"])
                                else 0
                            ),
                            "y_offset_pct": (
                                name_row["y_offset_pct"]
                                if pd.notna(name_row["y_offset_pct"])
                                else 0
                            ),
                        }
                    layout_dict["banner"] = banner_dict

            layout_dict = deep_round(layout_dict)
            target[layout] = layout_dict

        def write_yaml(file_path, data):
            with open(file_path, "w", encoding="utf-8") as f:
                yaml.dump(
                    data,
                    f,
                    allow_unicode=True,
                    sort_keys=False,
                    indent=2,
                    default_flow_style=False,
                )

        write_yaml(output_without_shop, config_data)
        write_yaml(output_with_shop, shop_config_data)

        if log_callback:
            log_callback(
                f"Created successfully: {output_without_shop} and {output_with_shop}"
            )
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
        self.root.title("CampaignOS Design Studio")
        self.root.geometry("1100x850")
        self.root.minsize(800, 600)
        if os.path.exists("icon.ico"):
            self.root.iconbitmap("icon.ico")

        self._setup_styles()

        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.design_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.design_tab, text="🎨 Generate Designs")
        self._make_scrollable(self.design_tab, self.build_design_tab)

        self.yaml_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.yaml_tab, text="📄 Convert Excel to YAML")
        self._make_scrollable(self.yaml_tab, self.build_yaml_tab)

        log_frame = ttk.LabelFrame(root, text="📋 Logs", padding=5)
        log_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))
        self.log_text = scrolledtext.ScrolledText(
            log_frame, height=12, state="normal", font=("Consolas", 9), wrap=tk.WORD
        )
        self.log_text.pack(fill=tk.BOTH, expand=True)

        self.log_handler = TextHandler(self.log_text)
        self.log_handler.setFormatter(
            logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")
        )
        logging.getLogger().addHandler(self.log_handler)
        logging.getLogger().setLevel(logging.INFO)

        self.status_var = tk.StringVar()
        self.status_var.set("Ready")
        status_bar = ttk.Label(
            root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W
        )
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)

    def _make_scrollable(self, parent, build_func):
        canvas = tk.Canvas(parent, highlightthickness=0, bg="#F8FAFC")
        scrollbar = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        build_func(scrollable_frame)

        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        canvas.bind_all("<MouseWheel>", _on_mousewheel)
        canvas.bind_all("<Button-4>", lambda e: canvas.yview_scroll(-1, "units"))
        canvas.bind_all("<Button-5>", lambda e: canvas.yview_scroll(1, "units"))

    def _setup_styles(self):
        style = ttk.Style()
        style.theme_use("clam")

        bg_light = "#F8FAFC"
        fg_dark = "#0F172A"
        primary = "#3B82F6"
        primary_hover = "#2563EB"
        border = "#E2E8F0"

        style.configure("TNotebook", background=bg_light)
        style.configure("TNotebook.Tab", padding=[12, 4], font=("Segoe UI", 10, "bold"))
        style.map(
            "TNotebook.Tab",
            background=[("selected", primary)],
            foreground=[("selected", "white")],
        )

        style.configure("TFrame", background=bg_light)
        style.configure(
            "TLabelframe",
            background=bg_light,
            foreground=fg_dark,
            bordercolor=border,
            relief=tk.RIDGE,
        )
        style.configure(
            "TLabelframe.Label",
            background=bg_light,
            foreground=primary,
            font=("Segoe UI", 10, "bold"),
        )

        style.configure(
            "TLabel", background=bg_light, foreground=fg_dark, font=("Segoe UI", 9)
        )
        style.configure(
            "TEntry", fieldbackground="white", bordercolor=border, focuscolor=primary
        )
        style.configure("TCombobox", fieldbackground="white", bordercolor=border)

        style.configure(
            "Accent.TButton",
            background=primary,
            foreground="white",
            borderwidth=0,
            focuscolor=primary,
        )
        style.map("Accent.TButton", background=[("active", primary_hover)])

        style.configure(
            "TButton",
            background=bg_light,
            foreground=fg_dark,
            borderwidth=1,
            relief=tk.RAISED,
        )
        style.map("TButton", background=[("active", "#E2E8F0")])

        style.configure("TCheckbutton", background=bg_light, foreground=fg_dark)
        style.configure("TRadiobutton", background=bg_light, foreground=fg_dark)

    def build_design_tab(self, parent):
        main_frame = ttk.Frame(parent)
        main_frame.pack(fill=tk.BOTH, expand=True)

        input_frame = ttk.LabelFrame(main_frame, text="📂 Input / Output", padding=10)
        input_frame.pack(fill=tk.X, pady=(0, 10))

        config_frame = ttk.LabelFrame(
            main_frame, text="⚙️ Configuration Files", padding=10
        )
        config_frame.pack(fill=tk.X, pady=(0, 10))

        output_frame = ttk.LabelFrame(main_frame, text="🖼️ Output Settings", padding=10)
        output_frame.pack(fill=tk.X, pady=(0, 10))

        extra_frame = ttk.LabelFrame(main_frame, text="🔧 Advanced Options", padding=10)
        extra_frame.pack(fill=tk.X, pady=(0, 10))

        # input_frame
        row = 0
        ttk.Label(input_frame, text="Excel File:").grid(
            row=row, column=0, sticky=tk.W, padx=5, pady=5
        )
        self.input_entry = ttk.Entry(input_frame, width=60)
        self.input_entry.grid(row=row, column=1, sticky=tk.EW, padx=5, pady=5)
        ttk.Button(input_frame, text="Browse", command=self.browse_input).grid(
            row=row, column=2, padx=5, pady=5
        )
        row += 1

        ttk.Label(input_frame, text="Sheet Name:").grid(
            row=row, column=0, sticky=tk.W, padx=5, pady=5
        )
        self.sheet_entry = ttk.Entry(input_frame, width=20)
        self.sheet_entry.insert(0, "Design")
        self.sheet_entry.grid(row=row, column=1, sticky=tk.W, padx=5, pady=5)
        row += 1

        ttk.Label(input_frame, text="Output Folder:").grid(
            row=row, column=0, sticky=tk.W, padx=5, pady=5
        )
        self.output_entry = ttk.Entry(input_frame, width=60)
        self.output_entry.insert(0, "./out")
        self.output_entry.grid(row=row, column=1, sticky=tk.EW, padx=5, pady=5)
        ttk.Button(input_frame, text="Browse", command=self.browse_output).grid(
            row=row, column=2, padx=5, pady=5
        )

        input_frame.columnconfigure(1, weight=1)

        # config_frame
        ttk.Label(config_frame, text="Config Layout:").grid(
            row=0, column=0, sticky=tk.W, padx=5, pady=5
        )
        self.config_layout_entry = ttk.Entry(config_frame, width=50)
        self.config_layout_entry.insert(0, "config_layout.yaml")
        self.config_layout_entry.grid(row=0, column=1, sticky=tk.EW, padx=5, pady=5)
        ttk.Button(
            config_frame,
            text="Browse",
            command=lambda: self.browse_file(self.config_layout_entry),
        ).grid(row=0, column=2, padx=5, pady=5)

        ttk.Label(config_frame, text="Config Sizes:").grid(
            row=1, column=0, sticky=tk.W, padx=5, pady=5
        )
        self.config_sizes_entry = ttk.Entry(config_frame, width=50)
        self.config_sizes_entry.insert(0, "config_sizes.yaml")
        self.config_sizes_entry.grid(row=1, column=1, sticky=tk.EW, padx=5, pady=5)
        ttk.Button(
            config_frame,
            text="Browse",
            command=lambda: self.browse_file(self.config_sizes_entry),
        ).grid(row=1, column=2, padx=5, pady=5)

        config_frame.columnconfigure(1, weight=1)

        # output_frame
        ttk.Label(output_frame, text="DPI:").grid(
            row=0, column=0, sticky=tk.W, padx=5, pady=5
        )
        self.dpi_entry = ttk.Entry(output_frame, width=10)
        self.dpi_entry.insert(0, "72")
        self.dpi_entry.grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)

        ttk.Label(output_frame, text="Format:").grid(
            row=0, column=2, sticky=tk.W, padx=15, pady=5
        )
        self.format_combo = ttk.Combobox(
            output_frame,
            values=["png", "pdf", "jpg", "tiff"],
            state="readonly",
            width=8,
        )
        self.format_combo.set("png")
        self.format_combo.grid(row=0, column=3, sticky=tk.W, padx=5, pady=5)

        self.preview_var = tk.BooleanVar()
        ttk.Checkbutton(
            output_frame, text="Preview (low resolution)", variable=self.preview_var
        ).grid(row=1, column=0, columnspan=2, sticky=tk.W, padx=5, pady=5)

        self.debug_var = tk.BooleanVar()
        ttk.Checkbutton(
            output_frame, text="Debug logging", variable=self.debug_var
        ).grid(row=1, column=2, columnspan=2, sticky=tk.W, padx=5, pady=5)

        output_frame.columnconfigure(1, weight=1)

        # extra_frame
        self.sample_var = tk.BooleanVar()
        self.sample_spin = ttk.Spinbox(
            extra_frame, from_=1, to=100, width=5, state="disabled"
        )
        ttk.Checkbutton(
            extra_frame,
            text="Generate sample (first N rows):",
            variable=self.sample_var,
            command=self.toggle_sample,
        ).grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.sample_spin.grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)

        self.test_sample_var = tk.BooleanVar()
        ttk.Checkbutton(
            extra_frame,
            text="Generate test sample (no Excel required)",
            variable=self.test_sample_var,
        ).grid(row=1, column=0, columnspan=2, sticky=tk.W, padx=5, pady=5)

        shop_frame = ttk.Frame(extra_frame)
        shop_frame.grid(row=2, column=0, columnspan=2, sticky=tk.W, padx=5, pady=5)
        self.with_shop_var = tk.BooleanVar()
        self.without_shop_var = tk.BooleanVar()
        ttk.Checkbutton(
            shop_frame,
            text="With shop name",
            variable=self.with_shop_var,
            command=self.toggle_shop,
        ).pack(side=tk.LEFT, padx=5)
        ttk.Checkbutton(
            shop_frame,
            text="Without shop name",
            variable=self.without_shop_var,
            command=self.toggle_shop,
        ).pack(side=tk.LEFT, padx=5)

        margin_frame = ttk.Frame(extra_frame)
        margin_frame.grid(row=3, column=0, columnspan=2, sticky=tk.W, padx=5, pady=5)
        self.margin_var = tk.BooleanVar()
        ttk.Checkbutton(
            margin_frame,
            text="Add safety margin",
            variable=self.margin_var,
            command=self.toggle_margin,
        ).pack(side=tk.LEFT)
        self.material_var = tk.StringVar(value="Mesh")
        self.material_combo = ttk.Combobox(
            margin_frame,
            textvariable=self.material_var,
            values=["Mesh", "Vinyl", "Paper"],
            state="readonly",
            width=8,
        )
        self.material_combo.pack(side=tk.LEFT, padx=10)
        ttk.Label(margin_frame, text="Margin (cm):").pack(side=tk.LEFT)
        self.margin_cm_entry = ttk.Entry(margin_frame, width=5, state="disabled")
        self.margin_cm_entry.insert(0, "0.5")
        self.margin_cm_entry.pack(side=tk.LEFT, padx=5)
        self.material_combo.state(["disabled"])

        self.label_var = tk.BooleanVar()
        ttk.Checkbutton(
            extra_frame,
            text="Add shop label (sequential number)",
            variable=self.label_var,
        ).grid(row=4, column=0, columnspan=2, sticky=tk.W, padx=5, pady=5)

        # خيار تخطي الملفات الموجودة
        self.skip_existing_var = tk.BooleanVar()
        ttk.Checkbutton(
            extra_frame,
            text="Skip existing files (do not overwrite)",
            variable=self.skip_existing_var,
        ).grid(row=5, column=0, columnspan=2, sticky=tk.W, padx=5, pady=5)

        logfile_frame = ttk.Frame(main_frame)
        logfile_frame.pack(fill=tk.X, pady=5)
        ttk.Label(logfile_frame, text="Log file (optional):").pack(side=tk.LEFT)
        self.logfile_entry = ttk.Entry(logfile_frame, width=50)
        self.logfile_entry.pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)
        ttk.Button(
            logfile_frame,
            text="Browse",
            command=lambda: self.browse_file(self.logfile_entry, save=True),
        ).pack(side=tk.LEFT, padx=5)

        run_frame = ttk.Frame(main_frame)
        run_frame.pack(fill=tk.X, pady=10)
        self.run_button = ttk.Button(
            run_frame,
            text="🚀 Run",
            command=self.start_generation,
            style="Accent.TButton",
        )
        self.run_button.pack()
        
        batch_frame = ttk.Frame(main_frame)
        batch_frame.pack(fill=tk.X, pady=5)
        ttk.Button(
            batch_frame,
            text="⚡ Batch Process All Provinces",
            command=self.start_batch,
            style="Accent.TButton",
        ).pack(pady=5)

        extra_frame.columnconfigure(1, weight=1)

    def build_yaml_tab(self, parent):
        main_frame = ttk.Frame(parent)
        main_frame.pack(fill=tk.BOTH, expand=True)

        frame = ttk.LabelFrame(main_frame, text="📁 Conversion", padding=10)
        frame.pack(fill=tk.X, pady=10)

        row = 0
        ttk.Label(frame, text="Excel File:").grid(
            row=row, column=0, sticky=tk.W, padx=5, pady=5
        )
        self.yaml_excel_entry = ttk.Entry(frame, width=60)
        self.yaml_excel_entry.grid(row=row, column=1, sticky=tk.EW, padx=5, pady=5)
        ttk.Button(frame, text="Browse", command=self.browse_yaml_excel).grid(
            row=row, column=2, padx=5, pady=5
        )
        row += 1

        ttk.Label(frame, text="Output YAML (without shop):").grid(
            row=row, column=0, sticky=tk.W, padx=5, pady=5
        )
        self.yaml_out_without_entry = ttk.Entry(frame, width=60)
        self.yaml_out_without_entry.insert(0, "done_config_layout.yaml")
        self.yaml_out_without_entry.grid(
            row=row, column=1, sticky=tk.EW, padx=5, pady=5
        )
        ttk.Button(
            frame,
            text="Browse",
            command=lambda: self.browse_file(self.yaml_out_without_entry, save=True),
        ).grid(row=row, column=2, padx=5, pady=5)
        row += 1

        ttk.Label(frame, text="Output YAML (with shop):").grid(
            row=row, column=0, sticky=tk.W, padx=5, pady=5
        )
        self.yaml_out_with_entry = ttk.Entry(frame, width=60)
        self.yaml_out_with_entry.insert(0, "done_shop_config_layout.yaml")
        self.yaml_out_with_entry.grid(row=row, column=1, sticky=tk.EW, padx=5, pady=5)
        ttk.Button(
            frame,
            text="Browse",
            command=lambda: self.browse_file(self.yaml_out_with_entry, save=True),
        ).grid(row=row, column=2, padx=5, pady=5)
        row += 1

        self.yaml_convert_button = ttk.Button(
            frame,
            text="🔄 Convert",
            command=self.start_yaml_conversion,
            style="Accent.TButton",
        )
        self.yaml_convert_button.grid(row=row, column=0, columnspan=3, pady=10)

        frame.columnconfigure(1, weight=1)

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
            filename = filedialog.asksaveasfilename(
                defaultextension=".yaml", filetypes=[("YAML files", "*.yaml")]
            )
        else:
            filename = filedialog.askopenfilename(
                filetypes=[("YAML files", "*.yaml"), ("All files", "*.*")]
            )
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
            self.sample_spin.state(["!disabled"])
        else:
            self.sample_spin.state(["disabled"])

    def toggle_shop(self):
        if self.with_shop_var.get() and self.without_shop_var.get():
            self.with_shop_var.set(False)
            self.without_shop_var.set(True)

    def toggle_margin(self):
        if self.margin_var.get():
            self.material_combo.state(["!disabled"])
            self.margin_cm_entry.config(state="normal")
        else:
            self.material_combo.state(["disabled"])
            self.margin_cm_entry.config(state="disabled")

    def start_generation(self):
        self.status_var.set("Processing...")
        self.run_button.config(state=tk.DISABLED)

        if self.test_sample_var.get():
            input_file = None
        else:
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

        margin_cm = 0.0
        material = "Mesh"
        if self.margin_var.get():
            try:
                margin_cm = float(self.margin_cm_entry.get())
            except ValueError:
                margin_cm = 0.0
            material = self.material_var.get()
        add_label = self.label_var.get()
        skip_existing = self.skip_existing_var.get()  # قراءة القيمة

        def target():
            try:
                run_generation(
                    input_file,
                    sheet,
                    output_dir,
                    config_layout,
                    config_sizes,
                    dpi,
                    preview,
                    sample_n,
                    test_sample,
                    with_shop,
                    without_shop,
                    out_format,
                    debug,
                    log_file,
                    safety_margin_cm=margin_cm,
                    material=material,
                    add_label=add_label,
                    skip_existing=skip_existing,
                )
            except Exception as e:
                logging.exception("Generation failed")
            finally:
                self.run_button.config(state=tk.NORMAL)

        threading.Thread(target=target, daemon=True).start()

    def start_batch(self):
        self.status_var.set("Processing...")
        base_output = self.output_entry.get()
        cities_dir = "Cities"
        if not os.path.isdir(cities_dir):
            messagebox.showerror("Error", f"Folder '{cities_dir}' not found.")
            return

        excel_files = [f for f in os.listdir(cities_dir) if f.endswith(".xlsx")]
        if not excel_files:
            messagebox.showinfo("Info", f"No Excel files found in '{cities_dir}'.")
            return

        sheet = self.sheet_entry.get()
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

        margin_cm = 0.0
        material = "Mesh"
        if self.margin_var.get():
            try:
                margin_cm = float(self.margin_cm_entry.get())
            except ValueError:
                margin_cm = 0.0
            material = self.material_var.get()
        add_label = self.label_var.get()
        skip_existing = self.skip_existing_var.get()  # قراءة القيمة

        def process_all():
            for file in excel_files:
                province = os.path.splitext(file)[0]
                output_dir = os.path.join(base_output, province)
                input_file = os.path.join(cities_dir, file)
                logging.info(f"Processing province: {province}")
                try:
                    run_generation(
                        input_file,
                        sheet,
                        output_dir,
                        config_layout,
                        config_sizes,
                        dpi,
                        preview,
                        sample_n,
                        test_sample,
                        with_shop,
                        without_shop,
                        out_format,
                        debug,
                        log_file,
                        safety_margin_cm=margin_cm,
                        material=material,
                        add_label=add_label,
                        skip_existing=skip_existing,
                    )
                except Exception as e:
                    logging.exception(f"Error processing {province}: {e}")
            logging.info("Batch processing completed.")

        self.run_button.config(state=tk.DISABLED)
        threading.Thread(target=process_all, daemon=True).start()

    def start_yaml_conversion(self):
        self.status_var.set("Processing...")
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


if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
