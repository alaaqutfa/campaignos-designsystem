import pandas as pd
import yaml
import numpy as np

# قراءة ملف Excel
df = pd.read_excel("artboard_export.xlsx", sheet_name="ورقة1")

# استبدال القيم الخالية أو "-" بـ None
df = df.replace({"-": None, np.nan: None})

# تحويل الأعمدة المئوية إلى أرقام (مع التعامل مع الأخطاء)
pct_cols = ['width_pct', 'height_pct', 'x_offset_pct', 'y_offset_pct', 'font_size_pct']
for col in pct_cols:
    if col in df.columns:
        df[col] = pd.to_numeric(df[col], errors='coerce')

# دالة لتقريب أي قيمة رقمية إلى منزلتين عشريتين
def round_number(x):
    if isinstance(x, (float, np.floating)):
        return round(float(x), 2)
    elif isinstance(x, (int, np.integer)):
        return round(float(x), 2)  # لتحويل الأعداد الصحيحة إلى أرقام عشرية أيضاً
    return x

# دالة لتمرير على جميع العناصر وتطبيق التقريب
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

config_data = {}        # for whithout_shop
shop_config_data = {}   # for whith_shop

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
        # إزالة القيم الفارغة
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

    # معالجة images (مرتبة حسب index)
    images = group[group['element_type'] == 'image'].sort_values('index')
    if not images.empty:
        images_list = []
        for _, row in images.iterrows():
            image = {
                'anchor': row['anchor'],
                'image': row['image'],
                'width_pct': row['width_pct'],
                'height_pct': row['height_pct'],
                'x_offset_pct': row['x_offset_pct'],
                'y_offset_pct': row['y_offset_pct']
            }
            image = {k: v for k, v in image.items() if v is not None and not (isinstance(v, float) and np.isnan(v))}
            images_list.append(image)
        layout_dict['images'] = images_list

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

            # banner_logo
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

            # banner_shop_name
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

    # تطبيق التقريب العميق على كل القيم الرقمية
    layout_dict = deep_round(layout_dict)

    # إضافة الـ layout إلى الهدف
    target[layout] = layout_dict

# دالة لتصدير YAML مع تنسيق جيد
def write_yaml(file_path, data):
    with open(file_path, 'w', encoding='utf-8') as f:
        yaml.dump(data, f, allow_unicode=True, sort_keys=False, indent=2, default_flow_style=False)

# كتابة الملفين
write_yaml('done_config_layout.yaml', config_data)
write_yaml('done_shop_config_layout.yaml', shop_config_data)

print("Created Successfully: config_layout.yaml و shop_config_layout.yaml")