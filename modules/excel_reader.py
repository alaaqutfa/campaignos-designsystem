import pandas as pd
import os
from typing import List, Dict, Any


def read_excel(file_path: str, sheet_name: str = 'Sheet1') -> List[Dict[str, Any]]:
    """
        يقرأ ملف Excel ويحول كل صف إلى قاموس بالمفاتيح التالية:
            - city
            - address
            - shop_name
            - width
            - height
            - quantity
            - material
            - sqm
            - text
            - print_file_name
    """

    if not os.path.exists(file_path):
        raise FileNotFoundError(f"ملف Excel غير موجود: {file_path}")

    df = pd.read_excel(file_path, sheet_name=sheet_name, header=0, dtype=str)

    # توحيد أسماء الأعمدة بشكل ديناميكي
    actual_columns = {}
    for col in df.columns:
        col_lower = col.strip().lower()

        if 'city' in col_lower or 'مدينة' in col_lower:
            actual_columns[col] = 'city'

        elif 'address' in col_lower or 'عنوان' in col_lower:
            actual_columns[col] = 'address'

        elif 'shop' in col_lower or 'اسم المحل' in col_lower:
            actual_columns[col] = 'shop_name'

        elif (('w' in col_lower and 'cm' in col_lower) or 'العرض' in col_lower):
            actual_columns[col] = 'width'

        elif (('h' in col_lower and 'cm' in col_lower) or 'الارتفاع' in col_lower):
            actual_columns[col] = 'height'

        elif 'qty' in col_lower or 'كمية' in col_lower:
            actual_columns[col] = 'quantity'

        elif 'material' in col_lower or 'مادة' in col_lower:
            actual_columns[col] = 'material'

        elif 'sqm' in col_lower or 'متر مربع' in col_lower:
            actual_columns[col] = 'sqm'

        elif 'text' in col_lower or 'نص' in col_lower:
            actual_columns[col] = 'text'

        elif 'print' in col_lower or 'اسم الملف' in col_lower:
            actual_columns[col] = 'print_file_name'

    df.rename(columns=actual_columns, inplace=True)

    # التحقق من الأعمدة الأساسية
    required = ['city', 'address', 'shop_name', 'width', 'height', 'quantity', 'material']
    for req in required:
        if req not in df.columns:
            raise ValueError(
                f"العمود المطلوب '{req}' غير موجود. الأعمدة الحالية: {list(df.columns)}"
            )

    # تحويل القيم الرقمية
    df['width'] = pd.to_numeric(df['width'], errors='coerce')
    df['height'] = pd.to_numeric(df['height'], errors='coerce')
    df['quantity'] = pd.to_numeric(df['quantity'], errors='coerce').fillna(1).astype(int)

    # حساب SQM
    if 'sqm' not in df.columns or df['sqm'].isnull().all():
        df['sqm'] = (df['width'] * df['height'] * df['quantity']) / 10000
    else:
        df['sqm'] = pd.to_numeric(df['sqm'], errors='coerce')

    # إنشاء اسم ملف الطباعة
    if 'print_file_name' not in df.columns or df['print_file_name'].isnull().any():
        def build_name(row):
            try:
                return (
                    f"{row['city']} - {row['address']} - {row['shop_name']} - "
                    f"{int(row['width'])}x{int(row['height'])} - "
                    f"{row['material']} - QTY{int(row['quantity'])}"
                )
            except Exception:
                return "invalid_row"

        df['print_file_name'] = df.apply(build_name, axis=1)
    else:
        df['print_file_name'] = df['print_file_name'].fillna('')

    # تنظيف اسم الملف
    df['print_file_name'] = df['print_file_name'].apply(
        lambda x: str(x).replace('/', '_').replace('\\', '_').replace(':', '_')
    )

    return df.to_dict(orient='records')


if __name__ == '__main__':
    test_file = 'Cities/Anbar.xlsx'

    if os.path.exists(test_file):
        data = read_excel(test_file)
        for row in data:
            print(row)
    else:
        print(f"ملف الاختبار غير موجود: {test_file}")