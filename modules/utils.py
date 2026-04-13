import os
import logging
import re
from typing import List, Tuple, Optional

def px_to_cm(px, dpi):
    return (px / dpi) * 2.54

def cm_to_pixels(cm: float, dpi: int = 72) -> int:
    """
    تحويل السنتيمتر إلى بكسل بناءً على دقة الطباعة (DPI).

    Args:
        cm: الطول بالسنتيمتر
        dpi: دقة الطباعة (نقطة لكل بوصة)، الافتراضي 300

    Returns:
        عدد البكسل (عدد صحيح)
    """
    # 1 بوصة = 2.54 سم
    inches = cm / 2.54
    pixels = int(inches * dpi)
    return pixels

def ensure_dir(path: str) -> None:
    """
    التأكد من وجود المجلد المحدد، وإنشائه إذا لم يكن موجودًا.

    Args:
        path: مسار المجلد
    """
    os.makedirs(path, exist_ok=True)

def setup_logging(debug: bool = False, log_file: Optional[str] = None) -> None:
    """
    إعداد نظام التسجيل (logging) للمشروع.

    Args:
        debug: إذا كان True، يتم ضبط المستوى على DEBUG، وإلا INFO
        log_file: مسار ملف لتسجيل السجلات (اختياري)
    """
    level = logging.DEBUG if debug else logging.INFO
    handlers = [logging.StreamHandler()]
    if log_file:
        ensure_dir(os.path.dirname(log_file))
        handlers.append(logging.FileHandler(log_file, encoding='utf-8'))

    logging.basicConfig(
        level=level,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S',
        handlers=handlers,
        force=True  # لإعادة تعيين أي إعدادات سابقة
    )

def sanitize_filename(filename: str, replacement: str = '_') -> str:
    """
    تنظيف اسم الملف من الأحغير المسموح بها في أنظمة الملفات.

    Args:
        filename: اسم الملف الأصلي
        replacement: الحرف الذي سيحل محل الأحرف غير المسموح بها

    Returns:
        اسم ملف آمن
    """
    # قائمة الأحرف غير المسموح بها في معظم أنظمة الملفات
    invalid_chars = r'[<>:"/\\|?*]'
    safe_name = re.sub(invalid_chars, replacement, filename)
    # إزالة المسافات الزائدة من الأطراف
    safe_name = safe_name.strip()
    # منع أسماء الملفات الفارغة
    if not safe_name:
        safe_name = 'unnamed'
    return safe_name

def parse_bold_segments(text: str) -> List[Tuple[str, bool]]:
    """
    يقسم النص إلى أجزاء عادية وأجزاء عريضة (الأرقام والوحدات).
    الأرقام وما يتبعها من وحدات (مثل Hz, K, mAh, سنوات) تعتبر عريضة.

    هذه الدالة مفيدة لمعالجة النصوص التي تحتوي على أرقام ووحدات
    وتريد تمييزها بخط عريض.

    Args:
        text: النص الكامل لعبارة واحدة

    Returns:
        قائمة من الأزواج (النص, هل هو عريض)
    """
    # وحدات محتملة بالعربية والإنجليزية
    units = r'(?:[a-zA-Z]+|سنة|سنوات|Hz|K|mAh|G|مم|سم|بكسل|ميجابكسل|جيجا|تيترا|شهور|أيام)'
    # النمط: رقم (قد يكون عشري) متبوعًا اختياريًا بمسافة ووحدة
    pattern = r'(\d+(?:\.\d+)?\s*(?:' + units + ')?)'

    segments = []
    last_end = 0
    for match in re.finditer(pattern, text):
        start, end = match.span()
        # جزء عادي قبل هذا المطابقة
        if start > last_end:
            segments.append((text[last_end:start], False))
        # الجزء المطابق (عريض)
        segments.append((text[start:end], True))
        last_end = end
    # باقي النص بعد آخر مطابقة
    if last_end < len(text):
        segments.append((text[last_end:], False))
    return segments

# اختبار بسيط للدوال (إذا تم تشغيل الملف مباشرة)
if __name__ == '__main__':
    # اختبار cm_to_pixels
    print(f"10 سم عند 300 DPI = {cm_to_pixels(10)} بكسل")

    # اختبار sanitize_filename
    bad_name = 'ملف: غير/مسموح? *|.txt'
    print(f"اسم آمن: {sanitize_filename(bad_name)}")

    # اختبار parse_bold_segments
    test_text = "بطارية 6500mAh فائقة النحافة وشاشة 120Hz"
    segments = parse_bold_segments(test_text)
    print("تقسيم النص إلى أجزاء:")
    for seg, bold in segments:
        print(f"  '{seg}' -> {'عريض' if bold else 'عادي'}")