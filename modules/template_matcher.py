import logging
from typing import List, Dict, Any

def match_template(ratio: float, templates: List[Dict[str, Any]], tolerance: float = 0.05) -> str:
    """
        Args:
            ratio: نسبة العرض/الارتفاع (W / H)
            templates: قائمة القوالب من ملف التكوين، كل عنصر يحتوي على 'name' و 'ratio'
            tolerance: أقصى فرق نسبي مسموح به لاعتبار التطابق دقيقًا

        Returns:
            اسم القالب الأقرب (مثل 'S2' أو 'default').

        تسجل تحذيرًا إذا كان الفرق بين النسبة المطلوبة وأقرب قالب > tolerance.
    """
    if not templates:
        logging.warning("There are no knowledge templates in the configuration file, use 'default'.")
        return 'default'

    # حساب الفرق المطلق لكل قالب
    best_match = None
    min_diff = float('inf')
    best_name = 'default'

    for tmpl in templates:
        tmpl_ratio = tmpl['ratio']
        diff = abs(ratio - tmpl_ratio)
        if diff < min_diff:
            min_diff = diff
            best_name = tmpl['name']
            best_match = tmpl

    # تحقق من tolerance
    if min_diff > tolerance:
        logging.warning(
            f"The ratio {ratio:.3f} does not exactly match any template (closest template: {best_name} with a difference of {min_diff:.3f} > {tolerance})."
        )
    else:
        logging.info(f"The ratio {ratio:.3f} matched the template {best_name} (difference {min_diff:.3f}).")

    return best_name

# دالة مساعدة لاستخراج القوالب من قاموس التكوين الكامل
def extract_templates_from_config(sizes_config: Dict[str, Any]) -> List[Dict[str, Any]]:
    """
        Args:
            sizes_config: المحتوى الكامل لملف التكوين (dict)

        Returns:
            قائمة القوالب (كل عنصر يحتوي على 'name' و 'ratio')
    """
    return sizes_config.get('templates', [])

def get_tolerance(sizes_config: Dict[str, Any]) -> float:
    """
        تستخرج قيمة التسامح من ملف التكوين.

        Args:
            sizes_config: المحتوى الكامل لملف التكوين

        Returns:
            قيمة التسامح (float)، الافتراضي 0.05
    """
    return sizes_config.get('tolerance', 0.05)