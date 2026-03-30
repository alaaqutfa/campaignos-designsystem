import os
import yaml
import json
from typing import Any, Dict


def load_yaml(file_path: str) -> Dict[str, Any]:
    """
    تحميل ملف YAML وإرجاع محتواه كقاموس.
    إذا كان الامتداد .json، يتم استخدام json.load بدلاً من ذلك.
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"Configuration file not found: {file_path}")

    ext = os.path.splitext(file_path)[1].lower()
    try:
        with open(file_path, "r", encoding="utf-8") as f:
            if ext in (".yaml", ".yml"):
                return yaml.safe_load(f)
            elif ext == ".json":
                return json.load(f)
            else:
                raise ValueError(
                    f"Unsupported extension: {ext}. Use .yaml, .yml, or .json"
                )
    except (yaml.YAMLError, json.JSONDecodeError) as e:
        raise ValueError(f"Error reading configuration file {file_path}: {e}")


def load_sizes(file_path: str) -> Dict[str, Any]:
    """
    تحميل ملف تكوين أحجام القوالب (config_sizes.yaml).
    يحتوي على templates (قائمة) و tolerance.
    """
    return load_yaml(file_path)


def load_layout(file_path: str) -> Dict[str, Any]:
    """
    تحميل ملف تكوين تخطيط العناصر (config_layout.yaml).
    يحتوي على إعدادات لكل قالب (S1..S11) وقالب افتراضي.
    """
    return load_yaml(file_path)


# دالة مساعدة لدمج التكوين مع القيم الافتراضية (اختياري)
def merge_with_defaults(
    config: Dict[str, Any], defaults: Dict[str, Any]
) -> Dict[str, Any]:
    """
    دمج تكوين معين مع القيم الافتراضية (للاستخدام الداخلي).
    """
    result = defaults.copy()
    result.update(config)
    return result
