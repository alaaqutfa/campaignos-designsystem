import pandas as pd
import yaml
import numpy as np


def convert_excel_to_yaml(excel_path, output_without_shop, output_with_shop):
    df = pd.read_excel(excel_path, sheet_name="artboard_export_need_check")
    df = df.replace({"-": None, np.nan: None})

    # numeric columns
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

    def deep_round(obj):
        if isinstance(obj, dict):
            return {k: deep_round(v) for k, v in obj.items()}
        elif isinstance(obj, list):
            return [deep_round(item) for item in obj]
        elif isinstance(obj, (float, np.floating)):
            return round(obj, 2)
        elif isinstance(obj, (int, np.integer)):
            return round(float(obj), 2)
        else:
            return obj

    grouped = df.groupby(["layout", "layout_type"])
    config_data = {}
    shop_config_data = {}

    for (layout, layout_type), group in grouped:
        target = config_data if layout_type == "whithout_shop" else shop_config_data
        layout_dict = {}

        # Collect all images (background, mobile, logo, icons) into one list
        images = []

        # Process backgrounds
        for _, row in group[
            group["element_type"].isin(["background", "background2"])
        ].iterrows():
            images.append(
                {
                    "type": "image",
                    "image": row["image"],
                    "anchor": row.get("anchor", "top-left"),
                    "width_pct": row["width_pct"],
                    "height_pct": row["height_pct"],
                    "x_offset_pct": row["x_offset_pct"],
                    "y_offset_pct": row["y_offset_pct"],
                }
            )

        # Process mobiles
        for _, row in group[
            group["element_type"].isin(["mobile", "mobile2"])
        ].iterrows():
            images.append(
                {
                    "type": "image",
                    "image": row["image"],
                    "anchor": row["anchor"],
                    "width_pct": row["width_pct"],
                    "x_offset_pct": row["x_offset_pct"],
                    "y_offset_pct": row["y_offset_pct"],
                }
            )

        # Process logos (only for without_shop)
        if layout_type == "whithout_shop":
            for _, row in group[
                group["element_type"].isin(["logo", "logo2"])
            ].iterrows():
                images.append(
                    {
                        "type": "image",
                        "image": row["image"],
                        "anchor": row["anchor"],
                        "width_pct": row["width_pct"],
                        "x_offset_pct": row["x_offset_pct"],
                        "y_offset_pct": row["y_offset_pct"],
                    }
                )

        # Process icons
        for _, row in group[group["element_type"] == "icon"].iterrows():
            images.append(
                {
                    "type": "image",
                    "image": row["image"],
                    "anchor": row["anchor"],
                    "width_pct": row["width_pct"],
                    "height_pct": row["height_pct"],
                    "x_offset_pct": row["x_offset_pct"],
                    "y_offset_pct": row["y_offset_pct"],
                }
            )
        for _, row in group[group["element_type"] == "image"].iterrows():
            images.append(
                {
                    "type": "image",
                    "image": row["image"],
                    "anchor": row["anchor"],
                    "width_pct": row["width_pct"],
                    "height_pct": row["height_pct"],
                    "x_offset_pct": row["x_offset_pct"],
                    "y_offset_pct": row["y_offset_pct"],
                }
            )

        # Sort by index if column exists, else keep order
        if "index" in group.columns:
            images.sort(key=lambda x: x.get("index", 0))
        layout_dict["images"] = images

        # Banner handling for with_shop
        if layout_type == "whith_shop":
            banner_row = group[group["element_type"] == "banner"]
            if not banner_row.empty:
                banner = banner_row.iloc[0]
                banner_dict = {
                    "width_pct": banner["width_pct"],
                    "height_pct": banner["height_pct"],
                    "background_color": banner["background_color"],
                }
                # Add logo and shop name if present
                logo_row = group[group["element_type"] == "banner_logo"]
                if not logo_row.empty:
                    banner_dict["logo"] = {
                        "image": logo_row.iloc[0]["image"],
                        "anchor": logo_row.iloc[0]["anchor"],
                        "width_pct": logo_row.iloc[0]["width_pct"],
                        "x_offset_pct": logo_row.iloc[0]["x_offset_pct"],
                        "y_offset_pct": logo_row.iloc[0]["y_offset_pct"],
                    }
                name_row = group[group["element_type"] == "banner_shop_name"]
                if not name_row.empty:
                    banner_dict["shop_name"] = {
                        "font_size_pct": name_row.iloc[0]["font_size_pct"],
                        "color": name_row.iloc[0]["background_color"],
                        "x_offset_pct": name_row.iloc[0]["x_offset_pct"],
                        "y_offset_pct": name_row.iloc[0]["y_offset_pct"],
                    }
                layout_dict["banner"] = banner_dict

        layout_dict = deep_round(layout_dict)
        target[layout] = layout_dict

    def write_yaml(path, data):
        with open(path, "w", encoding="utf-8") as f:
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
