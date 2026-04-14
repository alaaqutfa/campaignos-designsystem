import os
from PIL import Image, ImageDraw, ImageFont
import arabic_reshaper
from bidi.algorithm import get_display
import logging
from typing import Dict, Any, List, Tuple, Optional
from modules.utils import cm_to_pixels

logger = logging.getLogger(__name__)


class DesignRenderer:
    def __init__(self, layout_config: Dict[str, Any], assets_path: str = "./assets"):
        self.config = layout_config
        self.assets = assets_path
        self.fonts_cache = {}

    def _reshape(self, text: str) -> str:
        reshaped = arabic_reshaper.reshape(text)
        return get_display(reshaped)

    def _get_font(self, font_name: str, size: int) -> ImageFont.FreeTypeFont:
        cache_key = f"{font_name}_{size}"
        if cache_key in self.fonts_cache:
            return self.fonts_cache[cache_key]
        font_path = os.path.join(self.assets, "fonts", font_name)
        if not os.path.exists(font_path):
            try:
                font = ImageFont.truetype("arial.ttf", size)
            except:
                font = ImageFont.load_default()
        else:
            font = ImageFont.truetype(font_path, size)
        self.fonts_cache[cache_key] = font
        return font

    def _draw_image_layer(self, img, layer, width, height):
        img_path = os.path.join(self.assets, layer.get("image", "missing.png"))
        if not os.path.exists(img_path):
            return

        Image.MAX_IMAGE_PIXELS = None
        orig_img = Image.open(img_path).convert("RGBA")
        orig_w, orig_h = orig_img.size

        # أبعاد الإطار المطلوب من YAML
        target_w = None
        target_h = None
        if "width_pct" in layer:
            target_w = int(layer["width_pct"] / 100 * width)
        if "height_pct" in layer:
            target_h = int(layer["height_pct"] / 100 * height)

        fit = layer.get("fit", "fill").lower()  # القيمة الافتراضية cover

        # الحالة 1: لدينا عرض وارتفاع -> نطبق استراتيجية fit
        if target_w is not None and target_h is not None:
            if fit == "contain":
                scale = min(target_w / orig_w, target_h / orig_h)
                new_w = int(orig_w * scale)
                new_h = int(orig_h * scale)
                scaled_img = orig_img.resize((new_w, new_h), Image.Resampling.LANCZOS)
                canvas = Image.new("RGBA", (target_w, target_h), (0, 0, 0, 0))
                left = (target_w - new_w) // 2
                top = (target_h - new_h) // 2
                canvas.paste(scaled_img, (left, top))
                layer_img = canvas

            elif fit == "fill":
                # تشويه النسبة لملء الإطار
                layer_img = orig_img.resize((target_w, target_h), Image.Resampling.LANCZOS)

            elif fit == "full-width":
                scale = target_w / orig_w
                new_w = target_w
                new_h = int(orig_h * scale)
                layer_img = orig_img.resize((new_w, new_h), Image.Resampling.LANCZOS)

            else:  # cover (الافتراضي)
                scale = max(target_w / orig_w, target_h / orig_h)
                new_w = int(orig_w * scale)
                new_h = int(orig_h * scale)
                scaled_img = orig_img.resize((new_w, new_h), Image.Resampling.LANCZOS)
                left = (new_w - target_w) // 2
                top = (new_h - target_h) // 2
                cropped = scaled_img.crop((left, top, left + target_w, top + target_h))
                layer_img = cropped

        # الحالة 2: عرض فقط -> نحافظ على النسبة
        elif target_w is not None:
            new_h = int(target_w * orig_h / orig_w)
            layer_img = orig_img.resize((target_w, new_h), Image.Resampling.LANCZOS)

        # الحالة 3: ارتفاع فقط -> نحافظ على النسبة
        elif target_h is not None:
            new_w = int(target_h * orig_w / orig_h)
            layer_img = orig_img.resize((new_w, target_h), Image.Resampling.LANCZOS)

        else:
            layer_img = orig_img

        # رسم الطبقة (نفس الكود القديم)
        w_layer, h_layer = layer_img.size
        anchor = layer.get("anchor", "top-left")
        x_offset = int(layer.get("x_offset_pct", 0) / 100 * width)
        y_offset = int(layer.get("y_offset_pct", 0) / 100 * height)

        if anchor == "right-center":
            x = width - w_layer + x_offset
            y = (height - h_layer) // 2 + y_offset
        elif anchor == "left-center":
            x = x_offset
            y = (height - h_layer) // 2 + y_offset
        elif anchor == "center":
            x = (width - w_layer) // 2 + x_offset
            y = (height - h_layer) // 2 + y_offset
        elif anchor == "top-right":
            x = width - w_layer + x_offset
            y = y_offset
        elif anchor == "top-left":
            x = x_offset
            y = y_offset
        elif anchor == "bottom-right":
            x = width - w_layer + x_offset
            y = height - h_layer + y_offset
        elif anchor == "bottom-left":
            x = x_offset
            y = height - h_layer + y_offset
        else:
            x = x_offset
            y = y_offset

        img.paste(layer_img, (x, y), layer_img)

    def _draw_background(
        self, img: Image, bg_cfg: Dict[str, Any], canvas_width: int, canvas_height: int
    ):
        bg_type = bg_cfg.get("type", "color")
        if bg_type == "color":
            color = bg_cfg.get("color", "#FFFFFF")
            img.paste(color, [0, 0, canvas_width, canvas_height])
        elif bg_type == "image":
            img_path = os.path.join(self.assets, bg_cfg.get("image", "background.png"))
            if os.path.exists(img_path):
                Image.MAX_IMAGE_PIXELS = None
                bg_img = Image.open(img_path).convert("RGB")
                bg_img = bg_img.resize(
                    (canvas_width, canvas_height), Image.Resampling.LANCZOS
                )
                img.paste(bg_img, (0, 0))
            else:
                img.paste("#FFFFFF", [0, 0, canvas_width, canvas_height])

    def _draw_banner(
        self,
        draw,
        img,
        banner_cfg,
        original_width,
        original_height,
        margin,
        shop_name,
    ):
        banner_w = int(banner_cfg["width_pct"] / 100 * original_width)
        banner_h = int(banner_cfg["height_pct"] / 100 * original_height)

        if banner_cfg.get("center"):
            center_x = original_width // 2
            center_y = original_height // 2
            banner_x1 = center_x - banner_w // 2
            banner_y1 = center_y - banner_h // 2
        else:
            banner_x1 = 0
            banner_y1 = 0

        # إضافة margin لموقع البانر
        # banner_x1 += margin
        # banner_y1 += margin
        banner_x2 = banner_x1 + banner_w
        banner_y2 = banner_y1 + banner_h

        draw.rectangle(
            [(banner_x1, banner_y1), (banner_x2, banner_y2)],
            fill=banner_cfg["background_color"],
        )

        logo_cfg = banner_cfg.get("logo", {})
        logo_img = None
        logo_w = logo_h = 0

        logo_img_path = os.path.join(self.assets, logo_cfg.get("image", "logo.png"))
        if os.path.exists(logo_img_path):
            Image.MAX_IMAGE_PIXELS = None
            logo_img = Image.open(logo_img_path).convert("RGBA")
            logo_w = int(logo_cfg.get("width_pct", 10) / 100 * original_width)
            logo_h = int(logo_img.height * logo_w / logo_img.width)
            logo_img = logo_img.resize((logo_w, logo_h), Image.Resampling.LANCZOS)

        text_cfg = banner_cfg.get("shop_name", {})
        font_size = int(text_cfg.get("font_size_pct", 5) / 100 * banner_h)
        font = self._get_font("Bahij_TheSansArabic-Bold.ttf", font_size)
        color = text_cfg.get("color", "#FFFFFF")
        text = self._reshape(shop_name)

        bbox = draw.textbbox((0, 0), text, font=font)
        text_w = bbox[2] - bbox[0]
        text_h = bbox[3] - bbox[1]
        
        banner_type = banner_cfg.get("banner_type", "h").lower()

        # المواضع داخل البانر (بالنسبة إلى banner_x1, banner_y1)
        if banner_type == "h":
            spacing = int(0.04 * banner_w)
            x_logo = banner_x1 + int(logo_cfg.get("x_offset_pct", 2) / 100 * banner_w)
            y_logo = banner_y1 + int(logo_cfg.get("y_offset_pct", 10) / 100 * banner_h)
            x_text = x_logo + logo_w + spacing
            y_text = (y_logo + (logo_h - text_h) // 2) + 5
        elif banner_type == "s7":
            spacing = int(0.03 * banner_w)
            x_logo = banner_x1 + int(logo_cfg.get("x_offset_pct", 2) / 100 * banner_w)
            y_logo = banner_y1 + int(logo_cfg.get("y_offset_pct", 10) / 100 * banner_h)
            x_text = x_logo + logo_w + spacing
            y_text = (y_logo + (logo_h - text_h) // 2) + 5
        elif banner_type == "s8":
            spacing = int(0.02 * banner_w)
            x_logo = banner_x1 + int(logo_cfg.get("x_offset_pct", 2) / 100 * banner_w)
            y_logo = banner_y1 + int(logo_cfg.get("y_offset_pct", 10) / 100 * banner_h)
            x_text = x_logo + logo_w + spacing
            y_text = (y_logo + (logo_h - text_h) // 2) + 5
        else:
            spacing = int(0.08 * banner_h)
            x_logo = banner_x1 + int(logo_cfg.get("x_offset_pct", 2) / 100 * banner_w)
            y_logo = banner_y1 + int(logo_cfg.get("y_offset_pct", 10) / 100 * banner_h)
            x_text = banner_x1 + int(text_cfg.get("x_offset_pct", 50) / 100 * banner_w)
            y_text = y_logo + logo_h + spacing
            x_text = x_logo + (logo_w - text_w) // 2

        elements = []
        if logo_img:
            elements.append((x_logo, y_logo, logo_w, logo_h))
        elements.append((x_text, y_text, text_w, text_h))

        group_x1 = min(e[0] for e in elements)
        group_y1 = min(e[1] for e in elements)
        group_x2 = max(e[0] + e[2] for e in elements)
        group_y2 = max(e[1] + e[3] for e in elements)
        group_w = group_x2 - group_x1
        group_h = group_y2 - group_y1

        max_width = 0.85 * banner_w
        scale_factor = 1.0
        if group_w > max_width:
            scale_factor = max_width / group_w
            new_font_size = int(font_size * scale_factor)
            font = self._get_font("AktivGroteskEx_Md.ttf", new_font_size)
            bbox = draw.textbbox((0, 0), text, font=font)
            text_w = bbox[2] - bbox[0]
            text_h = bbox[3] - bbox[1]
            logo_w = int(logo_w * scale_factor)
            logo_h = int(logo_h * scale_factor)
            if logo_img:
                logo_img = logo_img.resize((logo_w, logo_h), Image.Resampling.LANCZOS)

            if banner_type == "h":
                x_logo = banner_x1 + int(
                    logo_cfg.get("x_offset_pct", 2) / 100 * banner_w
                )
                y_logo = banner_y1 + int(
                    logo_cfg.get("y_offset_pct", 10) / 100 * banner_h
                )
                x_text = x_logo + logo_w + spacing
                y_text = y_logo + (logo_h - text_h) // 2
            else:
                x_logo = banner_x1 + int(
                    logo_cfg.get("x_offset_pct", 2) / 100 * banner_w
                )
                y_logo = banner_y1 + int(
                    logo_cfg.get("y_offset_pct", 10) / 100 * banner_h
                )
                x_text = banner_x1 + int(
                    text_cfg.get("x_offset_pct", 50) / 100 * banner_w
                )
                y_text = y_logo + logo_h + spacing
                x_text = x_logo + (logo_w - text_w) // 2
                
            

            elements = []
            if logo_img:
                elements.append((x_logo, y_logo, logo_w, logo_h))
            elements.append((x_text, y_text, text_w, text_h))
            group_x1 = min(e[0] for e in elements)
            group_y1 = min(e[1] for e in elements)
            group_x2 = max(e[0] + e[2] for e in elements)
            group_y2 = max(e[1] + e[3] for e in elements)
            group_w = group_x2 - group_x1
            group_h = group_y2 - group_y1

        banner_center_x = banner_x1 + banner_w // 2
        banner_center_y = banner_y1 + banner_h // 2

        offset_x = banner_center_x - (group_x1 + group_w // 2)
        offset_y = banner_center_y - (group_y1 + group_h // 2)

        x_logo += offset_x
        y_logo += offset_y
        x_text += offset_x
        y_text += offset_y

        if logo_img:
            img.paste(logo_img, (x_logo, y_logo), logo_img)

        draw.text(
            (x_text, y_text),
            text,
            font=font,
            fill=color,
            anchor="lt",
        )

    def _create_label_image(
        self,
        text: str,
        max_width: int,
        max_height: int,
        font_name: str = "Bahij_TheSansArabic-Bold.ttf",
        padding: int = 12,
    ) -> Tuple[Image.Image, int, int]:
        """
        إنشاء صورة النص مع تصغير حجم الخط تلقائياً.
        بدون تدوير - التدوير يتم بعد ذلك.
        """
        for size in range(200, 9, -5):
            font = self._get_font(font_name, size)
            dummy = Image.new("RGB", (1, 1))
            d = ImageDraw.Draw(dummy)
            bbox = d.textbbox((0, 0), text, font=font)
            text_w = bbox[2] - bbox[0]
            text_h = bbox[3] - bbox[1]

            if text_w <= max_width and text_h <= max_height:
                img_w = text_w + 2 * padding
                img_h = text_h + 2 * padding
                text_img = Image.new("RGBA", (img_w, img_h), (255, 255, 255, 0))
                draw_text = ImageDraw.Draw(text_img)
                draw_text.text(
                    (padding - bbox[0], padding - bbox[1]),
                    text,
                    fill=(255, 0, 0, 255),
                    font=font,
                )
                # print(f"      ✓ Font size {size} fits: {text_w}x{text_h}")
                return text_img, img_w, img_h

        # print(f"      ⚠️ No font fits! Using smallest (10)")
        font = self._get_font(font_name, 10)
        dummy = Image.new("RGB", (1, 1))
        d = ImageDraw.Draw(dummy)
        bbox = d.textbbox((0, 0), text, font=font)
        text_w = bbox[2] - bbox[0]
        text_h = bbox[3] - bbox[1]
        img_w = text_w + 2 * padding
        img_h = text_h + 2 * padding
        text_img = Image.new("RGBA", (img_w, img_h), (255, 255, 255, 0))
        draw_text = ImageDraw.Draw(text_img)
        draw_text.text(
            (padding - bbox[0], padding - bbox[1]),
            text,
            fill=(255, 0, 0, 255),
            font=font,
        )
        return text_img, img_w, img_h

    def _draw_full_background(
        self, img: Image, layer: Dict[str, Any], canvas_width: int, canvas_height: int
    ):
        """رسم صورة الخلفية على كامل القماش (بدون إزاحة margin)"""
        img_path = os.path.join(self.assets, layer.get("image", "missing.png"))
        if not os.path.exists(img_path):
            # إذا لم توجد الصورة، نستخدم لون أبيض كخلفية
            img.paste("#FFFFFF", [0, 0, canvas_width, canvas_height])
            return

        Image.MAX_IMAGE_PIXELS = None
        bg_img = Image.open(img_path).convert("RGB")
        # تغيير حجم الصورة لتملأ القماش بالكامل
        bg_img = bg_img.resize((canvas_width, canvas_height), Image.Resampling.LANCZOS)
        img.paste(bg_img, (0, 0))

    def render(
        self,
        width: int,
        height: int,
        template: str,
        shop_name: Optional[str] = None,
        preview: bool = False,
        safety_margin: int = 0,
        add_label: bool = False,
        label_text: str = "",
        dpi: int = 72,
    ) -> Image.Image:
        tmpl_cfg = self.config.get(template, self.config.get("default", {}))

        # حفظ الأبعاد الأصلية للتصميم (بدون margin)
        original_w = width
        original_h = height

        # الأبعاد النهائية للقماش بعد إضافة margin
        canvas_w = width + 2 * safety_margin
        canvas_h = height + 2 * safety_margin

        # إنشاء الصورة الرئيسية بالحجم الموسع
        img = Image.new("RGB", (canvas_w, canvas_h), "white")
        draw = ImageDraw.Draw(img)

        # قائمة الطبقات من التكوين
        if "images" in tmpl_cfg and isinstance(tmpl_cfg["images"], list):
            all_layers = tmpl_cfg["images"]
        else:
            all_layers = []
            if "mobile_image" in tmpl_cfg:
                all_layers.append(tmpl_cfg["mobile_image"])
            if "logo" in tmpl_cfg:
                all_layers.append(tmpl_cfg["logo"])
            if "icons" in tmpl_cfg:
                all_layers.extend(tmpl_cfg["icons"])

        # فصل طبقة الخلفية (التي تحوي "bg" في اسم الصورة)
        bg_layer = None
        other_layers = []
        for layer in all_layers:
            img_path = layer.get("image", "")
            if "bg" in img_path.lower():  # البحث عن "bg" غير حساس للحالة
                bg_layer = layer
            else:
                other_layers.append(layer)

        # رسم الخلفية (إذا وجدت) على كامل canvas (بدون margin)
        if bg_layer:
            # نستخدم دالة خاصة لرسم الخلفية على canvas كامل
            self._draw_full_background(img, bg_layer, canvas_w, canvas_h)
        else:
            # إذا لم توجد طبقة bg، نرسم الخلفية من القسم القديم (background)
            if "background" in tmpl_cfg:
                self._draw_background(img, tmpl_cfg["background"], canvas_w, canvas_h)

        # رسم باقي الطبقات مع إزاحة margin
        for layer in other_layers:
            self._draw_image_layer(img, layer, canvas_w, canvas_h)

        # رسم البانر إذا وُجد
        if shop_name and "banner" in tmpl_cfg:
            self._draw_banner(
                draw,
                img,
                tmpl_cfg["banner"],
                canvas_w,
                canvas_h,
                safety_margin,
                shop_name,
            )

        # طباعة معلومات debug (مثل السابق)
        def px_to_cm(px, dpi_val):
            return (px / dpi_val) * 2.54

        # print("\n" + "=" * 55)
        # print("📐 RENDERER DEBUG INFO")
        # print(
        #     f"   Original: {original_w}x{original_h} px = {px_to_cm(original_w, dpi):.2f}x{px_to_cm(original_h, dpi):.2f} cm (DPI={dpi})"
        # )
        # if safety_margin > 0:
            # print(f"   Safety margin: {px_to_cm(safety_margin, dpi):.2f} cm each side")
            # print(
            #     f"   Canvas after margin: {canvas_w}x{canvas_h} px = {px_to_cm(canvas_w, dpi):.2f}x{px_to_cm(canvas_h, dpi):.2f} cm"
            # )

        # ========== إضافة الـ Label (بدون تغيير) ==========
        if add_label and label_text:
            label_thickness_px = cm_to_pixels(5, dpi)
            bg_color = "#FFFFFF"
            text_color = "#FF0000"
            max_ratio = 0.92
            padding = max(10, int(dpi * 0.08))

            shaped_text = self._reshape(label_text)
            # print(f"\n🏷️ Label text: '{label_text}'")
            # print(f"   Reshaped: '{shaped_text}'")
            # print(f"   Thickness: {label_thickness_px} px (5 cm)")

            # Portrait: label في الأعلى (شريط أفقي)
            if canvas_h > canvas_w:
                # print("   Orientation: Portrait (label on top)")
                new_w = canvas_w
                new_h = canvas_h + label_thickness_px
                new_img = Image.new("RGB", (new_w, new_h), bg_color)
                new_img.paste(img, (0, label_thickness_px))

                label_width = new_w
                label_height = label_thickness_px
                max_text_width = int(max_ratio * label_width)
                max_text_height = int(max_ratio * label_height)
                # print(f"   Label area: {label_width}x{label_height} px")
                # print(f"   Max text: {max_text_width}x{max_text_height} px")

                text_img, img_w, img_h = self._create_label_image(
                    shaped_text, max_text_width, max_text_height, padding=padding
                )
                x = (label_width - img_w) // 2
                y = (label_height - img_h) // 2
                # print(f"   Text image: {img_w}x{img_h} px, paste at ({x}, {y})")
                new_img.paste(text_img, (x, y), text_img)
                img = new_img

            # Landscape: label على اليسار (شريط عمودي)
            else:
                # print("   Orientation: Landscape (label on left)")
                new_w = canvas_w + label_thickness_px
                new_h = canvas_h
                new_img = Image.new("RGB", (new_w, new_h), bg_color)
                new_img.paste(img, (label_thickness_px, 0))

                label_width = label_thickness_px
                label_height = new_h
                max_original_width = int(max_ratio * label_height)
                max_original_height = int(max_ratio * label_width)
                # print(f"   Label area: {label_width}x{label_height} px")
                # print(
                #     f"   Max original text: {max_original_width}x{max_original_height} px"
                # )

                text_img, img_w, img_h = self._create_label_image(
                    shaped_text,
                    max_original_width,
                    max_original_height,
                    padding=padding,
                )
                rotated = text_img.rotate(-90, expand=True)
                # print(f"   Original text image: {img_w}x{img_h} px")
                # print(f"   Rotated text image: {rotated.width}x{rotated.height} px")
                x = (label_width - rotated.width) // 2
                y = (label_height - rotated.height) // 2
                # print(f"   Paste at ({x}, {y})")
                new_img.paste(rotated, (x, y), rotated)
                img = new_img

            # print("=" * 55 + "\n")

        # تصغير الصورة للمعاينة إذا لزم الأمر
        if preview:
            img.thumbnail(
                (1080, int(1080 * img.height / img.width)), Image.Resampling.LANCZOS
            )

        return img
