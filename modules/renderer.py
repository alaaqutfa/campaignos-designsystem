import os
from PIL import Image, ImageDraw, ImageFont
import arabic_reshaper
from bidi.algorithm import get_display
import re
from typing import Dict, Any, List, Tuple, Optional


class DesignRenderer:
    def __init__(self, layout_config: Dict[str, Any], assets_path: str = "./assets"):
        """
        Args:
            layout_config: قاموس التكوين (محمل من config_layout.yaml)
            assets_path: المسار إلى مجلد الأصول (صور، خطوط، أيقونات)
        """
        self.config = layout_config
        self.assets = assets_path
        self.fonts_cache = {}  # تخزين مؤقت للخطوط لتجنب إعادة التحميل

    def _reshape(self, text: str) -> str:
        """
        تشكيل النص العربي وعكس اتجاهه للعرض الصحيح في Pillow.

        Args:
            text: النص الأصلي (قد يحتوي على عربية وإنجليزية)

        Returns:
            النص المعاد تشكيله وجاهز للعرض
        """
        reshaped = arabic_reshaper.reshape(text)
        return get_display(reshaped)

    def _get_font(self, font_name: str, size: int) -> ImageFont.FreeTypeFont:
        """
        تحميل خط من مجلد assets/fonts مع التخزين المؤقت.

        Args:
            font_name: اسم ملف الخط (مثل 'NotoNaskhArabic-Regular.ttf')
            size: حجم الخط بالبكسل

        Returns:
            كائن الخط من Pillow
        """
        cache_key = f"{font_name}_{size}"
        if cache_key in self.fonts_cache:
            return self.fonts_cache[cache_key]

        font_path = os.path.join(self.assets, "fonts", font_name)
        if not os.path.exists(font_path):
            # محاولة استخدام خط النظام الافتراضي كبديل
            try:
                font = ImageFont.truetype("arial.ttf", size)
            except:
                font = ImageFont.load_default()
        else:
            font = ImageFont.truetype(font_path, size)
        self.fonts_cache[cache_key] = font
        return font

    def _parse_bold_segments(self, text: str) -> List[Tuple[str, bool]]:
        """
        يقسم النص إلى أجزاء عادية وأجزاء عريضة (الأرقام والوحدات).
        الأرقام وما يتبعها من وحدات (مثل Hz, K, mAh, سنوات) تعتبر عريضة.

        Args:
            text: النص الكامل لعبارة واحدة

        Returns:
            قائمة من الأزواج (النص, هل هو عريض)
        """
        # وحدات محتملة بالعربية والإنجليزية
        units = r"(?:[a-zA-Z]+|سنة|سنوات|Hz|K|mAh|G|مم|سم|بكسل|ميجابكسل|جيجا|تيترا|سنة|شهور|أيام)"
        # النمط: رقم (قد يكون عشري) متبوعًا اختياريًا بمسافة ووحدة
        pattern = r"(\d+(?:\.\d+)?\s*(?:" + units + ")?)"

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

    def _measure_text(self, draw: ImageDraw, text: str, font: ImageFont) -> int:
        """
        قياس عرض نص معين بعد تشكيله.

        Args:
            draw: كائن الرسم (لقياس النص)
            text: النص الأصلي
            font: الخط المستخدم

        Returns:
            العرض بالبكسل
        """
        reshaped = self._reshape(text)
        bbox = draw.textbbox((0, 0), reshaped, font=font)
        return bbox[2] - bbox[0]

    def _render_phrase(
        self,
        draw: ImageDraw,
        x: int,
        y: int,
        phrase: str,
        regular_font: ImageFont,
        bold_font: ImageFont,
        color: str,
        anchor: str = "rt",
    ) -> int:
        """
        رسم عبارة واحدة (قد تحتوي على أجزاء عريضة) وإرجاع عرضها الإجمالي.

        تعمل هذه الدالة عن طريق:
        1. تحليل العبارة إلى أجزاء.
        2. قياس عرض كل جزء.
        3. رسم الأجزاء من اليمين إلى اليسار (باتجاه RTL) لأن anchor = 'rt'.

        Args:
            draw: كائن الرسم
            x: الإحداثي الأفقي لنقطة الارتكاز (اليمين)
            y: الإحداثي الرأسي
            phrase: العبارة المراد رسمها
            regular_font: الخط العادي
            bold_font: الخط العريض
            color: لون النص (مثل '#000000')
            anchor: نقطة الارتكاز في النص (rt = right-top)

        Returns:
            العرض الإجمالي للعبارة بالبكسل
        """
        segments = self._parse_bold_segments(phrase)

        # قياس العرض الإجمالي
        total_width = 0
        for seg_text, is_bold in segments:
            font = bold_font if is_bold else regular_font
            total_width += self._measure_text(draw, seg_text, font)

        # رسم الأجزاء من اليمين إلى اليسار
        current_x = x
        for seg_text, is_bold in reversed(segments):  # نعكس لأننا نرسم من اليمين
            font = bold_font if is_bold else regular_font
            reshaped = self._reshape(seg_text)
            # قياس هذا الجزء
            seg_width = self._measure_text(draw, seg_text, font)
            # رسم الجزء
            draw.text((current_x, y), reshaped, font=font, fill=color, anchor="rt")
            # تحريك المؤشر لليسار بمقدار عرض الجزء
            current_x -= seg_width

        return total_width

    def _draw_background(
        self, img: Image, bg_cfg: Dict[str, Any], width: int, height: int
    ):
        """
        رسم الخلفية حسب التكوين (لون، تدرج، أو صورة).

        Args:
            img: الصورة الرئيسية
            bg_cfg: إعدادات الخلفية
            width, height: أبعاد الصورة
        """
        bg_type = bg_cfg.get("type", "color")
        if bg_type == "color":
            # خلفية بلون واحد
            color = bg_cfg.get("color", "#FFFFFF")
            # إذا كان اللون عبارة عن سلسلة تبدأ بـ #، نفترض أنها hex
            if isinstance(color, str) and color.startswith("#"):
                img.paste(color, [0, 0, width, height])
            else:
                # قد يكون RGB tuple
                img.paste(color, [0, 0, width, height])
        elif bg_type == "gradient":
            # يمكن تطويرها لاحقاً
            color1 = bg_cfg.get("color1", "#FFFFFF")
            color2 = bg_cfg.get("color2", "#CCCCCC")
            # رسم تدرج خطي عمودي
            draw = ImageDraw.Draw(img)
            for y in range(height):
                r = int(y / height * 255)
                # مزج بسيط
                draw.line([(0, y), (width, y)], fill=(r, r, r))
        elif bg_type == "image":
            # صورة خلفية
            img_path = os.path.join(self.assets, bg_cfg.get("image", "background.png"))
            if os.path.exists(img_path):
                bg_img = Image.open(img_path).convert("RGB")
                bg_img = bg_img.resize((width, height), Image.Resampling.LANCZOS)
                img.paste(bg_img, (0, 0))
            else:
                # إذا لم توجد الصورة، نستخدم لون افتراضي
                img.paste("#FFFFFF", [0, 0, width, height])

    def _draw_mobile_image(
        self,
        draw: ImageDraw,
        img: Image,
        mobile_cfg: Dict[str, Any],
        width: int,
        height: int,
    ):
        """
        وضع صورة الموبايل في التصميم.

        Args:
            draw: كائن الرسم (غير مستخدم هنا ولكن للاتساق)
            img: الصورة الرئيسية (سنقوم باللصق عليها)
            mobile_cfg: إعدادات صورة الموبايل
            width, height: أبعاد التصميم
        """
        img_path = os.path.join(self.assets, mobile_cfg.get("image", "mobile.png"))
        if not os.path.exists(img_path):
            return  # لا توجد صورة

        mobile_img = Image.open(img_path).convert("RGBA")

        # حساب عرض الصورة كنسبة من العرض الكلي
        mobile_w = int(mobile_cfg["width_pct"] / 100 * width)
        mobile_img = mobile_img.resize(
            (mobile_w, int(mobile_img.height * mobile_w / mobile_img.width)),
            Image.Resampling.LANCZOS,
        )

        # تحديد نقطة الارتكاز (anchor) والإزاحة
        anchor = mobile_cfg.get("anchor", "center-center")
        x_offset = int(mobile_cfg.get("x_offset_pct", 0) / 100 * width)
        y_offset = int(mobile_cfg.get("y_offset_pct", 0) / 100 * height)

        if anchor == "right-center":
            x = width - mobile_img.width + x_offset
            y = (height - mobile_img.height) // 2 + y_offset
        elif anchor == "left-center":
            x = x_offset
            y = (height - mobile_img.height) // 2 + y_offset
        elif anchor == "center":
            x = (width - mobile_img.width) // 2 + x_offset
            y = (height - mobile_img.height) // 2 + y_offset
        elif anchor == "top-right":
            x = width - mobile_img.width + x_offset
            y = y_offset
        elif anchor == "top-left":
            x = x_offset
            y = y_offset
        elif anchor == "bottom-right":
            x = width - mobile_img.width + x_offset
            y = height - mobile_img.height + y_offset
        elif anchor == "bottom-left":
            x = x_offset
            y = height - mobile_img.height + y_offset
        else:
            # افتراضي: right-center
            x = width - mobile_img.width + x_offset
            y = (height - mobile_img.height) // 2 + y_offset

        img.paste(mobile_img, (x, y), mobile_img)

    def _draw_logo(
        self,
        draw: ImageDraw,
        img: Image,
        logo_cfg: Dict[str, Any],
        width: int,
        height: int,
    ):
        """
        وضع الشعار الرئيسي.

        Args:
            draw: كائن الرسم
            img: الصورة الرئيسية
            logo_cfg: إعدادات الشعار
            width, height: أبعاد التصميم
        """
        img_path = os.path.join(self.assets, logo_cfg.get("image", "logo.png"))
        if not os.path.exists(img_path):
            return

        logo_img = Image.open(img_path).convert("RGBA")
        logo_w = int(logo_cfg["width_pct"] / 100 * width)
        logo_img = logo_img.resize(
            (logo_w, int(logo_img.height * logo_w / logo_img.width)),
            Image.Resampling.LANCZOS,
        )

        anchor = logo_cfg.get("anchor", "top-left")
        x_offset = int(logo_cfg.get("x_offset_pct", 0) / 100 * width)
        y_offset = int(logo_cfg.get("y_offset_pct", 0) / 100 * height)

        # تحديد الإحداثيات بناءً على نقطة الارتكاز
        if anchor == "top-left":
            x = x_offset
            y = y_offset
        elif anchor == "top-right":
            x = width - logo_img.width + x_offset
            y = y_offset
        elif anchor == "bottom-left":
            x = x_offset
            y = height - logo_img.height + y_offset
        elif anchor == "bottom-right":
            x = width - logo_img.width + x_offset
            y = height - logo_img.height + y_offset
        elif anchor == "center":
            x = (width - logo_img.width) // 2 + x_offset
            y = (height - logo_img.height) // 2 + y_offset
        else:
            x = x_offset
            y = y_offset

        img.paste(logo_img, (x, y), logo_img)

    def _draw_icons(
        self,
        draw: ImageDraw,
        img: Image,
        icons_cfg: List[Dict[str, Any]],
        width: int,
        height: int,
    ):
        """
        Args:
            draw: كائن الرسم
            img: الصورة الرئيسية
            icons_cfg: قائمة إعدادات الأيقونات
            width, height: أبعاد التصميم
        """
        for icon_cfg in icons_cfg:
            img_name = icon_cfg.get("image", "icon.png")
            img_path = os.path.join(self.assets, "icons", img_name)
            if not os.path.exists(img_path):
                continue

            icon_img = Image.open(img_path).convert("RGBA")
            icon_w = int(icon_cfg["width_pct"] / 100 * width)
            icon_h = int(icon_cfg["height_pct"] / 100 * height)
            icon_img = icon_img.resize((icon_w, icon_h), Image.Resampling.LANCZOS)

            anchor = icon_cfg.get("anchor", "top-left")
            x_offset = int(icon_cfg.get("x_offset_pct", 0) / 100 * width)
            y_offset = int(icon_cfg.get("y_offset_pct", 0) / 100 * height)

            if anchor == "bottom-left":
                x = x_offset
                y = height - icon_img.height + y_offset
            elif anchor == "bottom-right":
                x = width - icon_img.width + x_offset
                y = height - icon_img.height + y_offset
            elif anchor == "top-left":
                x = x_offset
                y = y_offset
            else:
                x = x_offset
                y = height - icon_img.height + y_offset

            img.paste(icon_img, (x, y), icon_img)

    def _draw_banner(
        self,
        draw,
        img,
        banner_cfg,
        width,
        height,
        shop_name,
    ):
        # ======================
        # حساب أبعاد البانر
        # ======================
        banner_w = int(banner_cfg["width_pct"] / 100 * width)
        banner_h = int(banner_cfg["height_pct"] / 100 * height)

        # ======================
        # تحديد مكان البانر
        # ======================
        if banner_cfg.get("center"):
            center_x = width // 2
            center_y = height // 2

            banner_x1 = center_x - banner_w // 2
            banner_y1 = center_y - banner_h // 2
        else:
            banner_x1 = 0
            banner_y1 = 0

        banner_x2 = banner_x1 + banner_w
        banner_y2 = banner_y1 + banner_h

        # رسم البانر
        draw.rectangle(
            [(banner_x1, banner_y1), (banner_x2, banner_y2)],
            fill=banner_cfg["background_color"],
        )

        # ======================
        # تحميل الشعار
        # ======================
        logo_cfg = banner_cfg.get("logo", {})
        logo_img = None
        logo_w = logo_h = 0

        logo_img_path = os.path.join(self.assets, logo_cfg.get("image", "logo.png"))
        if os.path.exists(logo_img_path):
            logo_img = Image.open(logo_img_path).convert("RGBA")

            logo_w = int(logo_cfg.get("width_pct", 10) / 100 * width)
            logo_h = int(logo_img.height * logo_w / logo_img.width)

            logo_img = logo_img.resize((logo_w, logo_h), Image.Resampling.LANCZOS)

        # ======================
        # إعداد النص
        # ======================
        text_cfg = banner_cfg.get("shop_name", {})
        font_size = int(text_cfg.get("font_size_pct", 5) / 100 * banner_h)
        font = self._get_font("NotoNaskhArabic-Bold.ttf", font_size)
        color = text_cfg.get("color", "#FFFFFF")

        text = self._reshape(shop_name)

        bbox = draw.textbbox((0, 0), text, font=font)
        text_w = bbox[2] - bbox[0]
        text_h = bbox[3] - bbox[1]

        # ======================
        # مواقع أولية داخل البانر
        # ======================
        x_logo = banner_x1 + int(logo_cfg.get("x_offset_pct", 2) / 100 * banner_w)
        y_logo = banner_y1 + int(logo_cfg.get("y_offset_pct", 10) / 100 * banner_h)

        x_text = banner_x1 + int(text_cfg.get("x_offset_pct", 50) / 100 * banner_w)
        y_text = banner_y1 + int(text_cfg.get("y_offset_pct", 50) / 100 * banner_h)

        # ======================
        # حساب bounding box للمجموعة
        # ======================
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

        # ======================
        # توسيط المجموعة داخل البانر
        # ======================
        banner_center_x = banner_x1 + banner_w // 2
        banner_center_y = banner_y1 + banner_h // 2

        offset_x = banner_center_x - (group_x1 + group_w // 2)
        offset_y = banner_center_y - (group_y1 + group_h // 2)

        # تطبيق الإزاحة
        x_logo += offset_x
        y_logo += offset_y
        x_text += offset_x
        y_text += offset_y

        # ======================
        # الرسم النهائي
        # ======================
        if logo_img:
            img.paste(logo_img, (x_logo, y_logo), logo_img)

        draw.text(
            (x_text, y_text),
            text,
            font=font,
            fill=color,
            anchor="lt",
        )

    def render(
        self,
        width: int,
        height: int,
        template: str,
        shop_name: Optional[str] = None,
        preview: bool = False,
    ) -> Image.Image:
        """
        Args:
            width: عرض التصميم بالبكسل
            height: ارتفاع التصميم بالبكسل
            template: اسم القالب المطابق (مثل S2, default)
            shop_name: اسم المحل (إذا كان None، لا يضاف شريط)
            preview: إذا كان True، يتم تصغير الصورة للعرض السريع
        """
        # تحميل إعدادات القالب (إذا لم يوجد، نستخدم default)
        tmpl_cfg = self.config.get(template, self.config.get("default", {}))

        # إنشاء صورة فارغة
        img = Image.new("RGB", (width, height), "white")
        draw = ImageDraw.Draw(img)

        # رسم الخلفية
        if "background" in tmpl_cfg:
            self._draw_background(img, tmpl_cfg["background"], width, height)

        # رسم الخلفية
        if "background2" in tmpl_cfg:
            self._draw_background(img, tmpl_cfg["background2"], width, height)

        # رسم الأيقونات
        if "icons" in tmpl_cfg and isinstance(tmpl_cfg["icons"], list):
            self._draw_icons(draw, img, tmpl_cfg["icons"], width, height)

        # إضافة شريط اسم المحل إذا كان مطلوباً
        if shop_name and "banner" in tmpl_cfg:
            self._draw_banner(draw, img, tmpl_cfg["banner"], width, height, shop_name)

        # رسم صورة الموبايل
        if "mobile_image" in tmpl_cfg:
            self._draw_mobile_image(draw, img, tmpl_cfg["mobile_image"], width, height)

        # رسم صورة الموبايل
        if "mobile_image2" in tmpl_cfg:
            self._draw_mobile_image(draw, img, tmpl_cfg["mobile_image2"], width, height)

        # رسم الشعار الرئيسي
        if "logo" in tmpl_cfg:
            self._draw_logo(draw, img, tmpl_cfg["logo"], width, height)

        # رسم الشعار الرئيسي 2
        if "logo2" in tmpl_cfg:
            self._draw_logo(draw, img, tmpl_cfg["logo2"], width, height)

        # إذا كان preview، نصغر الصورة
        if preview:
            img.thumbnail((1080, int(1080 * height / width)), Image.Resampling.LANCZOS)

        return img
