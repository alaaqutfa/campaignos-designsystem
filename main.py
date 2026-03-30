import argparse
import logging
import os
import sys
from typing import List, Dict, Any
from modules.excel_reader import read_excel
from modules.config_loader import load_sizes, load_layout
from modules.template_matcher import (
    match_template,
    extract_templates_from_config,
    get_tolerance,
)
from modules.renderer import DesignRenderer
from modules.utils import cm_to_pixels, ensure_dir, setup_logging, sanitize_filename


def parse_arguments():
    """Parse command line arguments."""
    parser = argparse.ArgumentParser(
        description="Generate advertising designs from an Excel file.",
        epilog="Example: python main.py --input Cities/Anbar.xlsx --output ./out --config config_layout.yaml --sizes config_sizes.yaml",
    )
    parser.add_argument(
        "--input",
        "-i",
        required=True,
        help="Path to input Excel file (e.g., Cities/Anbar.xlsx)",
    )
    parser.add_argument(
        "--sheet",
        "-s",
        default="Sheet1",
        help="Sheet name inside Excel file (default: Sheet1)",
    )
    parser.add_argument(
        "--output",
        "-o",
        default="./out",
        help="Output directory for generated images/PDF (default: ./out)",
    )
    parser.add_argument(
        "--config",
        "-c",
        default="config_layout.yaml",
        help="Layout configuration file (YAML) (default: config_layout.yaml)",
    )
    parser.add_argument(
        "--sizes",
        default="config_sizes.yaml",
        help="Template sizes configuration file (YAML) (default: config_sizes.yaml)",
    )
    parser.add_argument(
        "--dpi",
        type=int,
        default=72,
        help="Output DPI for generated images (default: 72)",
    )
    parser.add_argument(
        "--preview", action="store_true", help="Generate low-resolution preview images"
    )
    parser.add_argument(
        "--generate-sample",
        type=int,
        metavar="N",
        help="Generate only first N rows (for testing)",
    )
    parser.add_argument(
        "--generate-test-sample",
        action="store_true",
        help="Generate only versions WITHOUT shop name",
    )
    parser.add_argument(
        "--with-shop",
        action="store_true",
        help="Generate only versions WITH shop name",
    )
    parser.add_argument(
        "--without-shop",
        action="store_true",
        help="Generate only versions WITHOUT shop name",
    )
    parser.add_argument(
        "--format",
        choices=["png", "pdf", "jpg", "tiff"],
        default="png",
        help="Output format (png, pdf, jpg or tiff, default: png)",
    )
    parser.add_argument("--debug", action="store_true", help="Enable debug logging")
    parser.add_argument("--log-file", help="Optional log file path")

    args = parser.parse_args()

    if args.with_shop and args.without_shop:
        parser.error("Cannot use --with-shop and --without-shop together.")

    return args


def main():
    """Main execution function."""
    args = parse_arguments()

    # Setup logging
    setup_logging(debug=args.debug, log_file=args.log_file)
    logger = logging.getLogger(__name__)

    try:
        # 1. Load configurations
        logger.info("Loading configuration files...")
        sizes_config = load_sizes(args.sizes)
        layout_config = load_layout("config_layout.yaml")
        shop_layout_config = load_layout("config_layout_shop.yaml")
        logger.info("   - Configuration files loaded successfully.")

        # 2. Read Excel data
        logger.info(f"Reading Excel file: {args.input}")
        rows = read_excel(args.input, sheet_name=args.sheet)
        logger.info(f"  - Loaded {len(rows)} rows from Excel.")

        # 3. Sample mode
        if args.generate_sample:
            rows = rows[: args.generate_sample]
            logger.info(f"  - Using first {args.generate_sample} rows for testing.")

        # 4. Initialize renderer
        renderer = DesignRenderer(layout_config, assets_path="./assets")
        shop_renderer = DesignRenderer(shop_layout_config, assets_path="./assets")

        # 5. Ensure output directory
        ensure_dir(args.output)

        # 6. Extract templates
        templates = extract_templates_from_config(sizes_config)
        tolerance = get_tolerance(sizes_config)
        
        fmt = args.format
        if fmt.lower() in ("jpg", "jpeg"):
            fmt = "JPEG"
        elif fmt.lower() == "tiff":
            fmt = "TIFF"
        elif fmt.lower() == "png":
            fmt = "PNG"
        elif fmt.lower() == "pdf":
            fmt = "PDF"

        # 7. Process rows
        total_designs = 0
        if args.generate_test_sample:
            for template in templates:
                base_filename = template['name']
                shop_name = "Shop Name"
                template_name = template['name']
                width_px = int(template['width'])
                height_px = int(template['height'])
                ratio = template['ratio']
                def save_design(img, suffix):
                    nonlocal total_designs
                    filename = f"{base_filename}{suffix}.{args.format}"
                    filepath = os.path.join(args.output, filename)
                    if args.format == "png":
                        img.save(filepath)
                    else:
                        img.save(filepath, fmt, resolution=args.dpi)
                    logger.info(f"Saved: {filepath}")
                    total_designs += 1
                img = renderer.render(
                        width=width_px,
                        height=height_px,
                        template=template_name,
                        shop_name=None,
                        preview=args.preview,
                    )
                save_design(img, " - Ar - 02")
                img = shop_renderer.render(
                            width=width_px,
                            height=height_px,
                            template=template_name,
                            shop_name=str(shop_name).strip(),
                            preview=args.preview,
                        )
                save_design(img, " - Ar - 01")
        else:
            for idx, row in enumerate(rows, 1):
                logger.info(
                    f"Processing row {idx}: {row.get('print_file_name', 'unknown')}"
                )

                try:
                    width_cm = float(row["width"])
                    height_cm = float(row["height"])
                    shop_name = row.get("shop_name", "")
                    base_filename = sanitize_filename(
                        row.get("print_file_name", f"design_{idx}")
                    )
                except (ValueError, KeyError) as e:
                    logger.error(f"Error in row {idx}: {e}. Skipping.")
                    continue

                if height_cm == 0:
                    logger.error(f"Height is zero in row {idx}. Skipping.")
                    continue

                ratio = width_cm / height_cm
                template_name = match_template(ratio, templates, tolerance)

                width_px = cm_to_pixels(width_cm, args.dpi)
                height_px = cm_to_pixels(height_cm, args.dpi)

                generate_with_shop = False
                generate_without_shop = False

                if args.with_shop:
                    generate_with_shop = True
                elif args.without_shop:
                    generate_without_shop = True
                else:
                    generate_with_shop = True
                    generate_without_shop = True

                def save_design(img, suffix):
                    nonlocal total_designs
                    filename = f"{base_filename}{suffix}.{args.format}"
                    filepath = os.path.join(args.output, filename)
                    
                    if args.format == "png":
                        img.save(filepath)
                    else:
                        img.save(filepath, fmt, resolution=args.dpi)
                    logger.debug(f"Saved: {filepath}")
                    total_designs += 1

                
                # Without shop name
                if generate_without_shop:
                    logger.debug("Generating version WITHOUT shop name...")
                    img = renderer.render(
                        width=width_px,
                        height=height_px,
                        template=template_name,
                        shop_name=None,
                        preview=args.preview,
                    )
                    save_design(img, " - Ar - 02")

                # With shop name
                if generate_with_shop:
                    if shop_name and str(shop_name).strip():
                        logger.debug("Generating version WITH shop name...")
                        img = shop_renderer.render(
                            width=width_px,
                            height=height_px,
                            template=template_name,
                            shop_name=str(shop_name).strip(),
                            preview=args.preview,
                        )
                        save_design(img, " - Ar - 01")
                    else:
                        logger.warning(f"Row {idx} has no shop name for WITHSHOP version.")

            logger.info(f"Done. Generated {total_designs} designs in: {args.output}")

    except Exception as e:
        logger.exception(f"Unexpected error occurred: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
