import argparse
import logging
import os
import sys
from modules.excel_reader import read_excel
from modules.config_loader import load_sizes, load_layout
from modules.template_matcher import (
    match_template,
    extract_templates_from_config,
    get_tolerance,
)
from modules.last_renderer import DesignRenderer
from modules.utils import cm_to_pixels, ensure_dir, setup_logging, sanitize_filename


def parse_arguments():
    parser = argparse.ArgumentParser(
        description="Generate advertising designs from an Excel file.",
        epilog="Example: python main.py --input Cities/Anbar.xlsx --output ./out --config config_layout.yaml --sizes config_sizes.yaml",
    )
    parser.add_argument(
        "--input",
        "-i",
        help="Path to input Excel file (required unless --generate-test-sample is used)",
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
        help="Generate test designs for all templates (no Excel required)",
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

    parser.add_argument(
        "--margin",
        type=float,
        default=0.0,
        help="Safety margin in centimeters (default: 0)",
    )
    parser.add_argument(
        "--material",
        default="Mesh",
        help="Material type for margin (default: Mesh)",
    )
    parser.add_argument(
        "--add-label",
        action="store_true",
        help="Add a label with sequential number and print file name",
    )
    parser.add_argument(
        "--batch",
        action="store_true",
        help="Process all Excel files in the 'Cities' folder",
    )

    args = parser.parse_args()

    if args.with_shop and args.without_shop:
        parser.error("Cannot use --with-shop and --without-shop together.")

    # Input is required unless test sample
    if not args.input and not args.generate_test_sample:
        parser.error("--input is required unless --generate-test-sample is used")

    return args


def setup_default_logging(debug=False, log_file=None):
    """Set up logging with a default file handler if none provided."""
    if log_file:
        setup_logging(debug=debug, log_file=log_file)
    else:
        default_log = "error.log"
        setup_logging(debug=debug, log_file=default_log)
        # Ensure console output as well
        root_logger = logging.getLogger()
        console = logging.StreamHandler()
        console.setLevel(logging.DEBUG if debug else logging.INFO)
        console.setFormatter(
            logging.Formatter("%(asctime)s - %(name)s - %(levelname)s - %(message)s")
        )
        root_logger.addHandler(console)


def main():
    args = parse_arguments()
    setup_default_logging(debug=args.debug, log_file=args.log_file)
    logger = logging.getLogger(__name__)

    # --- Test sample mode (no Excel) ---
    if args.generate_test_sample:
        logger.info(
            "Test sample mode: generating designs for all templates (no Excel required)."
        )
        try:
            sizes_config = load_sizes(args.sizes)
            layout_config = load_layout("config_layout.yaml")
            shop_layout_config = load_layout("config_layout_shop.yaml")
            templates = extract_templates_from_config(sizes_config)
            renderer = DesignRenderer(layout_config, assets_path="./assets")
            shop_renderer = DesignRenderer(shop_layout_config, assets_path="./assets")
            ensure_dir(args.output)

            fmt = args.format
            if fmt.lower() in ("jpg", "jpeg"):
                fmt = "JPEG"
            elif fmt.lower() == "tiff":
                fmt = "TIFF"
            elif fmt.lower() == "png":
                fmt = "PNG"
            elif fmt.lower() == "pdf":
                fmt = "PDF"

            total = 0
            for template in templates:
                base_filename = template["name"]
                shop_name = "Shop Name"
                template_name = template["name"]
                width_px = int(template["width"])
                height_px = int(template["height"])

                def save(img, suffix):
                    nonlocal total
                    filename = f"{base_filename}{suffix}.{args.format}"
                    filepath = os.path.join(args.output, filename)
                    if args.format == "png":
                        img.save(filepath)
                    else:
                        img.save(filepath, fmt, resolution=args.dpi)
                    logger.info(f"Saved: {filepath}")
                    total += 1

                # Without shop name
                img = renderer.render(
                    width=width_px,
                    height=height_px,
                    template=template_name,
                    shop_name=None,
                    preview=args.preview,
                    safety_margin=cm_to_pixels(args.margin, args.dpi),
                    add_label=args.add_label,
                    label_text=None,
                )
                save(img, " - 01")

                # With shop name
                img = shop_renderer.render(
                    width=width_px,
                    height=height_px,
                    template=template_name,
                    shop_name=shop_name,
                    preview=args.preview,
                    safety_margin=cm_to_pixels(args.margin, args.dpi),
                    add_label=args.add_label,
                    label_text=None,
                )
                save(img, "")

            logger.info(
                f"Test sample generation completed. Generated {total} designs in {args.output}"
            )
            return
        except Exception as e:
            logger.exception(f"Test sample generation failed: {e}")
            sys.exit(1)

    # --- Batch mode ---
    if args.batch:
        cities_dir = "Cities"
        if not os.path.isdir(cities_dir):
            logger.error(f"Folder '{cities_dir}' not found.")
            sys.exit(1)
        excel_files = [f for f in os.listdir(cities_dir) if f.endswith(".xlsx")]
        if not excel_files:
            logger.error(f"No Excel files found in '{cities_dir}'.")
            sys.exit(1)
        for file in excel_files:
            province = os.path.splitext(file)[0]
            input_file = os.path.join(cities_dir, file)
            output_dir = os.path.join(args.output, province)
            logger.info(f"Processing province: {province}")
            try:
                generate_from_file(input_file, output_dir, args, logger)
            except Exception as e:
                logger.exception(f"Error processing {province}: {e}")
        logger.info("Batch processing completed.")
        return

    # --- Single file mode ---
    generate_from_file(args.input, args.output, args, logger)


def generate_from_file(input_file, output_dir, args, logger):
    """Generate designs for a single Excel file."""
    # Load configurations
    logger.info("Loading configuration files...")
    sizes_config = load_sizes(args.sizes)
    layout_config = load_layout("config_layout.yaml")
    shop_layout_config = load_layout("config_layout_shop.yaml")
    logger.info("   - Configuration files loaded successfully.")

    # Read Excel
    logger.info(f"Reading Excel file: {input_file}")
    try:
        rows = read_excel(input_file, sheet_name=args.sheet)
    except Exception as e:
        logger.error(f"Failed to read Excel file: {e}")
        sys.exit(1)

    logger.info(f"  - Loaded {len(rows)} rows from Excel.")

    # Validate rows
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
        sys.exit(1)

    # Sample mode
    if args.generate_sample:
        valid_rows = valid_rows[: args.generate_sample]
        logger.info(f"  - Using first {args.generate_sample} rows for testing.")

    # Initialize renderer
    renderer = DesignRenderer(layout_config, assets_path="./assets")
    shop_renderer = DesignRenderer(shop_layout_config, assets_path="./assets")

    ensure_dir(output_dir)

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

    total_designs = 0
    shop_counter = {}

    for idx, row in valid_rows:
        logger.info(f"Processing row {idx}: {row.get('print_file_name', 'unknown')}")

        width_cm = float(row["width"])
        height_cm = float(row["height"])
        shop_name = row.get("text", "")
        base_filename = sanitize_filename(row.get("print_file_name", f"design_{idx}"))

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

        margin_px = cm_to_pixels(args.margin, args.dpi)

        label_text = None
        if args.add_label:
            if shop_name not in shop_counter:
                shop_counter[shop_name] = 0
            shop_counter[shop_name] += 1
            label_text = f"[{shop_counter[shop_name]}] {base_filename}"

        def save_design(img, suffix):
            nonlocal total_designs
            filename = f"{base_filename}{suffix}.{args.format}"
            filepath = os.path.join(output_dir, filename)
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
                safety_margin=margin_px,
                add_label=args.add_label,
                label_text=label_text,
            )
            save_design(img, " - 01")

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
                    safety_margin=margin_px,
                    add_label=args.add_label,
                    label_text=label_text,
                )
                save_design(img, "")
            else:
                logger.warning(f"Row {idx} has no shop name for WITHSHOP version.")

    logger.info(f"Done. Generated {total_designs} designs in: {output_dir}")


if __name__ == "__main__":
    main()
