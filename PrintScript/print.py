#!/usr/bin/python3
import os               # Filesystem operations
import subprocess       # Run system commands
import time             # Sleep between cycles
import json             # Handle JSON settings
import sys              # For custom stdout sink
from loguru import logger  # Advanced logging

# Remove default logger configuration
logger.remove()

# Add a colorful console sink with bold formatting and emojis.
logger.add(
    sys.stdout,
    colorize=True,
    format=(
        "<bold><cyan>🚀 {time:YYYY-MM-DD HH:mm:ss.SSS}</cyan></bold> | "
        "<level>{level: <8}</level> | "
        "<bold><yellow>{name}</yellow></bold>:"
        "<bold><magenta>{function}</magenta></bold>:"
        "<bold><blue>{line}</blue></bold> - "
        "<level>{message}</level>"
    ),
    level="DEBUG"
)

# Add a file sink that clears the log file on restart.
logger.add(
    "printer_monitor.log",
    mode="w",                # Clear log file on restart
    rotation="5 MB",         # Rotate after 5MB
    retention="10 days",     # Keep logs for 10 days
    compression="zip",       # Compress rotated logs
    level="DEBUG",
    format="{time:YYYY-MM-DD HH:mm:ss.SSS} | {level: <8} | {name}:{function}:{line} - {message}"
)


def get_default_printer():
    """
    Retrieves the system's default printer using 'lpstat -d'.
    Returns the printer name as a string or None if not found.
    """
    try:
        result = subprocess.run(['lpstat', '-d'], capture_output=True, text=True, check=True)
        printer_name = result.stdout.strip().split(': ')[1]
        return printer_name
    except (subprocess.CalledProcessError, IndexError):
        logger.error("❌ Failed to retrieve the system default printer.")
        return None


def load_settings(settings_file="print_settings.json"):
    """
    Loads settings from the JSON file.
    """
    try:
        with open(settings_file, 'r') as f:
            settings = json.load(f)
        return settings
    except FileNotFoundError:
        logger.error(f"❌ Settings file '{settings_file}' not found.")
        return None
    except json.JSONDecodeError:
        logger.error(f"❌ Error parsing '{settings_file}'. Check JSON syntax.")
        return None


def check_printer_availability(printer_name):
    """
    Checks if the specified printer is available using 'lpstat -p <printer_name>'.
    Returns True if available, False otherwise.
    """
    result = subprocess.run(['lpstat', '-p', printer_name], capture_output=True)
    if result.returncode == 0:
        return True
    else:
        logger.error(f"❌ Printer '{printer_name}' is NOT available.")
        return False


def main():
    logger.info("🔥 Starting printer monitoring script...")

    # Initialize caches for settings, target folder, and printer.
    prev_settings = None
    prev_target_folder = None
    prev_printer_name = None

    while True:
        # Load settings from file.
        settings = load_settings()
        if settings is None:
            time.sleep(5)
            continue

        # Log new settings only when they change.
        if prev_settings != settings:
            logger.info(f"🛠 Settings loaded: {settings}")
            prev_settings = settings

        # Resolve and cache the target folder path.
        raw_target_folder = settings.get('target_folder', '~/Downloads')
        target_folder = os.path.expanduser(raw_target_folder)
        if prev_target_folder != target_folder:
            logger.info(f"📂 Resolved target folder path: {target_folder}")
            prev_target_folder = target_folder

        keywords = settings.get('keywords', [])
        use_default_printer = settings.get('use_default_printer', True)
        explicit_printer_name = settings.get('explicit_printer_name', None)

        # Determine which printer to use.
        if use_default_printer:
            current_printer = get_default_printer()
            if not current_printer:
                logger.error("❌ No default printer found. Retrying in 10 seconds...")
                time.sleep(10)
                continue
            if prev_printer_name != current_printer:
                logger.info(f"🖨 Using system default printer: {current_printer}")
                prev_printer_name = current_printer
            printer_name = current_printer
        else:
            if not explicit_printer_name:
                logger.error("❌ Explicit printer name is empty in settings. Retrying in 10 seconds...")
                time.sleep(10)
                continue
            if prev_printer_name != explicit_printer_name:
                logger.info(f"🖨 Using explicitly set printer: {explicit_printer_name}")
                prev_printer_name = explicit_printer_name
            printer_name = explicit_printer_name

        # Verify printer availability.
        if not check_printer_availability(printer_name):
            time.sleep(10)
            continue

        # Ensure the target folder is accessible.
        if not os.path.exists(target_folder) or not os.access(target_folder, os.R_OK):
            logger.error(f"❌ Target folder '{target_folder}' not accessible. Retrying in 10 seconds...")
            time.sleep(10)
            continue

        # Log scanning message using number of keywords only.
        logger.info(f"🔍 Scanning {target_folder} (looking for {len(keywords)} keywords)")

        # Process files matching keywords.
        for file in os.listdir(target_folder):
            if any(keyword in file for keyword in keywords):
                file_path = os.path.join(target_folder, file)
                logger.info(f"📄 Found matching file: {file_path}. Sending to printer.")
                try:
                    print_result = subprocess.run(
                        ['lp', '-d', printer_name, file_path],
                        capture_output=True,
                        text=True
                    )
                    if print_result.returncode == 0:
                        logger.success(f"✅ Successfully printed {file}. Deleting file.")
                        os.remove(file_path)
                    else:
                        logger.error(f"❌ Failed to print {file}. lp output: {print_result.stderr}")
                except Exception as e:
                    logger.exception(f"❌ Exception occurred while printing {file}: {e}")

        logger.debug("⏳ Cycle complete. Sleeping for 3 seconds before next scan.")
        time.sleep(3)


if __name__ == "__main__":
    main()
