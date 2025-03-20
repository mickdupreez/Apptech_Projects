#!/usr/local/bin/python3
"""
Ultra-locked Auto-Print Monitor Script

This script continuously monitors a specified folder for files whose names contain certain keywords.
When a matching file is found, it is automatically printed using either the system's default printer or
an explicitly specified printer. After successful printing, the file is removed. The script uses detailed
emoji-enhanced logging to provide clear, step-by-step traceability, making it easy for anyone—even those
with no coding experience—to understand its operation.

All sections and lines of this script are explained with comprehensive comments.
"""

# Import necessary modules for file system interaction, process management, time operations,
# JSON configuration handling, system operations, and logging with emoji-aware formatting.
import os                   # Provides functions to interact with the operating system (e.g., file handling).
import subprocess           # Enables running external commands and programs.
import time                 # Offers functions to handle delays and time-related operations.
import json                 # Allows reading and writing JSON formatted data.
import sys                  # Provides access to system-specific parameters and functions.
from loguru import logger   # Advanced logging library for simple, effective logging.
from wcwidth import wcswidth  # Computes the display width of Unicode characters (handles emojis accurately).

# Set up configuration constants for log message formatting.
BASE_LONGEST_LINE_LENGTH = 141    # The ideal total length for a log message line.
OFFSET_COMPENSATION = 2           # Extra characters to adjust for potential display misalignments.
# Calculate the actual maximum line length by applying a fixed offset adjustment.
LONGEST_LINE_LENGTH = BASE_LONGEST_LINE_LENGTH - 13 + OFFSET_COMPENSATION

# Define a constant static prefix that will appear at the start of each log message.
STATIC_PREFIX = "| APPTECH_AUTO_PRINT | 2025-03-20 11:17:00.628 | TRACE    | "

def padded_message(record):
    """
    Enhance a log message by adding appropriate spaces so that visual markers align properly.

    This function receives a log record containing a 'message' field and calculates the necessary
    padding based on the width of a static prefix and the message itself. This is particularly
    important for correct alignment when using emojis, which may have different visual widths.

    Parameters:
        record (dict): Dictionary with a key 'message' containing the log message text.

    Returns:
        str: The modified log message string with extra spaces appended to ensure proper alignment.
    """
    msg = record['message']  # Get the raw log message.
    prefix_length = wcswidth(STATIC_PREFIX)  # Determine how wide the static prefix is when displayed.
    # Calculate the full width of the message line including a starting emoji.
    line_length = prefix_length + wcswidth(f"🔸 {msg} ")
    # Compute the number of spaces required to reach the desired overall line length.
    spaces_needed = max(0, LONGEST_LINE_LENGTH - line_length - 1)
    # Return the message with a starting and ending emoji, padded with the calculated spaces.
    return f"🔸 {msg}{' ' * spaces_needed} 🔸"

def load_settings(settings_file=None):
    """
    Retrieve configuration settings from a JSON file.

    This function attempts to load settings from a file named 'print_settings.json' located in the same
    directory as the script if no specific path is provided. It parses the file into a dictionary. If the
    file cannot be found or contains invalid JSON, it prints an error message and returns None.

    Parameters:
        settings_file (str, optional): The path to the JSON configuration file. Defaults to None.

    Returns:
        dict or None: The settings as a dictionary if loaded successfully, otherwise None.
    """
    if settings_file is None:
        # Determine the directory where the script is located.
        script_dir = os.path.dirname(os.path.realpath(__file__))
        # Build the default path to the settings file.
        settings_file = os.path.join(script_dir, "print_settings.json")

    try:
        # Open the configuration file in read mode.
        with open(settings_file, 'r') as f:
            settings = json.load(f)  # Parse the JSON data.
        return settings  # Return the loaded settings.
    except FileNotFoundError:
        # Inform the user that the configuration file could not be found.
        print(f"❌ ERROR: Settings file not found: {settings_file}")
        return None  # Return None to signal an error.
    except json.JSONDecodeError:
        # Inform the user that the configuration file contains invalid JSON.
        print(f"🚫 ERROR: Invalid JSON syntax in {settings_file}")
        return None  # Return None to signal an error.

# Immediately load the settings to determine log level and other configuration options.
early_settings = load_settings()
# Retrieve the logging level from settings; default to "TRACE" if not specified.
log_level = early_settings.get("log_level", "TRACE").upper() if early_settings else "TRACE"

# Configure the logger with custom formatting and multiple output destinations.
# First, remove any pre-existing logger configurations to start fresh.
logger.remove()

# Add a console logger to output colored log messages to the terminal.
logger.add(
    sys.stdout,  # Output messages to the console.
    colorize=True,  # Enable colored formatting.
    format=(
        "<bold><magenta>| APPTECH_AUTO_PRINT |</magenta></bold> "  # Display a bold, magenta prefix.
        "<bold><cyan>{time:YYYY-MM-DD HH:mm:ss.SSS}</cyan></bold> | "  # Show the timestamp in bold cyan.
        "<bold><level>{level: <8}</level></bold> | {extra[padded]}"     # Present the log level and padded message.
    ),
    level=log_level,  # Use the log level determined from the settings.
    enqueue=True  # Use asynchronous logging to prevent blocking operations.
)

# Add a file logger to write log messages to "printer_monitor.log" with rotation and compression.
logger.add(
    "printer_monitor.log",  # Filename for the log file.
    mode="w",  # Open in write mode to start fresh.
    rotation="5 MB",  # Rotate the log file when it reaches 5 MB.
    retention="10 days",  # Keep rotated log files for 10 days.
    compression="zip",  # Compress rotated log files to save disk space.
    colorize=False,  # Disable colored output for the log file.
    level=log_level,  # Use the log level determined from the settings.
    format="| APPTECH_AUTO_PRINT | {time:YYYY-MM-DD HH:mm:ss.SSS} | {level: <8} | {extra[padded]}",  # File log message format.
    enqueue=True  # Use asynchronous logging.
)

def padded_log_method(level):
    """
    Create a custom logging function for a specific logging level that applies message padding.

    This function returns a wrapper that formats log messages using the padded_message function
    before passing them to the logger. It ensures that all log messages are uniformly formatted.

    Parameters:
        level (str): The logging level (e.g., "DEBUG", "INFO", "ERROR").

    Returns:
        function: A wrapper function that logs messages at the specified level with padding.
    """
    def wrapper(msg, *args, **kwargs):
        # Format the message with any given parameters.
        formatted_message = msg.format(*args)
        # Generate the padded version of the message.
        padded = padded_message({'message': formatted_message})
        # Log the message using the logger, binding the padded message as extra context.
        logger.bind(padded=padded).log(level, msg, *args, **kwargs)
    return wrapper

# Create specialized logger functions for various logging levels with the padded format.
logger.ptrace = padded_log_method("TRACE")     # Detailed trace-level logging.
logger.pdebug = padded_log_method("DEBUG")       # Debug-level logging for development details.
logger.pinfo = padded_log_method("INFO")         # Informational messages.
logger.psuccess = padded_log_method("SUCCESS")   # Indications of successful operations.
logger.pwarning = padded_log_method("WARNING")   # Warning messages about potential issues.
logger.perror = padded_log_method("ERROR")       # Error messages for failed operations.
logger.pcritical = padded_log_method("CRITICAL") # Critical error messages for severe failures.

def get_default_printer():
    """
    Detect the default printer on the system using the 'lpstat -d' command.

    This function runs a system command to retrieve the default printer's name. It logs each step of
    the process and returns the printer name if successful; otherwise, it logs a warning and returns None.

    Returns:
        str or None: The name of the default printer if found; None if not found.
    """
    logger.pdebug("🔍 Detecting system default printer...")  # Inform that printer detection has begun.
    logger.ptrace("🛠️ TRACE: Running subprocess: lpstat -d")  # Detail the exact command to be executed.
    try:
        # Run the command 'lpstat -d' and capture its output.
        result = subprocess.run(['lpstat', '-d'], capture_output=True, text=True, check=True)
        # Process the output to get a concise status message.
        short_status = result.stdout.strip().split('.')[0]
        logger.ptrace(f"🛠️ TRACE: lpstat OK — {short_status}")
        # Extract the printer name from the command output.
        printer_name = result.stdout.strip().split(': ')[1]
        logger.psuccess(f"🖨️ Default printer detected: {printer_name}")  # Log the successful detection.
        return printer_name  # Return the detected printer name.
    except (subprocess.CalledProcessError, IndexError):
        # Log a warning if an error occurs while retrieving the default printer.
        logger.pwarning("⚠️ No default printer found, retrying in 10 seconds.")
        return None  # Return None to indicate failure.

def check_printer_availability(printer_name):
    """
    Verify that a specific printer is available by using the 'lpstat -p' command.

    The function executes the command to check the printer's status and logs the output. It returns True
    if the printer is ready to print, or False if the printer is unavailable.

    Parameters:
        printer_name (str): The name of the printer to check.

    Returns:
        bool: True if the printer is available; otherwise, False.
    """
    logger.pdebug(f"🔎 Checking printer: {printer_name} availability...")  # Indicate that printer availability is being checked.
    logger.ptrace(f"🛠️ TRACE: Running subprocess: lpstat -p {printer_name}")  # Log the command used for checking.
    # Run the command to check the printer status.
    result = subprocess.run(['lpstat', '-p', printer_name], capture_output=True, text=True)
    # Extract a short status description from the command output.
    short_status = result.stdout.strip().split('.')[0]
    logger.ptrace(f"🛠️ TRACE: lpstat OK — {short_status}")
    # Determine if the printer is available based on the return code.
    if result.returncode == 0:
        logger.psuccess(f"✅ Printer {printer_name} Waiting for files to print.")  # Log success if available.
        return True
    else:
        logger.pwarning(f"⚠️ Printer {printer_name} is not available, retrying...")  # Log a warning if unavailable.
        return False

def main():
    """
    Start the auto-print monitoring service.

    This function initializes the monitoring service by loading configuration settings, determining
    the target folder and printer, and entering an endless loop where it checks for files matching
    specified keywords. When a matching file is found, it sends a print command and deletes the file upon
    successful printing. It also handles issues like printer availability and folder access, and it logs
    every operation in detail.
    
    The service runs continuously until the user stops it with a keyboard interrupt (Ctrl+C).
    """
    logger.pinfo("🔥 APPTECH_AUTO_PRINT service started 🚀")  # Log the start of the service.
    settings = load_settings()  # Load configuration settings from the JSON file.
    if settings is None:
        logger.pcritical("❌ Settings load failure at startup. Exiting.")  # Log a critical error if settings fail to load.
        return

    logger.pinfo("📥 Settings loaded successfully.")  # Confirm that settings have been loaded.
    # Convert the target folder path to an absolute path (expanding the '~' if necessary).
    target_folder = os.path.expanduser(settings.get('target_folder', '~/Downloads'))
    keywords = settings.get('keywords', [])  # Get the list of keywords to filter files.
    use_default_printer = settings.get('use_default_printer', True)  # Determine if the default printer should be used.
    explicit_printer_name = settings.get('explicit_printer_name', None)  # Get the explicitly configured printer name.
    scan_interval = settings.get('scan_interval_seconds', 3)  # Define the delay between scan cycles.

    cycle_count = 0           # Counter to track the number of monitoring cycles.
    prev_printer_name = None  # Keep track of the previously used printer name to detect changes.

    try:
        # Begin an endless loop to monitor the target folder continuously.
        while True:
            cycle_count += 1  # Increment the cycle counter for each loop iteration.
            logger.pdebug(f"🔄 Starting cycle {cycle_count}")  # Log the beginning of a new cycle.
            logger.ptrace(f"🛠️ TRACE: Cycle {cycle_count} start acknowledged")  # Detailed trace log for cycle start.

            # Choose the printer to use based on the configuration.
            if use_default_printer:
                current_printer = get_default_printer()  # Retrieve the system default printer.
                if not current_printer:
                    # If no default printer is found, log the error and wait before trying again.
                    logger.pcritical("❌ Default printer not found. Retrying in 10 seconds...")
                    time.sleep(10)
                    continue  # Skip the rest of this cycle.
                if prev_printer_name != current_printer:
                    # If the printer has changed, log the update.
                    logger.pinfo(f"🖨 Switching to default printer: {current_printer}")
                    prev_printer_name = current_printer  # Update the previously used printer.
                printer_name = current_printer  # Set the printer for this cycle.
            else:
                # If not using the default printer, check for an explicit printer configuration.
                if not explicit_printer_name:
                    logger.perror("🚫 No explicit printer name configured. Retrying in 10 seconds...")
                    time.sleep(10)
                    continue  # Skip the rest of this cycle.
                if prev_printer_name != explicit_printer_name:
                    logger.pinfo(f"🖨 Using explicit printer: {explicit_printer_name}")
                    prev_printer_name = explicit_printer_name
                printer_name = explicit_printer_name

            # Check if the selected printer is available.
            if not check_printer_availability(printer_name):
                time.sleep(10)
                continue  # If printer is unavailable, skip this cycle and try again later.

            # Confirm that the target folder exists and is accessible.
            if not os.path.exists(target_folder) or not os.access(target_folder, os.R_OK):
                logger.pcritical(f"🚫 Folder access issue: {target_folder}. Retrying in 10s.")
                time.sleep(10)
                continue  # Skip the cycle if folder access is problematic.
            else:
                logger.pdebug(f"📂 Folder access check passed: {target_folder}")

            # Begin scanning the target folder for files that match the specified keywords.
            logger.pinfo(f"🔎 Scanning {target_folder} for {len(keywords)} keywords.")
            logger.ptrace(f"🛠️ TRACE: Searching for {len(keywords)} keywords...")

            files_found = False  # Flag to track whether any matching files are detected in this cycle.
            # Loop through each file in the target folder.
            for file in os.listdir(target_folder):
                # Check if any of the keywords (case-insensitive) are present in the filename.
                if any(keyword.lower() in file.lower() for keyword in keywords):
                    file_path = os.path.join(target_folder, file)  # Construct the full file path.
                    logger.pinfo(f"📄 File found: {file}. Printing...")  # Log that a matching file was found.
                    logger.ptrace(f"🛠️ TRACE: lp call: lp -d {printer_name} {file_path}")  # Log the print command details.

                    files_found = True  # Indicate that a file was found.
                    try:
                        # Attempt to print the file using the 'lp' command.
                        print_result = subprocess.run(
                            ['lp', '-d', printer_name, file_path],
                            capture_output=True,  # Capture both stdout and stderr.
                            text=True  # Process the output as text.
                        )
                        # Log portions of the output from the print command for debugging.
                        logger.ptrace(f"🛠️ TRACE: lp stdout: {print_result.stdout.strip()[:80]}...")
                        logger.ptrace(f"🛠️ TRACE: lp stderr: {print_result.stderr.strip()[:80]}...")

                        # Check if the print command executed successfully.
                        if print_result.returncode == 0:
                            logger.psuccess(f"✅ Printed: {file}. File removed.")  # Log successful printing.
                            os.remove(file_path)  # Remove the file after printing.
                            logger.ptrace(f"🛠️ TRACE: Deleted {file} post-print.")  # Trace deletion.
                        else:
                            # Log an error message if the print command failed.
                            logger.perror(f"❌ Print failure for {file}. Reason: {print_result.stderr}")
                    except Exception as e:
                        # Catch and log any unexpected exceptions during printing.
                        logger.perror(f"🚨 Exception encountered printing {file}: {e}")

            # If no matching files were found in the current cycle, log an informational message.
            if not files_found:
                logger.pinfo(f"ℹ️ No files found to print this cycle (cycle {cycle_count}).")
                logger.ptrace(f"💓 TRACE: Cycle {cycle_count} idle heartbeat.")

            # Log that the cycle is complete and the script is pausing before the next cycle.
            logger.pdebug(f"⏳ Cycle {cycle_count} complete. Sleeping for {scan_interval} seconds... 💤")
            logger.ptrace(f"🛠️ TRACE: Sleeping for {scan_interval}s; next cycle {cycle_count + 1}")
            time.sleep(scan_interval)  # Pause execution for the configured interval before the next cycle.

    except KeyboardInterrupt:
        # Gracefully handle a keyboard interrupt (Ctrl+C) to stop the script.
        logger.pinfo("👋 Shutdown request received. Exiting cleanly.")
        sys.exit(0)  # Exit the script with a success status.

# If this script is executed directly (and not imported as a module), start the monitoring service.
if __name__ == "__main__":
    main()
