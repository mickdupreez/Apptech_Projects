#!/usr/local/bin/python3
# This is the shebang line that tells the system to use the Python 3 interpreter located at /usr/local/bin/python3

"""
Auto-Print Monitor Script (print.py)

This script continuously monitors a specified folder for files whose names contain certain keywords.
When a matching file is found, it is automatically printed using either the system's default printer or
an explicitly specified printer. After successful printing, the file is removed from the folder.
Detailed, emoji-enhanced logging is provided throughout the process to ensure step-by-step traceability.

Usage:
    - Place a valid "print_settings.json" file in the same directory as this script.
    - Customize settings (target folder, keywords, printer settings, and scan interval) in the JSON file.
    - Run the script: python3 print.py

Dependencies:
    - loguru: For advanced logging with custom formatting.
    - wcwidth: To calculate the display width of Unicode strings (ensuring proper emoji alignment).
    - Standard Python libraries: os, subprocess, time, json, sys

Note:
    - Extensive error handling is implemented to manage issues like file access and printer availability.
    - Every function and critical section is documented to ensure clarity and ease of maintenance.
"""
# The above multi-line string provides detailed documentation about the script's purpose, usage, dependencies, and notes.

# Import the os module to interact with the operating system (e.g., file paths, directory listings).
import os
# Import the subprocess module to run external commands (e.g., lpstat and lp for printing).
import subprocess
# Import the time module to handle delays and timestamps.
import time
# Import the json module to load configuration settings from a JSON file.
import json
# Import the sys module for system-specific parameters and functions.
import sys
# Import the logger object from loguru for advanced logging capabilities.
from loguru import logger
# Import the wcswidth function from wcwidth to compute the display width of Unicode strings.
from wcwidth import wcswidth

# Define a constant for the base length of the longest log line, used for visual formatting.
BASE_LONGEST_LINE_LENGTH = 141
# Define an offset compensation value to fine-tune the log line width calculation.
OFFSET_COMPENSATION = 2
# Calculate the final longest line length by adjusting the base value with the compensation.
LONGEST_LINE_LENGTH = BASE_LONGEST_LINE_LENGTH - 13 + OFFSET_COMPENSATION
# The above calculation sets a constant value for formatting log messages with consistent width.

# Define a static prefix example to simulate the width of the log header.
STATIC_PREFIX = "| AUTO_PRINT | 20/3/25 11:17:00.628 | TRACE    | "  # This string is used to determine the display width for log formatting.

def padded_message(record):
    """
    Generate a padded log message to ensure uniform log line lengths.
    
    This function calculates the required padding based on the length of the static prefix and 
    the actual message. The message is padded with spaces so that the total line length aligns 
    with the configured longest line length.
    
    Args:
        record (dict): A dictionary containing the 'message' key with the log message.
    
    Returns:
        str: The padded message with additional spaces appended.
    """
    msg = record['message']  # Extract the log message from the record dictionary.
    prefix_length = wcswidth(STATIC_PREFIX)  # Compute the display width of the static prefix using wcwidth.
    # Calculate the combined width of the prefix and the message (plus one extra space for separation).
    line_length = prefix_length + wcswidth(msg + " ")
    # Determine how many spaces are needed to reach the longest line length, ensuring non-negative value.
    spaces_needed = max(0, LONGEST_LINE_LENGTH - line_length - 1)
    # Return the message concatenated with the necessary number of space characters.
    return f"{msg}{' ' * spaces_needed}"

def load_settings(settings_file=None):
    """
    Load the printing settings from a JSON configuration file.
    
    This function attempts to load a JSON file containing configuration settings. If no file path 
    is provided, it defaults to a file named "print_settings.json" in the same directory as the script.
    It handles errors such as missing files and invalid JSON syntax.
    
    Args:
        settings_file (str, optional): The file path to the JSON settings file. Defaults to None.
    
    Returns:
        dict or None: The settings as a dictionary if successful; otherwise, None.
    """
    # If no settings file is provided, determine the default settings file location.
    if settings_file is None:
        # Get the directory of the current script.
        script_dir = os.path.dirname(os.path.realpath(__file__))
        # Set the path to the default JSON file named "print_settings.json" in the same directory.
        settings_file = os.path.join(script_dir, "print_settings.json")
    try:
        # Open the settings file in read mode.
        with open(settings_file, 'r') as f:
            # Load the JSON content from the file into a Python dictionary.
            settings = json.load(f)
        # Return the loaded settings dictionary.
        return settings
    except FileNotFoundError:
        # Print an error message if the file is not found.
        print(f"❌ ERROR: Settings file not found: {settings_file}")
        # Return None to indicate failure to load settings.
        return None
    except json.JSONDecodeError:
        # Print an error message if the JSON file contains invalid syntax.
        print(f"🚫 ERROR: Invalid JSON syntax in {settings_file}")
        # Return None to indicate failure due to invalid JSON.
        return None

# Load settings early to determine the logging level.
early_settings = load_settings()
# Determine the log level from the settings, defaulting to "TRACE" if not set or if settings are missing.
log_level = early_settings.get("log_level", "TRACE").upper() if early_settings else "TRACE"

logger.remove()  # Remove any default logger configurations to allow custom formatting to be applied.

# Add a logger configuration to output logs to the system's standard output (console).
logger.add(
    sys.stdout,  # Specify that log messages should be sent to the console (stdout).
    colorize=True,  # Enable colored output for better readability.
    format=(
        "<bold><magenta>| AUTO_PRINT |</magenta></bold> "  # Format the static prefix with bold magenta.
        "<bold><cyan>{time:DD/M/YY HH:mm:ss.SSS}</cyan></bold> | "  # Format the timestamp with bold cyan.
        "<bold><level>{level: <8}</level></bold> | {extra[padded]}"  # Format the log level and include the padded message.
    ),
    level=log_level,  # Set the logging level based on the configuration settings.
    enqueue=True  # Enable asynchronous logging to improve performance.
)

# Add another logger configuration to write logs to a file named "printer_monitor.log".
logger.add(
    "printer_monitor.log",  # Specify the log file name.
    mode="w",  # Open the log file in write mode, overwriting the file on each run.
    rotation="5 MB",  # Rotate the log file when it reaches 5 MB in size.
    retention="10 days",  # Retain rotated log files for 10 days before deletion.
    compression="zip",  # Compress rotated log files using zip compression.
    colorize=False,  # Disable colored output for the file logs.
    level=log_level,  # Set the logging level for file logging.
    format="| AUTO_PRINT | {time:DD/M/YY HH:mm:ss.SSS} | {level: <8} | {extra[padded]}",  # Define the log format for file logs.
    enqueue=True  # Enable asynchronous logging for file output as well.
)

def padded_log_method(level):
    """
    Create a wrapper for logging methods that automatically pads messages.
    
    This function returns a wrapper function that formats the message using the provided log level,
    computes the padded version of the message, and then logs it using the bound logger.
    
    Args:
        level (str): The logging level (e.g., "TRACE", "DEBUG", "INFO").
    
    Returns:
        function: A wrapper function that logs messages with the specified level and padded formatting.
    """
    def wrapper(msg, *args, **kwargs):
        # Format the message using Python's str.format() with any additional arguments.
        formatted_message = msg.format(*args)
        # Compute the padded version of the formatted message using the padded_message function.
        padded = padded_message({'message': formatted_message})
        # Bind the padded message as extra context and log the message at the specified log level.
        logger.bind(padded=padded).log(level, msg, *args, **kwargs)
    # Return the wrapper function so it can be used for logging.
    return wrapper

# Create padded log methods for various log levels and attach them as custom methods to the logger.
logger.ptrace = padded_log_method("TRACE")       # For trace-level messages, used for detailed debugging information.
logger.pdebug = padded_log_method("DEBUG")         # For debug-level messages, used for general debugging.
logger.pinfo = padded_log_method("INFO")           # For informational messages about normal operations.
logger.psuccess = padded_log_method("SUCCESS")     # For success messages to indicate successful operations.
logger.pwarning = padded_log_method("WARNING")     # For warning messages about potential issues.
logger.perror = padded_log_method("ERROR")         # For error messages indicating failures.
logger.pcritical = padded_log_method("CRITICAL")   # For critical messages indicating severe issues.

def get_default_printer():
    """
    Detect and return the system's default printer using the 'lpstat' command.
    
    This function attempts to run the 'lpstat -d' command to retrieve the default printer.
    If successful, it extracts and returns the printer name. In case of errors or if no default 
    printer is found, it logs a warning and returns None.
    
    Returns:
        str or None: The default printer name if detected; otherwise, None.
    """
    logger.pdebug("🔍 Detecting system default printer...")  # Log a debug message indicating printer detection.
    logger.ptrace("🛠️ TRACE: Running subprocess: lpstat -d")  # Log a trace message before executing the command.
    try:
        # Execute the 'lpstat -d' command using subprocess.run to get the default printer,
        # capturing the output as text and raising an error if the command fails.
        result = subprocess.run(['lpstat', '-d'], capture_output=True, text=True, check=True)
        # Strip any extra whitespace from the output and take the first part of the output for tracing.
        short_status = result.stdout.strip().split('.')[0]
        logger.ptrace(f"🛠️ TRACE: lpstat OK — {short_status}")  # Log a trace message with the short status output.
        # Extract the printer name by splitting the output string on ': ' and taking the second part.
        printer_name = result.stdout.strip().split(': ')[1]
        logger.psuccess(f"🖨️ Default printer detected: {printer_name}")  # Log a success message with the detected printer.
        # Return the detected default printer name.
        return printer_name
    except (subprocess.CalledProcessError, IndexError):
        # Log a warning message if the lpstat command fails or the output is not as expected.
        logger.pwarning("⚠️ No default printer found, retrying in 10 seconds.")
        # Return None to indicate that no default printer could be detected.
        return None

def check_printer_availability(printer_name):
    """
    Check if the specified printer is available using the 'lpstat' command.
    
    This function runs the 'lpstat -p <printer_name>' command to verify that the printer is active.
    It logs the status and returns True if the printer is available, otherwise logs a warning and returns False.
    
    Args:
        printer_name (str): The name of the printer to check.
    
    Returns:
        bool: True if the printer is available; False otherwise.
    """
    logger.pdebug(f"🔎 Checking printer: {printer_name} availability...")  # Log a debug message indicating printer check.
    logger.ptrace(f"🛠️ TRACE: Running subprocess: lpstat -p {printer_name}")  # Log a trace message before executing the command.
    # Execute the 'lpstat -p' command for the given printer, capturing the output and return code.
    result = subprocess.run(['lpstat', '-p', printer_name], capture_output=True, text=True)
    # Strip any extra whitespace from the output and split it to obtain a short version for logging.
    short_status = result.stdout.strip().split('.')[0]
    logger.ptrace(f"🛠️ TRACE: lpstat OK — {short_status}")  # Log the trace message with the short status.
    if result.returncode == 0:
        # If the command succeeded (return code 0), log a success message indicating printer availability.
        logger.psuccess(f"✅ Printer {printer_name} waiting for files to print.")
        # Return True to indicate that the printer is available.
        return True
    else:
        # If the command did not succeed, log a warning message indicating the printer is not available.
        logger.pwarning(f"⚠️ Printer {printer_name} is not available, retrying...")
        # Return False to indicate that the printer is unavailable.
        return False

def main():
    """
    Main function to run the Auto-Print Monitor service.
    
    This function loads the configuration settings, enters an infinite loop to periodically scan
    a target folder for files matching specified keywords, and attempts to print any matching files
    using either the default printer or an explicitly specified printer. It includes error handling
    for missing settings, printer issues, and folder access problems.
    """
    logger.pinfo("🔥 AUTO_PRINT service started 🚀")  # Log an informational message indicating that the service has started.
    settings = load_settings()  # Load the configuration settings from the JSON file.
    if settings is None:
        # If settings could not be loaded, log a critical error message and exit the function.
        logger.pcritical("❌ Settings load failure at startup. Exiting.")
        return  # Exit the main function due to critical error.

    logger.pinfo("📥 Settings loaded successfully.")  # Log an informational message indicating successful settings load.
    # Expand the target folder path; if not specified in settings, default to the user's Downloads folder.
    target_folder = os.path.expanduser(settings.get('target_folder', '~/Downloads'))
    # Retrieve the list of keywords from settings to search for in file names.
    keywords = settings.get('keywords', [])
    # Determine whether to use the system's default printer based on settings.
    use_default_printer = settings.get('use_default_printer', True)
    # Retrieve the explicit printer name from settings if provided.
    explicit_printer_name = settings.get('explicit_printer_name', None)
    # Get the scan interval (in seconds) from settings to determine how frequently to scan the folder.
    scan_interval = settings.get('scan_interval_seconds', 3)

    cycle_count = 0  # Initialize a counter to keep track of the number of monitoring cycles.
    prev_printer_name = None  # Initialize a variable to track the previously used printer name for logging.

    try:
        # Start an infinite loop to continuously monitor the target folder.
        while True:
            cycle_count += 1  # Increment the cycle counter at the beginning of each cycle.
            logger.pdebug(f"🔄 Starting cycle {cycle_count}")  # Log a debug message indicating the start of a new cycle.
            logger.ptrace(f"🛠️ TRACE: Cycle {cycle_count} start acknowledged")  # Log a trace message for cycle start.

            # Determine which printer to use based on the configuration setting.
            if use_default_printer:
                # Retrieve the default printer name using the get_default_printer() function.
                current_printer = get_default_printer()
                if not current_printer:
                    # If no default printer is detected, log a critical error and wait 10 seconds before retrying.
                    logger.pcritical("❌ Default printer not found. Retrying in 10 seconds...")
                    time.sleep(10)  # Pause execution for 10 seconds.
                    continue  # Skip the rest of the loop and start the next cycle.
                # If the detected default printer has changed from the previous cycle, log the change.
                if prev_printer_name != current_printer:
                    logger.pinfo(f"🖨 Switching to default printer: {current_printer}")
                    prev_printer_name = current_printer  # Update the previous printer name.
                # Set the printer_name variable to the currently detected default printer.
                printer_name = current_printer
            else:
                # If not using the default printer, check for an explicitly specified printer name.
                if not explicit_printer_name:
                    # If no explicit printer name is provided, log an error and wait 10 seconds before retrying.
                    logger.perror("🚫 No explicit printer name configured. Retrying in 10 seconds...")
                    time.sleep(10)  # Pause execution for 10 seconds.
                    continue  # Skip the rest of the loop and start the next cycle.
                # If the explicit printer name has changed from the previous cycle, log the change.
                if prev_printer_name != explicit_printer_name:
                    logger.pinfo(f"🖨 Using explicit printer: {explicit_printer_name}")
                    prev_printer_name = explicit_printer_name  # Update the previous printer name.
                # Set the printer_name variable to the explicit printer name from settings.
                printer_name = explicit_printer_name

            # Check if the chosen printer is currently available.
            if not check_printer_availability(printer_name):
                time.sleep(10)  # Pause execution for 10 seconds before rechecking if the printer is unavailable.
                continue  # Skip the rest of the loop and start the next cycle.

            # Check if the target folder exists and is accessible for reading.
            if not os.path.exists(target_folder) or not os.access(target_folder, os.R_OK):
                # If the folder does not exist or is not readable, log a critical error.
                logger.pcritical(f"🚫 Folder access issue: {target_folder}. Retrying in 10s.")
                time.sleep(10)  # Pause execution for 10 seconds before retrying.
                continue  # Skip the rest of the loop and start the next cycle.
            else:
                # Log a debug message confirming that the target folder is accessible.
                logger.pdebug(f"📂 Folder access check passed: {target_folder}")

            # Log an informational message about scanning the target folder for the specified number of keywords.
            logger.pinfo(f"🔎 Scanning {target_folder} for {len(keywords)} keywords.")
            logger.ptrace(f"🛠️ TRACE: Searching for {len(keywords)} keywords...")  # Log a trace message for scanning.

            files_found = False  # Initialize a flag to indicate whether any matching files are found during this cycle.
            # Iterate over each file in the target folder.
            for file in os.listdir(target_folder):
                # Check if any keyword (case-insensitive) exists within the current file name.
                if any(keyword.lower() in file.lower() for keyword in keywords):
                    # Construct the full path of the file by joining the target folder path with the file name.
                    file_path = os.path.join(target_folder, file)
                    # Log an informational message indicating that a matching file was found and will be printed.
                    logger.pinfo(f"📄 File found: {file}. Printing...")
                    # Log a trace message indicating the command that will be executed to print the file.
                    logger.ptrace(f"🛠️ TRACE: lp call: lp -d {printer_name} {file_path}")
                    files_found = True  # Set the flag to True as a matching file has been found.
                    try:
                        # Execute the 'lp' command using subprocess.run to print the file, capturing output and errors.
                        print_result = subprocess.run(
                            ['lp', '-d', printer_name, file_path],
                            capture_output=True,
                            text=True
                        )
                        # Log a trace message showing the first 80 characters of the command's standard output.
                        logger.ptrace(f"🛠️ TRACE: lp stdout: {print_result.stdout.strip()[:80]}...")
                        # Log a trace message showing the first 80 characters of the command's standard error.
                        logger.ptrace(f"🛠️ TRACE: lp stderr: {print_result.stderr.strip()[:80]}...")
                        if print_result.returncode == 0:
                            # If the print command was successful (return code 0), log a success message.
                            logger.psuccess(f"✅ Printed: {file}. File removed.")
                            os.remove(file_path)  # Remove the file after it has been printed successfully.
                            # Log a trace message indicating that the file has been deleted.
                            logger.ptrace(f"🛠️ TRACE: Deleted {file} post-print.")
                        else:
                            # If the print command failed (non-zero return code), log an error message with the stderr output.
                            logger.perror(f"❌ Print failure for {file}. Reason: {print_result.stderr}")
                    except Exception as e:
                        # If an exception occurs during the print process, log an error message with the exception details.
                        logger.perror(f"🚨 Exception encountered printing {file}: {e}")

            if not files_found:
                # If no matching files were found during this cycle, log an informational message.
                logger.pinfo(f"ℹ️ No files found to print this cycle (cycle {cycle_count}).")
                # Log a trace message to indicate that the cycle completed without printing any files.
                logger.ptrace(f"💓 TRACE: Cycle {cycle_count} idle heartbeat.")

            # Log a debug message indicating that the cycle is complete and the script will sleep before the next cycle.
            logger.pdebug(f"⏳ Cycle {cycle_count} complete. Sleeping for {scan_interval} seconds... 💤")
            # Log a trace message with details about the sleep duration and upcoming cycle.
            logger.ptrace(f"🛠️ TRACE: Sleeping for {scan_interval}s; next cycle {cycle_count + 1}")
            time.sleep(scan_interval)  # Pause execution for the specified scan interval before starting the next cycle.

    except KeyboardInterrupt:
        # Catch a KeyboardInterrupt (e.g., when the user presses Ctrl+C) to allow for graceful shutdown.
        logger.pinfo("👋 Shutdown request received. Exiting cleanly.")
        sys.exit(0)  # Exit the program with a status code of 0 indicating a clean shutdown.

# This conditional ensures that main() is only called when this script is executed directly,
# and not when it is imported as a module in another script.
if __name__ == "__main__":
    main()  # Call the main function to start the Auto-Print Monitor service.
