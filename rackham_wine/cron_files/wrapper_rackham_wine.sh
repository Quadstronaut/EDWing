#!/bin/bash

# Define variables for key directories and files
SCRIPT_DIR="/home/foobar/cron_files/rackham_wine"
PYTHON_SCRIPT="rackham_wine.py"
LOG_FILE="crontab.log"
PYTHON_INTERPRETER="/usr/bin/python3"

# Set the home directory for the cron environment
export HOME="/home/foobar"

# Navigate to the script's directory using the variable
cd "$SCRIPT_DIR" || { echo "Failed to change directory to $SCRIPT_DIR" >&2; exit 1; }

# Run the Python script using the variables
"$PYTHON_INTERPRETER" "$SCRIPT_DIR/$PYTHON_SCRIPT" >> "$SCRIPT_DIR/$LOG_FILE" 2>&1
