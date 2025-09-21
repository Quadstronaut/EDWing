#!/bin/bash

# Define variables
SCRIPT_DIR="/home/quadstronaut/cron_files/rackham_wine"
PYTHON_SCRIPT="rackham_wine.py"
LOG_FILE="crontab.log"
PYTHON_INTERPRETER="/usr/bin/python3"

# Set the home directory for the cron environment
export HOME="/home/quadstronaut/cron_files/rackham_wine"

# Navigate to the script's directory
cd "$SCRIPT_DIR" || { echo "Failed to change directory to $SCRIPT_DIR" >&2; exit 1; }

# Run the Python script with xvfb-run
xvfb-run -a "$PYTHON_INTERPRETER" "$SCRIPT_DIR/$PYTHON_SCRIPT" >> "$SCRIPT_DIR/$LOG_FILE" 2>&1
