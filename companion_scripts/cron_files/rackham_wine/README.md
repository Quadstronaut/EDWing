# Elite Dangerous Wine Price Scraper: Crontab Troubleshooting
This document summarizes the steps taken to troubleshoot and fix a common issue where a Python Selenium script, designed to scrape commodity prices from Inara.cz, failed to run when scheduled with crontab.

## The Problem: Headless Browser and Crontab
The crontab environment is different from a user's terminal session. It does not have access to a graphical display or many of the environment variables available in a standard shell. The Python script's use of Selenium, which launches a browser even in headless mode, requires a display environment to function correctly. Without it, the script would fail with a WebDriverException and an error message indicating that Chrome failed to start.

## The Solution: Using Xvfb (X virtual framebuffer)
The most reliable solution is to provide a virtual display environment for the headless browser. We achieved this by using Xvfb, a display server that performs graphical operations in memory, without needing a physical screen.

By running the Python script with xvfb-run, we simulated a graphical environment, allowing the Chrome browser to start and the Selenium script to execute successfully.

sudo apt-get install xvfb