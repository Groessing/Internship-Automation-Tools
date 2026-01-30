# Internship-Automation-Tools
This repository contains five automation scripts developed during my internship at Greenpeace CEE (July – August).
All credentials, tokens, and passwords have been removed for security.

## schedule-scraping.py
Logs in to TimeTac, navigates to Arbeitszeitmodelle (working schedules), and extracts names and IDs of schedules. The results are written to a Google Sheet.
**Tech:** Python, Selenium

## schedule-duplicator.py
Logs in to TimeTac, navigates to Arbeitszeitmodelle, and automates copying, renaming, and editing schedule templates.
**Tech:** Python, Selenium

## meross-daily-power_consumption.py
Fetches yesterday’s power consumption for each device and writes the data to a Google Sheet.
**Tech:** Python, meross-iot

## recruitee-candidate-sync.js
Reads candidate data from a Google Sheet, creates new candidates via the Recruitee API, updates missing fields, and sets candidates as disqualified when needed.
**Tech:** JavaScript, Google Apps Script, REST API

## workspaceone-device-sync.js
Retrieves all Windows devices that logged in within the last 14 days via the Workspace ONE API, along with installed applications. Writes the data to a Google Sheet and highlights outdated app versions.
**Tech:** JavaScript, Google Apps Script, REST API
