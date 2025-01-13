# Expiry Date Notification System

A Google Apps Script application that automatically sends email notifications when items are due to expire.

## Features

- Monitors a Google Spreadsheet for expiration dates
- Sends automated email notifications on the day items expire
- Configurable column names and email templates
- Runs hourly checks automatically
- Detailed logging for monitoring and debugging

## Setup

1. Create a Google Spreadsheet with the following columns:
   - "Data de Expiração" (Expiry Date)
   - "Email" (Recipient Email)
   - "Nome do Item" (Item Name)

2. Open Script Editor in Google Sheets (Extensions > Apps Script)
3. Copy the code from `index.js` into the script editor
4. Run the `setup()` function to initialize the automatic triggers

## Configuration

The system can be configured through the `CONFIG` object:
