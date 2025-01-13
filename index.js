// Configurable variables
const CONFIG = {
    // Sheet column names
    COLUMNS: {
        EXPIRY_DATE: 'Data de Expiração',
        EMAIL: 'Email', 
        ITEM_NAME: 'Nome do Item'
    },
    // Email template
    EMAIL: {
        SUBJECT: 'Notificação de Expiração',
        TEMPLATE: `
Olá,

Este é um aviso automático para informar que o item "{itemName}" expira hoje ({date}).

Por favor, tome as ações necessárias.

Atenciosamente,
Sistema Automático de Notificações`
    },
    // Trigger settings
    TRIGGER: {
        FREQUENCY_HOURS: 1
    }
};

function checkExpiryDates() {
    Logger.log('Starting checkExpiryDates function');
    
    // Get the active spreadsheet and the first sheet
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    Logger.log(`Active sheet name: ${sheet.getName()}`);
    
    // Get all data from the sheet
    const data = sheet.getDataRange().getValues();
    Logger.log(`Total rows in sheet: ${data.length}`);
    
    // Assuming headers are in the first row
    const headers = data[0];
    Logger.log(`Headers found: ${headers.join(', ')}`);
    
    // Find the column indices
    const dateColumnIndex = headers.indexOf(CONFIG.COLUMNS.EXPIRY_DATE);
    const emailColumnIndex = headers.indexOf(CONFIG.COLUMNS.EMAIL);
    const nameColumnIndex = headers.indexOf(CONFIG.COLUMNS.ITEM_NAME);
    
    Logger.log(`Column indices - Date: ${dateColumnIndex}, Email: ${emailColumnIndex}, Name: ${nameColumnIndex}`);
    
    // Verify required columns exist
    if (dateColumnIndex === -1 || emailColumnIndex === -1) {
      Logger.log(`Error: Required columns not found. Please ensure "${CONFIG.COLUMNS.EXPIRY_DATE}" and "${CONFIG.COLUMNS.EMAIL}" columns exist`);
      return;
    }
    
    // Get today's date and reset time to midnight for proper comparison
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    Logger.log(`Today's date (midnight): ${today}`);
    
    // Loop through each row (skip header row)
    for (let i = 1; i < data.length; i++) {
      Logger.log(`Processing row ${i}`);
      const row = data[i];
      const expiryDate = new Date(row[dateColumnIndex]);
      const recipientEmail = row[emailColumnIndex];
      const itemName = nameColumnIndex !== -1 ? row[nameColumnIndex] : 'Item';
      
      Logger.log(`Row ${i} data - Date: ${expiryDate}, Email: ${recipientEmail}, Item: ${itemName}`);
      
      // Skip invalid data
      if (!expiryDate || !recipientEmail) {
        Logger.log(`Skipping row ${i} due to invalid data`);
        continue;
      }
      
      // Reset time part of expiry date for proper comparison
      expiryDate.setHours(0, 0, 0, 0);
      
      // Check if date is today
      if (expiryDate.getTime() === today.getTime()) {
        Logger.log(`Match found for row ${i} - sending notification`);
        try {
          // Send email notification
          sendExpiryNotification(recipientEmail, itemName, expiryDate);
          Logger.log(`Email sent to ${recipientEmail} about expiration of ${itemName}`);
        } catch (error) {
          Logger.log(`Error sending email to ${recipientEmail}: ${error.toString()}`);
        }
      } else {
        Logger.log(`No match for row ${i} - expiry date: ${expiryDate}, today: ${today}`);
      }
    }
    Logger.log('Finished processing all rows');
}

function sendExpiryNotification(email, itemName, expiryDate) {
    Logger.log(`Starting sendExpiryNotification for ${email}`);
    const formattedDate = Utilities.formatDate(expiryDate, Session.getScriptTimeZone(), 'dd/MM/yyyy');
    Logger.log(`Formatted date: ${formattedDate}`);
    
    const body = CONFIG.EMAIL.TEMPLATE
        .replace('{itemName}', itemName)
        .replace('{date}', formattedDate);
    
    Logger.log('Attempting to send email...');
    MailApp.sendEmail({
      to: email,
      subject: CONFIG.EMAIL.SUBJECT,
      body: body
    });
    Logger.log('Email sent successfully');
}

// Create time-based trigger to run hourly
function createTimeDrivenTrigger() {
    Logger.log('Creating new time-driven trigger');
    ScriptApp.newTrigger('checkExpiryDates')
      .timeBased()
      .everyHours(CONFIG.TRIGGER.FREQUENCY_HOURS)
      .create();
    Logger.log('Trigger created successfully');
}

// Initial setup function
function setup() {
    Logger.log('Starting setup function');
    // Remove any existing triggers to avoid duplicates
    const triggers = ScriptApp.getProjectTriggers();
    Logger.log(`Found ${triggers.length} existing triggers`);
    
    triggers.forEach(trigger => {
      Logger.log(`Deleting trigger: ${trigger.getHandlerFunction()}`);
      ScriptApp.deleteTrigger(trigger);
    });
    
    // Create new trigger
    createTimeDrivenTrigger();
    
    Logger.log('Setup completed: Hourly check configured');
}