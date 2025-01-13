// Configurable variables
const CONFIG = {
    // Sheet column names
    COLUMNS: {
        EXPIRY_DATE: 'DATA PREVISTA',
        EMAIL: 'RESPONSÁVEL TÉCNICO',
        ITEM_NAME: 'SERVIÇO',
        PROJECT_NAME: 'PROJETO',
        STAGE: 'ETAPA / FAMÍLIA'
    },
    // Email template
    EMAIL: {
        SUBJECT: 'Notificação de Expiração',
        TEMPLATE: `
Olá,

Este é um aviso automático para informar que {itemText}{stageText}{projectText}expira hoje ({date}).

Por favor, tome as ações necessárias.

Atenciosamente,
Sistema Automático de Notificações`,
        MANAGER: 'prof.jansenfaria@gmail.com'
    },
    // Trigger settings
    TRIGGER: {
        FREQUENCY_HOURS: 1,
        MAX_RETRIES: 3, // Number of retries for failed operations
        RETRY_DELAY_MS: 1000 // Delay between retries in milliseconds
    }
};

// Main function to check expiry dates across all sheets
async function checkExpiryDates() {
    try {
        Logger.log('Starting checkExpiryDates function');
        const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
        const sheets = spreadsheet.getSheets();

        // Process sheets in parallel for better performance
        await Promise.all(sheets.map(async sheet => {
            Logger.log(`Processing sheet: ${sheet.getName()}`);
            await processSheet(sheet);
        }));
    } catch (error) {
        Logger.log(`Error in checkExpiryDates: ${error.message}`);
        throw error;
    }
}

// Validate email format
function isValidEmail(email) {
    if (!email || typeof email !== 'string') return false;
    const emailRegex = /^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/;
    return emailRegex.test(email.trim());
}

// Extract and validate multiple emails from a string
function extractEmails(emailString) {
    if (!emailString) return [];
    return emailString.split(/[,;]/)
        .map(email => email.trim())
        .filter(email => isValidEmail(email));
}

// Process individual sheet data
async function processSheet(sheet) {
    try {
        const data = sheet.getDataRange().getValues();
        Logger.log(`Total rows in sheet ${sheet.getName()}: ${data.length}`);

        if (data.length <= 1) {
            Logger.log('Sheet empty or contains only headers - skipping');
            return;
        }

        const headers = data[0];
        const columnIndices = getColumnIndices(headers);

        if (!columnIndices.dateColumnIndex) {
            Logger.log(`Required column "${CONFIG.COLUMNS.EXPIRY_DATE}" not found - skipping sheet`);
            return;
        }

        const today = getTodayMidnight();
        await processRows(data, columnIndices, today);

    } catch (error) {
        Logger.log(`Error processing sheet ${sheet.getName()}: ${error.message}`);
        throw error;
    }
}

// Get column indices from headers
function getColumnIndices(headers) {
    return {
        dateColumnIndex: headers.indexOf(CONFIG.COLUMNS.EXPIRY_DATE),
        emailColumnIndex: headers.indexOf(CONFIG.COLUMNS.EMAIL),
        nameColumnIndex: headers.indexOf(CONFIG.COLUMNS.ITEM_NAME),
        projectColumnIndex: headers.indexOf(CONFIG.COLUMNS.PROJECT_NAME),
        stageColumnIndex: headers.indexOf(CONFIG.COLUMNS.STAGE)
    };
}

// Get today's date at midnight
function getTodayMidnight() {
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    return today;
}

// Process rows with retry mechanism
async function processRows(data, columnIndices, today) {
    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const rowData = extractRowData(row, columnIndices);

        if (!isValidRowData(rowData)) {
            Logger.log(`Invalid data in row ${i} - skipping`);
            continue;
        }

        if (isExpiringToday(rowData.expiryDate, today)) {
            await sendNotificationsWithRetry(rowData);
        }
    }
}

// Extract data from row
function extractRowData(row, indices) {
    return {
        expiryDate: new Date(row[indices.dateColumnIndex]),
        itemName: indices.nameColumnIndex !== -1 ? row[indices.nameColumnIndex] : '',
        projectName: indices.projectColumnIndex !== -1 ? row[indices.projectColumnIndex] : '',
        stage: indices.stageColumnIndex !== -1 ? row[indices.stageColumnIndex] : '',
        emails: indices.emailColumnIndex !== -1 ? extractEmails(row[indices.emailColumnIndex].toString()) : []
    };
}

// Validate row data
function isValidRowData(rowData) {
    return rowData.expiryDate instanceof Date &&
           !isNaN(rowData.expiryDate) &&
           rowData.emails.length > 0;
}

// Check if date is today
function isExpiringToday(expiryDate, today) {
    expiryDate.setHours(0, 0, 0, 0);
    return expiryDate.getTime() === today.getTime();
}

// Send notifications with retry mechanism
async function sendNotificationsWithRetry(rowData) {
    for (const email of rowData.emails) {
        let retries = 0;
        while (retries < CONFIG.TRIGGER.MAX_RETRIES) {
            try {
                await sendExpiryNotification(email, rowData.itemName, rowData.projectName, rowData.stage, rowData.expiryDate);
                break;
            } catch (error) {
                retries++;
                if (retries === CONFIG.TRIGGER.MAX_RETRIES) {
                    Logger.log(`Failed to send email to ${email} after ${retries} attempts: ${error.message}`);
                    throw error;
                }
                await Utilities.sleep(CONFIG.TRIGGER.RETRY_DELAY_MS);
            }
        }
    }
}

// Send individual notification
async function sendExpiryNotification(email, itemName, projectName, stage, expiryDate) {
    Logger.log(`Sending notification to ${email}`);
    const formattedDate = Utilities.formatDate(expiryDate, Session.getScriptTimeZone(), 'dd/MM/yyyy');

    const itemText = itemName ? `o item "${itemName}" ` : '';
    const stageText = stage ? `da etapa "${stage}" ` : '';
    const projectText = projectName ? `do projeto "${projectName}" ` : '';

    const body = CONFIG.EMAIL.TEMPLATE
        .replace('{itemText}', itemText)
        .replace('{projectText}', projectText)
        .replace('{stageText}', stageText)
        .replace('{date}', formattedDate);

    const emailParams = {
        to: email,
        cc: CONFIG.EMAIL.MANAGER,
        subject: CONFIG.EMAIL.SUBJECT,
        body: body
    };

    MailApp.sendEmail(emailParams);
    Logger.log(`Notification sent successfully to ${email}`);
}

// Create time-based trigger
function createTimeDrivenTrigger() {
    Logger.log('Creating time-driven trigger');
    ScriptApp.newTrigger('checkExpiryDates')
        .timeBased()
        .everyHours(CONFIG.TRIGGER.FREQUENCY_HOURS)
        .create();
    Logger.log('Trigger created successfully');
}

// Setup function with error handling
function setup() {
    try {
        Logger.log('Starting setup');
        const triggers = ScriptApp.getProjectTriggers();
        Logger.log(`Removing ${triggers.length} existing triggers`);

        triggers.forEach(trigger => {
            ScriptApp.deleteTrigger(trigger);
        });

        createTimeDrivenTrigger();
        Logger.log('Setup completed successfully');
    } catch (error) {
        Logger.log(`Setup failed: ${error.message}`);
        throw error;
    }
}