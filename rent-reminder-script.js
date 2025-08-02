/**
 * Enhanced Tenant Rent Reminder Script with HTML Email Template
 * This script checks for upcoming rent due dates and sends professional HTML reminder emails
 */

// CONFIGURATION - Update these values for your mom's property management
const CONFIG = {
  landlordName: "Margaret", // Replace with actual name
  landlordEmail: "destade45@example.com", // Replace with actual email
  landlordPhone: "(555) 123-4567", // Replace with actual phone
  propertyName: "J5 Property Management Co.", // Replace with property name
  reminderThreshold: 5 // Days before rent due date to send reminder
};

function checkRentReminders() {
  // Get the active spreadsheet and sheet
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getActiveSheet();
  
  // Get all data from the sheet
  const data = sheet.getDataRange().getValues();
  
  // Find column indices from headers
  const headers = data[0];
  const tenantNameCol = headers.indexOf('Tenant Name');
  const phoneCol = headers.indexOf('Phone Number');
  const rentAmountCol = headers.indexOf('Rent Amount');
  const amountPaidCol = headers.indexOf('Amount Paid');
  const rentBalanceCol = headers.indexOf('Rent Balance');
  const rentDueDateCol = headers.indexOf('Rent Due Date');
  const daysUntilExpirationCol = headers.indexOf('Days Until Expiration');
  const tenantEmailCol = headers.indexOf('Tenant Email');
  const statusCol = headers.indexOf('Status');
  
  // Validate required columns
  if (tenantNameCol === -1 || rentDueDateCol === -1 || tenantEmailCol === -1 || statusCol === -1) {
    console.error('Required columns not found. Please check column headers.');
    return;
  }
  
  const today = new Date();
  
  // Process each tenant row
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    
    // Skip empty rows
    if (!row[tenantNameCol] || !row[tenantEmailCol] || !row[rentDueDateCol]) {
      continue;
    }
    
    const tenantName = row[tenantNameCol];
    const tenantEmail = row[tenantEmailCol];
    const rentDueDate = new Date(row[rentDueDateCol]);
    const currentStatus = row[statusCol];
    const rentAmount = row[rentAmountCol] || 0;
    const rentBalance = row[rentBalanceCol] || 0;
    
    // Calculate days until rent is due
    const timeDiff = rentDueDate.getTime() - today.getTime();
    const daysUntilDue = Math.ceil(timeDiff / (1000 * 3600 * 24));
    
    // Check if reminder should be sent
    if (daysUntilDue <= CONFIG.reminderThreshold && daysUntilDue >= 0 && currentStatus !== 'Sent') {
      try {
        // Send HTML reminder email
        sendHtmlReminderEmail(tenantName, tenantEmail, rentDueDate, rentAmount, rentBalance, daysUntilDue);
        
        // Update status to 'Sent'
        sheet.getRange(i + 1, statusCol + 1).setValue('Sent');
        
        console.log(`HTML reminder sent to ${tenantName} (${tenantEmail})`);
        
      } catch (error) {
        console.error(`Failed to send reminder to ${tenantName}: ${error.message}`);
      }
    }
  }
}

/**
 * Sends a professional HTML rent reminder email
 */
function sendHtmlReminderEmail(tenantName, tenantEmail, rentDueDate, rentAmount, rentBalance, daysUntilDue) {
  const subject = `Rent Reminder - Payment Due ${formatDate(rentDueDate)}`;
  
  // Get the HTML template
  let htmlTemplate = getHtmlEmailTemplate();
  
  // Determine time description
  let timeDescription;
  if (daysUntilDue === 0) {
    timeDescription = `today (${formatDate(rentDueDate)})`;
  } else if (daysUntilDue === 1) {
    timeDescription = `tomorrow (${formatDate(rentDueDate)})`;
  } else {
    timeDescription = `in ${daysUntilDue} days on ${formatDate(rentDueDate)}`;
  }
  
  // Replace placeholders in template
  htmlTemplate = htmlTemplate
    .replace(/\[TENANT_NAME\]/g, tenantName)
    .replace(/\[TIME_DESCRIPTION\]/g, timeDescription)
    .replace(/\[DUE_DATE\]/g, formatDate(rentDueDate))
    .replace(/\[RENT_AMOUNT\]/g, formatCurrency(rentAmount))
    .replace(/\[RENT_BALANCE\]/g, formatCurrency(rentBalance))
    .replace(/\[DAYS_UNTIL_DUE\]/g, daysUntilDue)
    .replace(/\[LANDLORD_NAME\]/g, CONFIG.landlordName)
    .replace(/\[LANDLORD_EMAIL\]/g, CONFIG.landlordEmail)
    .replace(/\[LANDLORD_PHONE\]/g, CONFIG.landlordPhone)
    .replace(/\[PROPERTY_NAME\]/g, CONFIG.propertyName)
    .replace(/\[BALANCE_DISPLAY\]/g, rentBalance > 0 ? 'flex' : 'none');
  
  // Create plain text version as fallback
  const plainTextBody = createPlainTextVersion(tenantName, timeDescription, rentAmount, rentBalance, daysUntilDue);
  
  // Send the email
  MailApp.sendEmail({
    to: tenantEmail,
    subject: subject,
    body: plainTextBody,
    htmlBody: htmlTemplate
  });
}

/**
 * Returns the HTML email template
 */
function getHtmlEmailTemplate() {
  return `<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Rent Reminder</title>
    <style>
        body, table, td, p, a, li, blockquote {
            -webkit-text-size-adjust: 100%;
            -ms-text-size-adjust: 100%;
        }
        
        table, td {
            mso-table-lspace: 0pt;
            mso-table-rspace: 0pt;
        }
        
        img {
            -ms-interpolation-mode: bicubic;
            border: 0;
            height: auto;
            line-height: 100%;
            outline: none;
            text-decoration: none;
        }
        
        body {
            margin: 0 !important;
            padding: 0 !important;
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background-color: #f4f6f9;
        }
        
        .email-container {
            max-width: 600px;
            margin: 0 auto;
            background-color: #ffffff;
        }
        
        .header {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            padding: 30px 20px;
            text-align: center;
        }
        
        .header h1 {
            color: #ffffff;
            margin: 0;
            font-size: 28px;
            font-weight: 300;
            letter-spacing: 1px;
        }
        
        .header-subtitle {
            color: rgba(255, 255, 255, 0.9);
            margin: 8px 0 0 0;
            font-size: 16px;
            font-weight: 300;
        }
        
        .content {
            padding: 40px 30px;
        }
        
        .greeting {
            font-size: 18px;
            color: #333333;
            margin: 0 0 20px 0;
            font-weight: 500;
        }
        
        .message {
            font-size: 16px;
            color: #555555;
            line-height: 1.6;
            margin: 0 0 30px 0;
        }
        
        .rent-details {
            background-color: #f8f9fc;
            border: 1px solid #e3e8ee;
            border-radius: 8px;
            padding: 25px;
            margin: 30px 0;
        }
        
        .rent-details h3 {
            color: #333333;
            margin: 0 0 20px 0;
            font-size: 18px;
            font-weight: 600;
            border-bottom: 2px solid #667eea;
            padding-bottom: 8px;
        }
        
        .detail-row {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin: 12px 0;
            padding: 8px 0;
            border-bottom: 1px solid #eef1f5;
        }
        
        .detail-row:last-child {
            border-bottom: none;
            margin-bottom: 0;
        }
        
        .detail-label {
            font-weight: 600;
            color: #666666;
            font-size: 14px;
        }
        
        .detail-value {
            font-weight: 700;
            color: #333333;
            font-size: 16px;
        }
        
        .amount-due {
            color: #e74c3c;
            font-size: 20px;
        }
        
        .due-date {
            color: #f39c12;
        }
        
        .cta-section {
            text-align: center;
            margin: 35px 0;
        }
        
        .cta-button {
            display: inline-block;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: #ffffff !important;
            text-decoration: none;
            padding: 15px 30px;
            border-radius: 6px;
            font-weight: 600;
            font-size: 16px;
            letter-spacing: 0.5px;
            box-shadow: 0 4px 15px rgba(102, 126, 234, 0.3);
        }
        
        .contact-section {
            background-color: #fafbfc;
            border-radius: 8px;
            padding: 25px;
            margin: 30px 0;
            text-align: center;
        }
        
        .contact-title {
            color: #333333;
            font-size: 16px;
            font-weight: 600;
            margin: 0 0 15px 0;
        }
        
        .contact-info {
            color: #666666;
            font-size: 14px;
            line-height: 1.5;
            margin: 8px 0;
        }
        
        .contact-info a {
            color: #667eea;
            text-decoration: none;
            font-weight: 500;
        }
        
        .footer {
            background-color: #f8f9fc;
            padding: 25px;
            text-align: center;
            border-top: 1px solid #e3e8ee;
        }
        
        .footer-text {
            color: #888888;
            font-size: 14px;
            line-height: 1.5;
            margin: 0;
        }
        
        @media only screen and (max-width: 600px) {
            .email-container {
                width: 100% !important;
                max-width: 100% !important;
            }
            
            .content {
                padding: 25px 20px !important;
            }
            
            .rent-details {
                padding: 20px 15px !important;
            }
            
            .header {
                padding: 25px 15px !important;
            }
            
            .header h1 {
                font-size: 24px !important;
            }
            
            .detail-row {
                flex-direction: column;
                align-items: flex-start;
                gap: 5px;
            }
            
            .detail-value {
                font-size: 18px !important;
            }
            
            .cta-button {
                padding: 12px 25px !important;
                font-size: 15px !important;
            }
        }
    </style>
</head>
<body>
    <div class="email-container">
        <div class="header">
            <h1>Rent Reminder</h1>
            <p class="header-subtitle">Payment Due Soon</p>
        </div>
        
        <div class="content">
            <p class="greeting">Dear [TENANT_NAME],</p>
            
            <p class="message">
                This is a friendly reminder that your rent payment is due <strong>[TIME_DESCRIPTION]</strong>. 
                We appreciate your prompt attention to this matter.
            </p>
            
            <div class="rent-details">
                <h3>Payment Details</h3>
                
                <div class="detail-row">
                    <span class="detail-label">Due Date:</span>
                    <span class="detail-value due-date">[DUE_DATE]</span>
                </div>
                
                <div class="detail-row">
                    <span class="detail-label">Rent Amount:</span>
                    <span class="detail-value amount-due">$[RENT_AMOUNT]</span>
                </div>
                
                <div class="detail-row" style="display: [BALANCE_DISPLAY];">
                    <span class="detail-label">Outstanding Balance:</span>
                    <span class="detail-value amount-due">$[RENT_BALANCE]</span>
                </div>
                
                <div class="detail-row">
                    <span class="detail-label">Days Remaining:</span>
                    <span class="detail-value">[DAYS_UNTIL_DUE] days</span>
                </div>
            </div>
            
            <div class="cta-section">
                <a href="mailto:[LANDLORD_EMAIL]?subject=Rent Payment - [TENANT_NAME]" class="cta-button">
                    Contact for Payment Details
                </a>
            </div>
            
            <p class="message">
                Please ensure your payment is submitted on time to avoid any late fees. 
                If you have already made your payment, please disregard this reminder.
            </p>
            
            <div class="contact-section">
                <h4 class="contact-title">Questions or Concerns?</h4>
                <p class="contact-info">
                    <strong>[LANDLORD_NAME]</strong><br>
                    <a href="mailto:[LANDLORD_EMAIL]">[LANDLORD_EMAIL]</a><br>
                    <a href="tel:[LANDLORD_PHONE]">[LANDLORD_PHONE]</a>
                </p>
            </div>
        </div>
        
        <div class="footer">
            <p class="footer-text">
                This is an automated reminder. Thank you for being a valued tenant.<br>
                <small>© 2024 [PROPERTY_NAME]. All rights reserved.</small>
            </p>
        </div>
    </div>
</body>
</html>`;
}

/**
 * Creates a plain text version for email clients that don't support HTML
 */
function createPlainTextVersion(tenantName, timeDescription, rentAmount, rentBalance, daysUntilDue) {
  let plainText = `Dear ${tenantName},\n\n`;
  plainText += `This is a friendly reminder that your rent payment is due ${timeDescription}.\n\n`;
  plainText += `Payment Details:\n`;
  plainText += `• Amount Due: $${formatCurrency(rentAmount)}\n`;
  
  if (rentBalance > 0) {
    plainText += `• Outstanding Balance: $${formatCurrency(rentBalance)}\n`;
  }
  
  plainText += `• Days Remaining: ${daysUntilDue} days\n\n`;
  plainText += `Please ensure your payment is submitted on time to avoid any late fees.\n\n`;
  plainText += `If you have any questions or concerns, please contact:\n`;
  plainText += `${CONFIG.landlordName}\n`;
  plainText += `Email: ${CONFIG.landlordEmail}\n`;
  plainText += `Phone: ${CONFIG.landlordPhone}\n\n`;
  plainText += `Thank you for being a valued tenant.\n\n`;
  plainText += `Best regards,\n${CONFIG.landlordName}`;
  
  return plainText;
}

/**
 * Formats a date to a readable string
 */
function formatDate(date) {
  const options = { 
    year: 'numeric', 
    month: 'long', 
    day: 'numeric' 
  };
  return date.toLocaleDateString('en-US', options);
}

/**
 * Formats currency values
 */
function formatCurrency(amount) {
  return Number(amount).toLocaleString('en-US', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  });
}

/**
 * Sets up a daily trigger to run the rent reminder check
 */
function setupDailyTrigger() {
  // Delete existing triggers to avoid duplicates
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'checkRentReminders') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // Create new daily trigger
  ScriptApp.newTrigger('checkRentReminders')
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .create();
    
  console.log('Daily trigger set up successfully. HTML emails will be sent daily at 9 AM.');
}

/**
 * Manual test function
 */
function testReminderScript() {
  console.log('Running manual test with HTML emails...');
  checkRentReminders();
  console.log('Test completed. Check the logs and your email for HTML results.');
}
