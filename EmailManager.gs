/**
 * WHITE HOUSE TENANT MANAGEMENT SYSTEM
 * Email Manager - EmailManager.gs
 * 
 * This module handles all email communications with updated timing:
 * - Reminders: One week before due date (25th)
 * - Invoices: On due date (1st)
 * - Late alerts: Day after due date (2nd) and weekly (9th, 16th, 23rd)
 */

const EmailManager = {

  /**
   * Send rent reminders to all active tenants (one week before due date)
   */
  sendRentReminders() {
    try {
      console.log('Sending rent reminders...');
      
      const tenants = this._getActiveTenants();
      let sentCount = 0;
      
      tenants.forEach(tenant => {
        if (tenant.email && tenant.name) {
          const subject = `Rent Reminder - Room ${tenant.roomNumber}`;
          const body = this._createRentReminderEmail(tenant);
          
          try {
            GmailApp.sendEmail(tenant.email, subject, body);
            sentCount++;
            console.log(`Rent reminder sent to ${tenant.name} (${tenant.email})`);
          } catch (emailError) {
            console.error(`Failed to send reminder to ${tenant.email}:`, emailError);
          }
        }
      });
      
      console.log(`Rent reminders sent to ${sentCount} tenants`);
      return `Rent reminders sent to ${sentCount} tenants`;
    } catch (error) {
      console.error('Error sending rent reminders:', error);
      throw error;
    }
  },

  /**
   * Send late payment alerts to tenants with overdue payments
   * Now runs day after due date and weekly thereafter
   */
  sendLatePaymentAlerts() {
    try {
      console.log('Sending late payment alerts...');
      
      const tenants = this._getActiveTenants();
      const currentDate = new Date();
      let sentCount = 0;
      
      tenants.forEach(tenant => {
        if (this._isPaymentLate(tenant, currentDate) && tenant.email && tenant.name) {
          const subject = `Payment Overdue Notice - Room ${tenant.roomNumber}`;
          const body = this._createLatePaymentEmail(tenant);
          
          try {
            GmailApp.sendEmail(tenant.email, subject, body);
            sentCount++;
            console.log(`Late payment alert sent to ${tenant.name} (${tenant.email})`);
          } catch (emailError) {
            console.error(`Failed to send late payment alert to ${tenant.email}:`, emailError);
          }
        }
      });
      
      console.log(`Late payment alerts sent to ${sentCount} tenants`);
      return `Late payment alerts sent to ${sentCount} tenants`;
    } catch (error) {
      console.error('Error sending late payment alerts:', error);
      throw error;
    }
  },

  /**
   * Send monthly invoices to all active tenants (on due date)
   */
  sendMonthlyInvoices() {
    try {
      console.log('Sending monthly invoices...');
      
      const tenants = this._getActiveTenants();
      let sentCount = 0;
      
      tenants.forEach(tenant => {
        if (tenant.email && tenant.name) {
          const subject = `Monthly Rent Invoice - Room ${tenant.roomNumber}`;
          const body = this._createMonthlyInvoiceEmail(tenant);
          
          try {
            GmailApp.sendEmail(tenant.email, subject, body);
            sentCount++;
            console.log(`Monthly invoice sent to ${tenant.name} (${tenant.email})`);
          } catch (emailError) {
            console.error(`Failed to send invoice to ${tenant.email}:`, emailError);
          }
        }
      });
      
      console.log(`Monthly invoices sent to ${sentCount} tenants`);
      return `Monthly invoices sent to ${sentCount} tenants`;
    } catch (error) {
      console.error('Error sending monthly invoices:', error);
      throw error;
    }
  },

  /**
   * Test function to preview rent reminder email
   */
  testRentReminder() {
    const tenants = this._getActiveTenants();
    if (tenants.length > 0) {
      const testTenant = tenants[0];
      console.log('Testing rent reminder with tenant:', testTenant.name);
      
      const subject = `TEST - Rent Reminder - Room ${testTenant.roomNumber}`;
      const body = this._createRentReminderEmail(testTenant);
      
      console.log('Subject:', subject);
      console.log('Body:', body);
      
      return 'Test rent reminder created (check logs for content)';
    } else {
      return 'No active tenants found for testing';
    }
  },

  /**
   * Get all active tenants from the Tenant sheet
   * @private
   */
  _getActiveTenants() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.TENANT);
    if (!sheet) {
      throw new Error('Tenant sheet not found');
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const tenants = [];
    
    // Find column indices
    const roomNumberCol = headers.indexOf('Room Number');
    const rentalPriceCol = headers.indexOf('Rental Price');
    const negotiatedPriceCol = headers.indexOf('Negotiated Price');
    const nameCol = headers.indexOf('Current Tenant Name');
    const emailCol = headers.indexOf('Tenant Email');
    const phoneCol = headers.indexOf('Tenant Phone');
    const statusCol = headers.indexOf('Room Status');
    const lastPaymentCol = headers.indexOf('Last Payment Date');
    const paymentStatusCol = headers.indexOf('Payment Status');
    
    // Process each row (skip header)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      // Only include active tenants (not vacant rooms)
      if (row[statusCol] && row[statusCol].toString().toLowerCase() === 'occupied' && row[nameCol]) {
        tenants.push({
          roomNumber: row[roomNumberCol],
          rentalPrice: row[rentalPriceCol],
          negotiatedPrice: row[negotiatedPriceCol],
          name: row[nameCol],
          email: row[emailCol],
          phone: row[phoneCol],
          status: row[statusCol],
          lastPaymentDate: row[lastPaymentCol],
          paymentStatus: row[paymentStatusCol]
        });
      }
    }
    
    return tenants;
  },

  /**
   * Check if a tenant's payment is late (updated logic)
   * Now considers payment late if:
   * - Today is after the 1st (due date) AND
   * - Payment status is 'Late' or 'Overdue' OR
   * - Last payment date is before current month's 1st
   * @private
   */
  _isPaymentLate(tenant, currentDate) {
    const currentDay = currentDate.getDate();
    
    // Only consider late if we're past the due date (1st of month)
    if (currentDay < 2) return false;
    
    // Check payment status
    if (tenant.paymentStatus === 'Late' || tenant.paymentStatus === 'Overdue') {
      return true;
    }
    
    // Check if last payment was before this month's due date
    if (tenant.lastPaymentDate) {
      const lastPayment = new Date(tenant.lastPaymentDate);
      const currentMonth = currentDate.getMonth();
      const currentYear = currentDate.getFullYear();
      const thisMonthsDueDate = new Date(currentYear, currentMonth, 1);
      
      return lastPayment < thisMonthsDueDate;
    }
    
    // No payment date recorded, consider late
    return true;
  },

  /**
   * Create rent reminder email content (sent one week before due date)
   * @private
   */
  _createRentReminderEmail(tenant) {
    const nextMonth = new Date();
    nextMonth.setMonth(nextMonth.getMonth() + 1);
    const monthName = nextMonth.toLocaleString('default', { month: 'long', year: 'numeric' });
    const rentAmount = tenant.negotiatedPrice || tenant.rentalPrice || 'N/A';
    
    return `
Dear ${tenant.name},

This is a friendly reminder that your rent payment for ${monthName} is due on the 1st of the month.

Room Details:
- Room Number: ${tenant.roomNumber}
- Monthly Rent: ${rentAmount}
- Due Date: ${monthName} 1st

Please ensure your payment is submitted by the due date. This is just a courtesy reminder to help you stay on track.

If you have already made your payment, please disregard this message.

Thank you for being a valued tenant!

Best regards,
${EMAIL_CONFIG.MANAGEMENT_TEAM}
${EMAIL_CONFIG.PROPERTY_NAME}
    `.trim();
  },

  /**
   * Create late payment email content (sent after due date)
   * @private
   */
  _createLatePaymentEmail(tenant) {
    const currentMonth = new Date().toLocaleString('default', { month: 'long', year: 'numeric' });
    const rentAmount = tenant.negotiatedPrice || tenant.rentalPrice || 'N/A';
    const today = new Date();
    const dayOfMonth = today.getDate();
    
    let urgencyLevel = '';
    if (dayOfMonth === 2) {
      urgencyLevel = 'Your rent payment is now one day overdue.';
    } else if (dayOfMonth <= 9) {
      urgencyLevel = 'Your rent payment is now over one week overdue.';
    } else if (dayOfMonth <= 16) {
      urgencyLevel = 'Your rent payment is now over two weeks overdue.';
    } else {
      urgencyLevel = 'Your rent payment is now over three weeks overdue.';
    }
    
    return `
Dear ${tenant.name},

${urgencyLevel}

Room Details:
- Room Number: ${tenant.roomNumber}
- Monthly Rent: ${rentAmount}
- Amount Due: ${rentAmount}
- Original Due Date: ${currentMonth} 1st

Please submit your payment immediately to bring your account current.

If you have already made your payment or are experiencing financial difficulties, please contact us immediately to discuss the situation.

We value you as a tenant and want to work with you to resolve this matter promptly.

Best regards,
${EMAIL_CONFIG.MANAGEMENT_TEAM}
${EMAIL_CONFIG.PROPERTY_NAME}
    `.trim();
  },

  /**
   * Create monthly invoice email content (sent on due date)
   * @private
   */
  _createMonthlyInvoiceEmail(tenant) {
    const currentDate = new Date();
    const currentMonth = currentDate.toLocaleString('default', { month: 'long', year: 'numeric' });
    const rentAmount = tenant.negotiatedPrice || tenant.rentalPrice || 'N/A';
    const dueDate = new Date(currentDate.getFullYear(), currentDate.getMonth(), 1);
    
    return `
Dear ${tenant.name},

Your monthly rent is due today. Please find your invoice details below:

RENT INVOICE - ${currentMonth}
Room Number: ${tenant.roomNumber}
Tenant: ${tenant.name}
Amount Due: ${rentAmount}
Due Date: ${dueDate.toLocaleDateString()}

Please ensure payment is submitted today to avoid your account becoming overdue.

If you have any questions about this invoice, please don't hesitate to contact us.

Thank you!

Best regards,
${EMAIL_CONFIG.MANAGEMENT_TEAM}
${EMAIL_CONFIG.PROPERTY_NAME}
    `.trim();
  }
};
