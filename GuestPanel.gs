/**
 * WHITE HOUSE TENANT MANAGEMENT SYSTEM
 * Guest Check-In Processing Panel - GuestPanel.gs
 * 
 * This module provides a user-friendly panel interface to process guest check-ins
 * from Google Form responses and update the Guest Rooms sheet.
 */

const GuestPanel = {

  /**
   * Show the guest check-in processing panel
   */
  showGuestCheckInPanel() {
    try {
      console.log('Opening Guest Check-In Processing Panel...');
      
      const checkIns = this._getGuestCheckIns();
      
      if (checkIns.length === 0) {
        SpreadsheetApp.getUi().alert(
          'No Check-Ins Found',
          'No guest check-ins found in the responses sheet. Make sure the Google Form is linked and has responses.',
          SpreadsheetApp.getUi().ButtonSet.OK
        );
        return;
      }
      
      const html = this._generatePanelHTML(checkIns);
      const htmlOutput = HtmlService.createHtmlOutput(html)
        .setWidth(900)
        .setHeight(700)
        .setTitle('🏨 Process Guest Check-Ins');
      
      SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Process Guest Check-Ins');
      
    } catch (error) {
      console.error('Error showing guest check-in panel:', error);
      SpreadsheetApp.getUi().alert('Error', 'Failed to load guest check-in panel: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
    }
  },

  /**
   * Process the selected check-in and update existing guest room row
   */
  processGuestCheckIn(checkInData) {
    try {
      console.log('Processing guest check-in:', checkInData);
      
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      const guestRoomsSheet = spreadsheet.getSheetByName(SHEET_NAMES.GUEST_ROOMS);
      
      if (!guestRoomsSheet) {
        throw new Error('Guest Rooms sheet not found');
      }
      
      // Parse the check-in data
      const data = JSON.parse(checkInData);
      
      // Find the row with the selected room number
      const roomRowIndex = this._findGuestRoomRow(guestRoomsSheet, data.roomNumber);
      if (roomRowIndex === -1) {
        throw new Error(`Room ${data.roomNumber} not found in Guest Rooms sheet`);
      }
      
      // Get existing room data to preserve some fields
      const existingData = guestRoomsSheet.getRange(roomRowIndex, 1, 1, 24).getValues()[0];
      
      // Calculate total amount
      const nights = parseInt(data.numberOfNights) || 1;
      const dailyRate = this._parseAmount(data.dailyRate);
      const totalAmount = dailyRate * nights;
      
      // Calculate check-out date
      const checkInDate = new Date(data.checkInDate);
      const checkOutDate = new Date(checkInDate);
      checkOutDate.setDate(checkOutDate.getDate() + nights);
      
      // Update the existing row with guest information
      const guestRow = [
        this._generateBookingId(),   // Booking ID
        data.roomNumber,             // Room Number
        existingData[2] || 'Guest Room', // Room Name (keep existing)
        existingData[3] || 'Standard',   // Room Type (keep existing)
        existingData[4] || '2',          // Max Occupancy (keep existing)
        existingData[5] || 'WiFi, TV',   // Amenities (keep existing)
        data.dailyRate,              // Daily Rate
        existingData[7] || '',       // Weekly Rate (keep existing)
        existingData[8] || '',       // Monthly Rate (keep existing)
        'Occupied',                  // Status
        new Date().toLocaleDateString(), // Last Cleaned
        '',                          // Maintenance Notes
        data.checkInDate,            // Check-In Date
        checkOutDate.toLocaleDateString(), // Check-Out Date
        data.numberOfNights,         // Number of Nights
        data.numberOfGuests,         // Number of Guests
        data.guestName,              // Current Guest
        data.purposeOfVisit,         // Purpose of Visit
        data.specialRequests || '',  // Special Requests
        data.source || 'Website',    // Source
        `$${totalAmount}`,           // Total Amount
        data.paymentStatus || 'Confirmed', // Payment Status
        'Checked-In',                // Booking Status
        `Checked in on ${new Date().toLocaleDateString()}` // Notes
      ];
      
      // Update the existing row
      guestRoomsSheet.getRange(roomRowIndex, 1, 1, guestRow.length).setValues([guestRow]);
      
      // Mark the check-in as processed in the responses sheet
      this._markCheckInAsProcessed(data.timestamp, data.guestName);
      
      console.log(`Check-in processed successfully for ${data.guestName} in Room ${data.roomNumber}`);
      return `✅ Check-in confirmed! ${data.guestName} has been checked into Room ${data.roomNumber}.`;
      
    } catch (error) {
      console.error('Error processing guest check-in:', error);
      throw new Error('Failed to process check-in: ' + error.message);
    }
  },

  /**
   * Get guest check-ins from form responses
   * @private
   */
  _getGuestCheckIns() {
    try {
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      
      // Look for the guest check-in form responses sheet
      const sheets = spreadsheet.getSheets();
      let responseSheet = null;
      
      // Find the guest check-in responses sheet
      for (let sheet of sheets) {
        const sheetName = sheet.getName().toLowerCase();
        if (sheetName.includes('guest check-in') || 
            (sheetName.includes('form responses') && sheetName.includes('3'))) {
          responseSheet = sheet;
          break;
        }
      }
      
      if (!responseSheet) {
        // If no specific sheet found, look for "Form Responses 3" (assuming this is the guest form)
        responseSheet = spreadsheet.getSheetByName('Form Responses 3');
      }
      
      if (!responseSheet || responseSheet.getLastRow() <= 1) {
        console.log('No guest check-in responses found');
        return [];
      }
      
      const data = responseSheet.getDataRange().getValues();
      const headers = data[0];
      const checkIns = [];
      
      console.log('Guest check-in response sheet headers:', headers);
      
      // Map form headers to our expected fields
      const headerMap = this._createGuestHeaderMap(headers);
      
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        
        // Skip if already processed
        const processedIndex = headers.indexOf('Processed');
        if (processedIndex !== -1 && row[processedIndex]) {
          continue;
        }
        
        const checkIn = {
          timestamp: row[0], // First column is always timestamp
          guestName: this._getFieldValue(row, headerMap.guestName) || 'Unknown Guest',
          email: this._getFieldValue(row, headerMap.email) || '',
          phone: this._getFieldValue(row, headerMap.phone) || '',
          roomNumber: this._getFieldValue(row, headerMap.roomNumber) || '',
          checkInDate: this._getFieldValue(row, headerMap.checkInDate) || '',
          numberOfNights: this._getFieldValue(row, headerMap.numberOfNights) || '1',
          numberOfGuests: this._getFieldValue(row, headerMap.numberOfGuests) || '1',
          purposeOfVisit: this._getFieldValue(row, headerMap.purposeOfVisit) || '',
          specialRequests: this._getFieldValue(row, headerMap.specialRequests) || '',
          rowIndex: i + 1 // Store row index for processing
        };
        
        checkIns.push(checkIn);
      }
      
      console.log(`Found ${checkIns.length} unprocessed guest check-ins`);
      return checkIns;
      
    } catch (error) {
      console.error('Error getting guest check-ins:', error);
      return [];
    }
  },

  /**
   * Create a mapping between guest form headers and our expected fields
   * @private
   */
  _createGuestHeaderMap(headers) {
    const map = {};
    
    headers.forEach((header, index) => {
      const lowerHeader = header.toLowerCase();
      
      // Be specific about Guest Name
      if (lowerHeader === 'guest name' || lowerHeader.startsWith('guest name')) {
        map.guestName = index;
      } else if (lowerHeader === 'name' && !map.guestName) {
        map.guestName = index;
      } else if (lowerHeader.includes('email')) {
        map.email = index;
      } else if (lowerHeader.includes('phone')) {
        map.phone = index;
      } else if (lowerHeader.includes('room number') || lowerHeader.includes('room')) {
        map.roomNumber = index;
      } else if (lowerHeader.includes('check-in date') || lowerHeader.includes('check in')) {
        map.checkInDate = index;
      } else if (lowerHeader.includes('number of nights') || lowerHeader.includes('nights')) {
        map.numberOfNights = index;
      } else if (lowerHeader.includes('number of guests') || lowerHeader.includes('guests')) {
        map.numberOfGuests = index;
      } else if (lowerHeader.includes('purpose of visit') || lowerHeader.includes('purpose')) {
        map.purposeOfVisit = index;
      } else if (lowerHeader.includes('special requests') || lowerHeader.includes('requests')) {
        map.specialRequests = index;
      }
    });
    
    // Debug logging
    console.log('Guest header mapping created:', map);
    console.log('Available headers:', headers);
    
    return map;
  },

  /**
   * Get available guest rooms from Guest Rooms sheet
   * @private
   */
  _getAvailableGuestRooms() {
    try {
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      const guestRoomsSheet = spreadsheet.getSheetByName(SHEET_NAMES.GUEST_ROOMS);
      
      if (!guestRoomsSheet || guestRoomsSheet.getLastRow() <= 1) {
        return [];
      }
      
      const data = guestRoomsSheet.getDataRange().getValues();
      const headers = data[0];
      const rooms = [];
      
      const roomNumberCol = headers.indexOf('Room Number');
      const dailyRateCol = headers.indexOf('Daily Rate');
      const statusCol = headers.indexOf('Status');
      const roomNameCol = headers.indexOf('Room Name');
      
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (row[roomNumberCol]) {
          const roomNumber = row[roomNumberCol].toString();
          const dailyRate = row[dailyRateCol] || '';
          const status = row[statusCol] || 'Available';
          const roomName = row[roomNameCol] || 'Guest Room';
          
          let statusDisplay = '';
          switch (status.toLowerCase()) {
            case 'occupied':
              statusDisplay = '(Occupied)';
              break;
            case 'maintenance':
              statusDisplay = '(Maintenance)';
              break;
            case 'reserved':
              statusDisplay = '(Reserved)';
              break;
            case 'cleaning':
              statusDisplay = '(Cleaning)';
              break;
            case 'available':
            default:
              statusDisplay = '(Available)';
              break;
          }
          
          rooms.push({
            number: roomNumber,
            name: roomName,
            dailyRate: dailyRate,
            status: status,
            display: `${roomNumber} - ${roomName} ${statusDisplay}`,
            available: status.toLowerCase() === 'available'
          });
        }
      }
      
      return rooms;
    } catch (error) {
      console.error('Error getting available guest rooms:', error);
      return [];
    }
  },

  /**
   * Find the row index for a specific guest room number
   * @private
   */
  _findGuestRoomRow(sheet, roomNumber) {
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const roomNumberCol = headers.indexOf('Room Number');
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][roomNumberCol].toString() === roomNumber.toString()) {
        return i + 1; // Return 1-based row index
      }
    }
    
    return -1; // Room not found
  },

  /**
   * Get field value safely
   * @private
   */
  _getFieldValue(row, index) {
    return (index !== undefined && row[index] !== undefined) ? row[index] : '';
  },

  /**
   * Generate a unique booking ID
   * @private
   */
  _generateBookingId() {
    const prefix = 'BK';
    const timestamp = Date.now().toString().slice(-6);
    return `${prefix}${timestamp}`;
  },

  /**
   * Parse amount from string
   * @private
   */
  _parseAmount(amount) {
    if (!amount) return 0;
    if (typeof amount === 'number') return amount;
    
    const cleanAmount = amount.toString().replace(/[$,\s]/g, '');
    return parseFloat(cleanAmount) || 0;
  },

  /**
   * Generate HTML for the guest check-in processing panel
   * @private
   */
  _generatePanelHTML(checkIns) {
    const checkInsJson = JSON.stringify(checkIns);
    const availableRooms = this._getAvailableGuestRooms();
    const roomsJson = JSON.stringify(availableRooms);
    
    return `
<!DOCTYPE html>
<html>
<head>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        .header { background: #1c4587; color: white; padding: 15px; margin: -20px -20px 20px -20px; }
        .selector-section {
            background: #f8f9fa;
            border: 1px solid #ddd;
            border-radius: 8px;
            padding: 20px;
            margin-bottom: 20px;
        }
        .checkin-card { 
            border: 1px solid #ddd; 
            border-radius: 8px; 
            margin: 15px 0; 
            padding: 15px; 
            background: #f9f9f9; 
            display: none;
        }
        .checkin-header { 
            background: #e8f4fd; 
            padding: 10px; 
            margin: -15px -15px 15px -15px; 
            border-radius: 8px 8px 0 0; 
            font-weight: bold; 
        }
        .checkin-details { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; margin: 10px 0; }
        .checkin-field { margin: 5px 0; }
        .checkin-field strong { color: #1c4587; }
        .approval-section { 
            background: #fff; 
            border: 1px solid #ccc; 
            border-radius: 5px; 
            padding: 15px; 
            margin-top: 15px; 
        }
        .form-group { margin: 20px 0; }
        .form-group label { display: block; font-weight: bold; margin-bottom: 10px; }
        .form-group input, .form-group select { 
            width: 100%; 
            padding: 12px; 
            border: 1px solid #ccc; 
            border-radius: 4px; 
            font-size: 14px;
            margin-bottom: 8px;
        }
        .form-group small {
            color: #666; 
            font-size: 11px; 
            margin-top: 8px; 
            display: block;
            line-height: 1.4;
        }
        .btn { 
            background: #22803c; 
            color: white; 
            padding: 10px 20px; 
            border: none; 
            border-radius: 4px; 
            cursor: pointer; 
            font-size: 14px; 
        }
        .btn:hover { background: #1a6b30; }
        .btn-reject { background: #cc0000; }
        .btn-reject:hover { background: #990000; }
        .status { padding: 10px; margin: 10px 0; border-radius: 4px; }
        .status.success { background: #d4edda; color: #155724; border: 1px solid #c3e6cb; }
        .status.error { background: #f8d7da; color: #721c24; border: 1px solid #f5c6cb; }
        .no-checkins { text-align: center; color: #666; margin: 50px 0; }
        .occupied { color: #cc0000; }
        .maintenance { color: #ff6d00; }
        .available { color: #22803c; }
        .cleaning { color: #ff9800; }
        .instruction { color: #666; font-style: italic; margin-bottom: 20px; text-align: center; }
    </style>
</head>
<body>
    <div class="header">
        <h2>🏨 Process Guest Check-Ins</h2>
        <p>Select a guest to review and process their check-in</p>
    </div>
    
    <div id="status" class="status" style="display: none;"></div>
    
    ${checkIns.length === 0 ? `
        <div class="no-checkins">
            <h3>No Pending Check-Ins</h3>
            <p>All check-ins have been processed or no new check-ins have been submitted.</p>
        </div>
    ` : `
        <div class="selector-section">
            <h3>Select Guest to Check In</h3>
            <div class="instruction">Choose a guest from the dropdown below to view their check-in details and process confirmation.</div>
            <div class="form-group">
                <label for="guest-selector">Guest Name:</label>
                <select id="guest-selector" onchange="showSelectedCheckIn()">
                    <option value="">-- Choose a guest --</option>
                    ${checkIns.map((guest, index) => `
                        <option value="${index}">${guest.guestName} (Check-in: ${new Date(guest.checkInDate).toLocaleDateString()})</option>
                    `).join('')}
                </select>
            </div>
        </div>
        
        ${checkIns.map((guest, index) => `
            <div class="checkin-card" id="checkin-${index}">
                <div class="checkin-header">
                    Check-In Request from ${guest.guestName}
                    <span style="float: right; font-size: 12px;">Submitted: ${new Date(guest.timestamp).toLocaleDateString()}</span>
                </div>
                
                <div class="checkin-details">
                    <div>
                        <div class="checkin-field"><strong>Guest Name:</strong> ${guest.guestName}</div>
                        <div class="checkin-field"><strong>Email:</strong> ${guest.email}</div>
                        <div class="checkin-field"><strong>Phone:</strong> ${guest.phone}</div>
                        <div class="checkin-field"><strong>Check-In Date:</strong> ${guest.checkInDate}</div>
                        <div class="checkin-field"><strong>Number of Nights:</strong> ${guest.numberOfNights}</div>
                    </div>
                    <div>
                        <div class="checkin-field"><strong>Number of Guests:</strong> ${guest.numberOfGuests}</div>
                        <div class="checkin-field"><strong>Requested Room:</strong> ${guest.roomNumber || 'Any available'}</div>
                        <div class="checkin-field"><strong>Purpose of Visit:</strong> ${guest.purposeOfVisit || 'Not specified'}</div>
                        ${guest.specialRequests ? `<div class="checkin-field"><strong>Special Requests:</strong> ${guest.specialRequests}</div>` : ''}
                    </div>
                </div>
                
                <div class="approval-section">
                    <h4 style="margin-bottom: 25px;">Check-In Confirmation</h4>
                    <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 25px; margin-bottom: 25px;">
                        <div class="form-group">
                            <label>Assign Room:</label>
                            <select id="room-${index}" onchange="updateRoomRate(${index})">
                                <option value="">Choose a room...</option>
                                ${availableRooms.map(room => `
                                    <option value="${room.number}" data-rate="${room.dailyRate}" data-name="${room.name}" class="${room.status.toLowerCase()}">${room.display} - ${room.dailyRate}/night</option>
                                `).join('')}
                            </select>
                        </div>
                        <div class="form-group">
                            <label>Daily Rate:</label>
                            <input type="text" id="rate-${index}" placeholder="e.g., $85" required>
                            <small>Rate per night for this guest</small>
                        </div>
                    </div>
                    <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 25px; margin-bottom: 25px;">
                        <div class="form-group">
                            <label>Payment Status:</label>
                            <select id="payment-${index}">
                                <option value="Confirmed">Confirmed</option>
                                <option value="Paid">Paid</option>
                                <option value="Pending">Pending</option>
                                <option value="Deposit Paid">Deposit Paid</option>
                            </select>
                        </div>
                        <div class="form-group">
                            <label>Booking Source:</label>
                            <select id="source-${index}">
                                <option value="Website">Website</option>
                                <option value="Phone">Phone</option>
                                <option value="Email">Email</option>
                                <option value="Walk-in">Walk-in</option>
                                <option value="Booking.com">Booking.com</option>
                                <option value="Airbnb">Airbnb</option>
                                <option value="Referral">Referral</option>
                            </select>
                        </div>
                    </div>
                    <div style="margin-top: 30px;">
                        <button class="btn" onclick="confirmCheckIn(${index})">✅ Confirm Check-In</button>
                        <button class="btn btn-reject" onclick="rejectCheckIn(${index})" style="margin-left: 10px;">❌ Reject Check-In</button>
                    </div>
                </div>
            </div>
        `).join('')}
    `}
    
    <script>
        const checkIns = ${checkInsJson};
        const availableRooms = ${roomsJson};
        
        function showSelectedCheckIn() {
            const selectedIndex = document.getElementById('guest-selector').value;
            
            // Hide all check-in cards
            checkIns.forEach((guest, index) => {
                document.getElementById('checkin-' + index).style.display = 'none';
            });
            
            // Show selected check-in card
            if (selectedIndex !== '') {
                document.getElementById('checkin-' + selectedIndex).style.display = 'block';
            }
        }
        
        function updateRoomRate(index) {
            const roomSelect = document.getElementById('room-' + index);
            const selectedOption = roomSelect.options[roomSelect.selectedIndex];
            if (selectedOption && selectedOption.dataset.rate) {
                // Auto-populate daily rate with room rate
                document.getElementById('rate-' + index).value = selectedOption.dataset.rate;
            }
        }
        
        function confirmCheckIn(index) {
            const guest = checkIns[index];
            const roomNumber = document.getElementById('room-' + index).value;
            const dailyRate = document.getElementById('rate-' + index).value;
            const paymentStatus = document.getElementById('payment-' + index).value;
            const source = document.getElementById('source-' + index).value;
            
            if (!roomNumber || !dailyRate) {
                showStatus('Please fill in all required fields (Room and Daily Rate).', 'error');
                return;
            }
            
            const selectedRoom = availableRooms.find(room => room.number === roomNumber);
            
            if (!selectedRoom) {
                showStatus('Selected room not found.', 'error');
                return;
            }
            
            const checkInData = {
                ...guest,
                roomNumber: roomNumber,
                dailyRate: dailyRate,
                paymentStatus: paymentStatus,
                source: source
            };
            
            showStatus('Processing check-in...', 'success');
            
            google.script.run
                .withSuccessHandler(function(result) {
                    showStatus(result, 'success');
                    // Reset the form
                    document.getElementById('guest-selector').value = '';
                    document.getElementById('checkin-' + index).style.display = 'none';
                    // Remove the processed check-in from the dropdown
                    const option = document.querySelector('#guest-selector option[value="' + index + '"]');
                    if (option) option.remove();
                })
                .withFailureHandler(function(error) {
                    showStatus('Error: ' + error.message, 'error');
                })
                .processGuestCheckIn(JSON.stringify(checkInData));
        }
        
        function rejectCheckIn(index) {
            if (confirm('Are you sure you want to reject this check-in? This action cannot be undone.')) {
                showStatus('Check-in rejected.', 'error');
                // Reset the form and hide the check-in
                document.getElementById('guest-selector').value = '';
                document.getElementById('checkin-' + index).style.display = 'none';
                // Remove the rejected check-in from the dropdown
                const option = document.querySelector('#guest-selector option[value="' + index + '"]');
                if (option) option.remove();
            }
        }
        
        function showStatus(message, type) {
            const status = document.getElementById('status');
            status.textContent = message;
            status.className = 'status ' + type;
            status.style.display = 'block';
            
            setTimeout(() => {
                status.style.display = 'none';
            }, 5000);
        }
    </script>
</body>
</html>
    `;
  },

  /**
   * Mark check-in as processed in the responses sheet
   * @private
   */
  _markCheckInAsProcessed(timestamp, guestName) {
    try {
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      const sheets = spreadsheet.getSheets();
      let responseSheet = null;
      
      // Find the guest check-in form responses sheet
      for (let sheet of sheets) {
        const sheetName = sheet.getName().toLowerCase();
        if (sheetName.includes('guest check-in') || 
            (sheetName.includes('form responses') && sheetName.includes('3'))) {
          responseSheet = sheet;
          break;
        }
      }
      
      if (!responseSheet) {
        responseSheet = spreadsheet.getSheetByName('Form Responses 3');
      }
      
      if (!responseSheet) return;
      
      const data = responseSheet.getDataRange().getValues();
      const headers = data[0];
      
      // Add "Processed" column if it doesn't exist
      let processedCol = headers.indexOf('Processed');
      if (processedCol === -1) {
        processedCol = headers.length;
        responseSheet.getRange(1, processedCol + 1).setValue('Processed');
      }
      
      // Find and mark the specific check-in as processed
      for (let i = 1; i < data.length; i++) {
        if (data[i][0].toString() === timestamp.toString()) {
          responseSheet.getRange(i + 1, processedCol + 1).setValue(`Checked In - ${new Date().toLocaleDateString()}`);
          break;
        }
      }
      
    } catch (error) {
      console.error('Error marking check-in as processed:', error);
    }
  }
};

/**
 * Wrapper function for menu integration
 */
function showGuestCheckInPanel() {
  return GuestPanel.showGuestCheckInPanel();
}

/**
 * Server-side function called from the HTML panel
 */
function processGuestCheckIn(checkInData) {
  return GuestPanel.processGuestCheckIn(checkInData);
}
