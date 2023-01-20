//apiDMARC runs through the report looking for domain names, makes an API call to retrieve DMARC record and pastes them into the relevant cell
//Written by Tass Kalfoglou - tasskalf@gmail.com

// API KEY: 

function apiDMARC() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var cell = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getActiveCell();
    
    var rowCount = 3;
    var lastRow = ss.getDataRange().getNumRows();
  
    while (rowCount <= lastRow) {
      //move to cell A3, store value of cell and check to see if cell contains the string 'Domain Group:'
      sheet.getRange(rowCount, 1).activateAsCurrentCell();
      var cellValue = ss.getCurrentCell().getValue();
      var dgTrue = cellValue.includes('Domain Group:');
      
      //if cell value equals blank or contains the string 'Domain Group:' ignore and add 1 to the counter
      if ( cellValue == '' || dgTrue == true ) {
        rowCount += 1;
      }
      
      //else perform the api call to get the DMARC record of the domain (cell value)
      else {
        var domain_name = '_dmarc.' + ss.getCurrentCell().getValue();
        //API parameters
        var options = {
          'method' : 'GET',
          'headers': {'X-Api-Key': 'API KEY'},
          'contentType': 'application/json',
        };
        //call API using urlFetch
        var apiCall = UrlFetchApp.fetch('https://api.api-ninjas.com/v1/dnslookup?domain=' + domain_name, options);  
        var response = apiCall.getContentText();
        console.log(response);
        //check if response contains the string 'v=DMARC1'
        var dmarcTrue = response.includes('v=DMARC1');
        
        //if DMARC record found
        if ( dmarcTrue == true ) {
            //check to see what DMARC policy is active
            var none = response.includes('p=none');
            var quarantine = response.includes('p=quarantine');
            var reject = response.includes('p=reject');
            
            //if p=none found
            if ( none == true ) {
                dmarcPolicy = 'p=none';
                sheet.getRange(rowCount, 3).activateAsCurrentCell().setValue(dmarcPolicy);
                sheet.getRange(rowCount, 1).activateAsCurrentCell();
                rowCount += 1;
            }
            
            //if p=reject found
            else if ( quarantine == true ) {
                dmarcPolicy = 'p=quarantine';
                sheet.getRange(rowCount, 3).activateAsCurrentCell().setValue(dmarcPolicy);
                sheet.getRange(rowCount, 1).activateAsCurrentCell();
                rowCount += 1;
            }
            
            //if p=reject found
            else if ( reject == true ) {
                dmarcPolicy = 'p=reject';
                sheet.getRange(rowCount, 3).activateAsCurrentCell().setValue(dmarcPolicy);
                sheet.getRange(rowCount, 1).activateAsCurrentCell();
                rowCount += 1;
            }
        }
        
        //if no DMARC record found
        else if (dmarcTrue == false ) {
          sheet.getRange(rowCount, 3).activateAsCurrentCell().setValue('No DMARC record');
          sheet.getRange(rowCount, 1).activateAsCurrentCell();
          rowCount += 1; 
        }
      }
    }
  }