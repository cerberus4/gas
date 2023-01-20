//apiSPF runs through the report looking for domain names, makes an API call to retrieve SPF record and pastes them into the relevant cell
//Written by Tass Kalfoglou - tasskalf@gmail.com

// API KEY: 

function apiSPF() {
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

    //else perform the api call to get the SPF record of the domain (cell value)
    else {
      var domain_name = ss.getCurrentCell().getValue();
      //API parameters
      var options = {
        'method' : 'GET',
        'headers': {'X-Api-Key': 'API KEY'},
        'contentType': 'application/json',
      };
      //call API using urlFetch
      var apiCall = UrlFetchApp.fetch('https://api.api-ninjas.com/v1/dnslookup?domain=' + domain_name, options);  
      var response = apiCall.getContentText();
      //remove unnecessary characters
      response = response.replace(/}/g,'');
      response = response.replace(/{/g,'');
      response = response.replace(/"/g,'');
      response = response.replace(/]/g,'');
      //split string by ','
      var splitString = response.split(',')
      var i = 0;
      var foundSPF = 0;
      //cycle through split string checking for the string 'value: v=spf1'
      while (i < splitString.length) {
        var searchSPF = splitString[i].includes('value: v=spf1');
        //if string has been found, paste results, and set foundSPF variable to 1
        if ( searchSPF == true ) {
          spfValue = splitString[i].replace(' value: ','');
          sheet.getRange(rowCount, 4).activateAsCurrentCell().setValue(spfValue);
          sheet.getRange(rowCount, 1).activateAsCurrentCell(); 
          foundSPF = 1;
        } 
        i++;
      }
      //if SPF record has not been found paste 'No SPF Record' into cell
      if ( foundSPF == 0 ) {
        sheet.getRange(rowCount, 4).activateAsCurrentCell().setValue('No SPF record');
        sheet.getRange(rowCount, 1).activateAsCurrentCell(); 
      }
      rowCount += 1;
    }
  }
}