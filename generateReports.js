//Generate Reports runs through the report looking for groupings and pastes them into a new spreadsheet titles by the name of the group
//Written by Tass Kalfoglou - tasskalf@gmail.com

function updateReport() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var cell = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getActiveCell();
  
  var rowCount = 1;
  var lastRow = ss.getDataRange().getNumRows();
  var reportSheetName = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();

  while (rowCount <= lastRow) {
    //go to cell A1 and store value of cell   
    sheet.getRange(rowCount, 1).activateAsCurrentCell();  
    var cellValue = ss.getCurrentCell().getValue();                   
    //Determine if the cell value contains the word Domain Group:
    var dgTrue = cellValue.includes('Domain Group:');
    if (dgTrue == true) {  
      //declare sheetname, cell row to copy from, second row count for while loop and blank cell counter
      var sheetName = cellValue.replace('Domain Group: ','');         
      var fromRow = rowCount;                                         
      var rowCount2 = rowCount + 1;                                       
      var blankCell = 0;                                    
      //move down 1 cell, get value of cell and determine if string contains word Domain Group: (second counter)
      sheet.getRange(rowCount2, 1).activateAsCurrentCell();           
      var cellValue2 = ss.getCurrentCell().getValue();
      var dgTrue2 = cellValue2.includes('Domain Group:');
      //while second counter cell value does not equal 'Domain Group:', keep moving down the rows.
      while (dgTrue2 == false) {       
        rowCount2 += 1;                                               
        sheet.getRange(rowCount2, 1).activateAsCurrentCell();         
        cellValue2 = ss.getCurrentCell().getValue();
        dgTrue2 = cellValue2.includes('Domain Group:');
        //if cell is blank add a counter to blankCell variable
        if (cellValue2 == '') {
          blankCell += 1;
        }
        //if number of blank cells equals 15 this means we are at the end of the script, performs function to copy last group
        if (blankCell == 15) {
          //declare row that group ends
          console.log(rowCount2);
          var toRow = rowCount2 - 13;
          var numRows = toRow - fromRow + 1; 
          //create copy of template sheet, create a new sheet and give it the name of variable sheetName
          var blankSheet = ss.getSheetByName('blank');
          blankSheet.copyTo(ss).setName(sheetName);
          var sourceSheet = ss.getSheetByName(reportSheetName);
          var sourceRange = sourceSheet.getRange(fromRow, 1, numRows, 9);
          var targetSheet = ss.getSheetByName(sheetName);
          sourceRange.copyTo(targetSheet.getRange(3, 1));
          break;
        }
      }
      //if second counter cell value contains the string 'Domain Group:' we are at the end of the group. 
      if (dgTrue2 == true) {  
        //declare ranges to copy from
        var toRow = rowCount2 -1;
        var numRows = toRow - fromRow + 1; 
        //create copy of template sheet, create a new sheet and give it the name of variable sheetName
        var blankSheet = ss.getSheetByName('blank');                  
        blankSheet.copyTo(ss).setName(sheetName);        
        var sourceSheet = ss.getSheetByName(reportSheetName);
        var sourceRange = sourceSheet.getRange(fromRow, 1, numRows, 9);
        var targetSheet = ss.getSheetByName(sheetName);
        sourceRange.copyTo(targetSheet.getRange(3, 1));
      }

    }
    rowCount += 1;
  }
}