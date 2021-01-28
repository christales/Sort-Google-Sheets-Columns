/**
* To be used with Google Spreadsheets with Google Apps Script
* Sorts the columns of a chosen sheet preserving any applied formatting or comments
*
* @param {string} sheetName - name of the selected sheet
* @param {number} startingColumn (optional, starting with column B by default) - indicates where the sorting should start. Col A = 1, Col B = 2,...
*
* by Chris Tales
*/

function sortSheet(sheetName, startingColumn) {

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);

  const initialHeaders = getCurrentHeaders(sheetName);
  startingColumn ? initialHeaders.splice(0, startingColumn-1) : initialHeaders.shift(); //remove IF your first column does not contain a descriptive first row
  initialHeaders.sort(); //sorting the headers alphabetically
  
  const uniqueHeaders = [...new Set(initialHeaders)]
  if (!uniqueHeaders[0]) {uniqueHeaders.shift()}

  const headersArray = [];
  for (const i in uniqueHeaders) {
    var itemAmt = initialHeaders.filter(x => x === uniqueHeaders[i]).length
    headersArray.push([uniqueHeaders[i], itemAmt])
  }

  let inspectedColumnArrIndex = startingColumn ? startingColumn -1 : 1; //array index: starting with col B by default
  let currentTargetIndex = startingColumn ? startingColumn : 2; //range index: starting with col B by default
  let searchedHeaderValue, count, currentHeaders, colToMove;
  
  for (const j in headersArray) {
    count = headersArray[j][1];
    while (count > 0) { //checking if there's any more instances of a particular header left to check against the existing sheet
      currentHeaders = getCurrentHeaders(sheetName) //updating the headers with each pass of a unique header
      searchedHeaderValue = headersArray[j][0]
      if (searchedHeaderValue === currentHeaders[inspectedColumnArrIndex]) {//checking if the searched header matches the currently inspected column
        currentSourceIndex = 1 + inspectedColumnArrIndex
        if (currentSourceIndex != currentTargetIndex){ //checking if the column needs to be moved
          colToMove = sheet.getRange(1, currentSourceIndex)
          sheet.moveColumns(colToMove, currentTargetIndex)
          currentSourceIndex++
        }
        count--; //removing the counter since the column was either moved or in the appropriate position
        currentTargetIndex++;
      }
      inspectedColumnArrIndex++; //moving to the next inspected column
    }
    inspectedColumnArrIndex = headersArray[j][1]; //resetting the columnArrayIndex to start after the already sorted columns
  }
}

function getCurrentHeaders(sheetName){
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  return sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0]
}
