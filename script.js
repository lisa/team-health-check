/*
* Written for SREP by Lisa Seelye <lseelye@redhat.com> in October 2019 with some extra bugfixes
*/

function onOpen() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet()
  var entries = [{
    name : "Refresh",
    functionName : "refreshLastUpdate"
  }]
  sheet.addMenu("Refresh Form Results", entries)
}

function refreshLastUpdate() {
  SpreadsheetApp.getActiveSpreadsheet().getRange('Z9').setValue(new Date().toTimeString());
}


/**
* Parse a result row
* @param {date} Parse the row for this date
* @param {string} Parse for this metric name (eg "Fun", or "Learning (Career Development)")
* @return The result for this metric
* @customfunction
*/
function parseRow(parseDate,ignored) {
  Logger.log("Running...")
  var formResponseSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses")
  var dataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Analysis")
  var index = 0 // index into form response sheet for columns with the data
  
  metricName = SpreadsheetApp.getActiveRange().offset(0, -1 * SpreadsheetApp.getActiveRange().getColumn() +1).getValue()
  
  // Find all the results in the resultSheet for the given date
  // indexes are in the order they appear in the result sheet, 0 based
  switch (metricName) {
    case "Teamwork":
      index = 2
      break
    case "Pawns or Players":
      index = 3
      break
    case "Health of Codebase":
      index = 4
      break
    case "Mission":
      index = 5
      break
    case "Learning (Career Development)":
      index = 6
      break
    case "Support (to do our job)":
      index = 7
      break
    case "Fun":
      index = 8
      break
  }

  if (index == 0) {
    return -1
  }
  
  var allDataForMetricName = formResponseSheet.getDataRange().getValues()
  
  var dataForDate = filterDataByMetricAndDate(allDataForMetricName, parseDate, index)
  
  
  var sum = 0
  var i = 0
  for (var d = 0; d < dataForDate.length; d++) {
    i = i + 1
    
    if (dataForDate[d] === 2) {
      // need to look at the team result of the previous entry, which is one to the left of the current cell
      sum = sum + parseFloat(SpreadsheetApp.getActiveRange().offset(0,-1).getValue())
    } else {
      sum = sum + dataForDate[d]
    }
  }

  return sum / i
  
}

/* Return an array of data filtered based on date and only for a specific index */
function filterDataByMetricAndDate(data,date,index) {
  result = []
  for (var r = 1; r < data.length; r++ ) {
    if (data[r][0].getFullYear() == date.getFullYear() && data[r][0].getMonth() == date.getMonth() && data[r][0].getDate() == date.getDate()) {
      result.push(data[r][index])
    }
  }
  return result;
}
