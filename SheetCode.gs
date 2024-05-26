var LIVE_URL = "TODO"

function sheet1_edit_C8_on_click() {
  editValue("C8")
}
function sheet1_edit_C12_on_click() {
  editValue("C12")
}
function sheet1_edit_C16_on_click() {
  editValue("C16")
}


function sheet2_edit_on_click() {
  // Check if we're allowed to edit the cell
  // In this sample, only empty cells can be edited
  var myRange = SpreadsheetApp.getActiveSheet().getCurrentCell()
  if( myRange.getValue() ) {
    SpreadsheetApp.getUi().alert("Please choose an empty cell to edit")
    return
  }
  // OK, we're allowed to Edit, unlock the cell for the current user
  var a1 = SpreadsheetApp.getActiveSheet().getCurrentCell().getA1Notation()
  var retVal = execCommand("start-cell-edit",
   {
    spreadsheetId: SpreadsheetApp.getActive().getId(),
    sheetName: SpreadsheetApp.getActiveSheet().getName(),
    a1Notation: a1,
    emailAddress: Session.getActiveUser().getEmail(),
    editTime: 2,
    verbose: true
  })
  myRange.setBackgroundRGB(201,218,248)
  myRange.activateAsCurrentCell()
  if( !retVal.ok) SpreadsheetApp.getUi().alert(retVal.errorMessage)
}



function editValue(a1Notation) {
  var me = Session.getActiveUser().getEmail()
  var ui = SpreadsheetApp.getUi()
  // Prompt for value, then set the value
  var newValue = ui.prompt("Enter a new value for " + a1Notation,ui.ButtonSet.OK_CANCEL).getResponseText()
  if(newValue) {
    var parameter = {
      spreadsheetId: SpreadsheetApp.getActive().getId(),
      sheetName: SpreadsheetApp.getActiveSheet().getName(),
      a1Notation: a1Notation,
      emailAddress: me,
      text: newValue,
      verbose: true
    }
    var retVal = execCommand("set-cell-text", parameter) 
    if(!retVal.ok) {
      SpreadsheetApp.getUi().alert("ERROR: " + retVal.errorMessage)
      return
    }
  }
  // Update audit log
  var auditRange = SpreadsheetApp.getActiveSheet().getRange(4,7,14,3)
  var myArray = auditRange.getValues()
  var newRow = [(new Date()).toLocaleString(),getMaskedEmail(me), a1Notation + ' edited to ' + newValue]
  myArray.splice(myArray.length-1,1)
  myArray.splice(0,0,newRow)
  var parameter = {
    spreadsheetId: SpreadsheetApp.getActive().getId(),
    sheetName: SpreadsheetApp.getActiveSheet().getName(),
    a1Notation: "G4",
    emailAddress: me,
    text: JSON.stringify(myArray),
    verbose: true
  }
  var retVal = execCommand("set-range-text", parameter)
  if(!retVal.ok) {
    SpreadsheetApp.getUi().alert("ERROR: " + retVal.errorMessage)
  }
}

 /**
 * Execute command
 * @param {string} myCommand LockUnlock command to execute
 * @param {Object} parameter Object
 * @param {string} parameter.spreadsheetId
 * @param {string} parameter.sheetName
 * @param {string} parameter.a1Notation cell address, eg 'A1'
 * @param {string} parameter.emailAddress
 * @param {number} parameter.editTime Length in minutes cell remains unlocked
 * @param {boolean} parameter.verbose
 * @return Success flag {ok: true/false, message: error message if ok is false}
 * @customfunction
 */
function execCommand(myCommand, parameter) {
  simulationMode = "LIVE"
  var myResponse
  var retVal = {ok:false}

  switch(simulationMode) {
    case "SIMULATE":
      parameter.simulate = true
    case "LIVE":
      var myURL = LIVE_URL + "/" + myCommand
      var params = {
        headers: { Authorization: 'Bearer ' + getToken()},
        muteHttpExceptions: true,
        method: 'post',
        payload: parameter
      }
      myResponse = UrlFetchApp.fetch(myURL, params)
      try {
        retVal = JSON.parse(myResponse.getContentText()) 
      } catch(e) {
        retVal = {ok:false, errorMessage: myResponse.getContentText()}
      }
      break
    case "LOCAL":
      var e = {
        pathInfo: myCommand,
        contentLength: 999.0,
        parameter: parameter,
        postData: {
          type: 'application/x-www-form-urlencoded',
          name: 'postData'
        }
      }
      myResponse = doPost(e)
      try {
        retVal = JSON.parse(myResponse.getContentText()) 
      } catch(e) {
        retVal = {ok:false, errorMessage: myResponse.getContentText()}
      }
      break
  }
  return retVal
}
function getMaskedEmail(emailAddress) {
  var retVal = ""
  try {
    var meSplit = emailAddress.split("@")
    var maskedName = meSplit[0][0] + "*".repeat(meSplit[0].length-1)
    var meSplit = meSplit[1].split(".")
    var maskedDomain = meSplit[0][0] + "*".repeat(meSplit[0].length-1)
    retVal = maskedName + "@" + maskedDomain + emailAddress.substring(maskedName.length + maskedDomain.length+1)
  } catch(e) {
    retVal = "Unknown email"
  }
  return retVal
}
function getToken() {
  return ScriptApp.getOAuthToken()
}

