var PRODUCT_VERSION = "001"
var PRODUCT_NAME = "Lock-Unlock-Sheet-v1"

/**
 * doPost is the WebApp command handler. Invoked by your spreadsheet code to perform a lock/unlock command.
 * This handler can be run in three modes:
 *  - Live.     Normal operation, deployed as a WebApp
 *  - Local.    If included as a file in the project, the doPost handler can be called directly (useful for testing/debugging)
 *  - Simulate. Deployed as a WebApp, but performs no action, except to return the paramater signature (useful for testing/debugging)
 * 
 * All parameters are set as fields in a parameter object
 * See: https://developers.google.com/apps-script/guides/web#request_parameters
 * 
 * doPost returns an object to the caller with "ok" flag. true = all good, false = something went wrong. 
 */
function doPost(e) {
  var retVal = {ok:false}
  // If simulate mode, return event object and exit...
  if( e.parameter.simulate) return ContentService.createTextOutput(JSON.stringify(e))
  // ... otherwise, process the command
  switch(e.pathInfo) {
    case "get-info":
      retVal = getInfo(e.parameter)
      break    
    case "set-cell-text":
      retVal = setCellText(e.parameter)
      break
    case "set-range-text":
      retVal = setRangeText(e.parameter)
      break
    case "lock-sheet":
      retVal = lockSheet(e.parameter)
      break
    case "unlock-sheet":
      retVal = unlockSheet(e.parameter)
      break
    case "start-cell-edit":
      retVal = startCellEdit(e.parameter)
      break
    case "end-cell-edit":
      retVal = endCellEdit(e.parameter)
      break
    case "generate-error":
      retVal = generateError(e.parameter)
      break
    default:
      retVal = {ok:false, errorMessage: PRODUCT_NAME + ":: Unknown command: " + e.pathInfo}
  }
  return ContentService.createTextOutput(JSON.stringify(retVal))
}
function doGet(e) {
  var versionMessage = PRODUCT_NAME + " " + PRODUCT_VERSION
  var retVal = ContentService.createTextOutput(JSON.stringify(versionMessage))
  return retVal
}
function getInfo(parameter) {
  return {
    ok: true,
    productName: PRODUCT_NAME,
    productVersion: PRODUCT_VERSION
  }
}
/**
 * Set value of a  cell
 * @param {Object} parameter
 * @param {string} parameter.spreadsheetId
 * @param {string} parameter.sheetName
 * @param {string} parameter.a1Notation
 * @param {string} parameter.text
 * @param {boolean} parameter.verbose
 * @return Success flag {ok: true/false, errorMessage: text if ok is false}
 * @customfunction
 */
function setCellText(parameter) {
  var retVal = {ok:false}
  if(parameter.verbose) logInstrumentation("setCellText", parameter)
  try {
    var ss = SpreadsheetApp.openById(parameter.spreadsheetId)
    var sht = ss.getSheetByName(parameter.sheetName)
    var rng = sht.getRange(parameter.a1Notation)
    var shtProtect = sht.getProtections(SpreadsheetApp.ProtectionType.SHEET)[0]
    var unProtList = shtProtect.getUnprotectedRanges()
    unProtList.push(rng)
    shtProtect.setUnprotectedRanges(unProtList)
    sht.getRange(parameter.a1Notation).setValue(parameter.text)
    unProtList.pop()
    shtProtect.setUnprotectedRanges(unProtList)
    retVal.ok = true
  } catch(e) {
    retVal.errorMessage = e.text
    console.log("setCellText:Error:: " + e.message + "\nparameter::" + JSON.stringify(parameter))
  }
  return retVal
}
/**
 * Set values for a range of cells
 * @param {Object} parameter
 * @param {string} parameter.spreadsheetId
 * @param {string} parameter.sheetName
 * @param {string} parameter.a1Notation range start
 * @param {string} parameter.text 2 dimension array of new values, stringified
 * @param {boolean} parameter.verbose
 * @return Success flag {ok: true/false, errorMessage: text if ok is false}
 * @customfunction
 */
function setRangeText(parameter) {
  var retVal = {ok:false}
  if(parameter.verbose) logInstrumentation("setRangeText", parameter)
  try {
    var ss = SpreadsheetApp.openById(parameter.spreadsheetId)
    var sht = ss.getSheetByName(parameter.sheetName)
    var myArray = JSON.parse(parameter.text)
    var beginRow = sht.getRange(parameter.a1Notation).getRow()
    var beginColumn = sht.getRange(parameter.a1Notation).getColumn()
    var rowCount = myArray.length
    var columnCount = myArray[0].length
    var rng = sht.getRange(beginRow,beginColumn,rowCount,columnCount)
    var shtProtect = sht.getProtections(SpreadsheetApp.ProtectionType.SHEET)[0]
    var unProtList = shtProtect.getUnprotectedRanges()
    unProtList.push(rng)
    rng.setValues(myArray)
    unProtList.pop()
    shtProtect.setUnprotectedRanges(unProtList)
    retVal.ok = true
  } catch(e) {
    retVal.errorMessage = e.text
    console.log("setRangeText:Error:: " + e.message + "\nparameter::" + JSON.stringify(parameter))
  }
  return retVal
}
/**
 * Remove any existing protection and lock sheet for editing
 * @param {Object} parameter Object
 * @param {string} parameter.spreadsheetId
 * @param {string} parameter.sheetName
 * @param {boolean} parameter.verbose
 * @return Success flag {ok: true/false, message: error message if ok is false}
 * @customfunction
 */
function lockSheet(parameter) {
  if(parameter.verbose) logInstrumentation("lockSheet", parameter)
  var ss = SpreadsheetApp.openById(parameter.spreadsheetId)
  var sht = ss.getSheetByName(parameter.sheetName)
  sht.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach( function(myProt) {
    myProt.remove()
  })
  sht.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach( function(myProt) {
    myProt.remove()
  })
  var prot = sht.protect()
  var me = Session.getEffectiveUser()
  prot.addEditor(me)
  prot.removeEditors(prot.getEditors());
  prot.setDomainEdit(false)
  return {ok:true}
}
/**
 * Remove any existing protection
 * @param {Object} parameter Object
 * @param {string} parameter.spreadsheetId
 * @param {string} parameter.sheetName
 * @param {boolean} parameter.verbose
 * @return Success flag {ok: true/false}
 * @customfunction
 */
function unlockSheet(parameter) {
  if(parameter.verbose) logInstrumentation("unlockSheet", parameter)
  var ss = SpreadsheetApp.openById(parameter.spreadsheetId)
  var sht = ss.getSheetByName(parameter.sheetName)
  sht.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach( function(myProt) {
    myProt.remove()
  })
  sht.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach( function(myProt) {
    myProt.remove()
  })
  return {ok:true}
}
/**
 * Unlock cell for a user to edit (also ending any other editing session)
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
function startCellEdit(parameter) {
  // end any current editing session
  var retVal = endCellEdit({
    spreadsheetId: parameter.spreadsheetId,
    sheetName: parameter.sheetName,
    emailAddress: parameter.emailAddress,
    verbose: parameter.verbose
  })
  if(!retVal.ok) return retVal
  if(typeof parameter.editTime === 'undefined') parameter.editTime = 5 
  var endTime = (new Date()).getTime() + (parameter.editTime * 60 * 1000)  
  var me =              Session.getEffectiveUser().getEmail()
  var ss =              SpreadsheetApp.openById(parameter.spreadsheetId)
  var sht =             ss.getSheetByName(parameter.sheetName)
  var rng =             sht.getRange(parameter.a1Notation)
  var shtProtect =      sht.getProtections(SpreadsheetApp.ProtectionType.SHEET)[0]
  var unProtectList =   shtProtect.getUnprotectedRanges()
  try {
    // exit if a cell is already in edit mode
    for(var i=0; i< unProtectList.length; i++) {
      if(unProtectList[i].getA1Notation() == rng.getA1Notation()) {
        retVal = {
          ok: false,
          errorMessage: "Cell is being edited by another person" 
        }
        return retVal
      }
    }
    // Add cell to unprotected list of sheet
    unProtectList.push(rng)
    shtProtect.setUnprotectedRanges(unProtectList)
    // Add a protecton to the range, and remove all editors except me + emailAddress
    var rngProtect = rng.protect()
    rngProtect.getEditors().forEach(function(email){
      if(email != me ) rngProtect.removeEditor(email)
    })    
    rngProtect.setDomainEdit(false)
    rngProtect.addEditor(parameter.emailAddress)
    rngProtect.setDescription(parameter.emailAddress + " " + endTime)
  } catch (e) {
    retVal = {
      ok: false,
      errorMessage: e.errorMessage
    }
  }
  return retVal
}
/**
 * Finish editing for a user, locking the cell. Also end any expired edit sessions
 * @param {Object} parameter Object
 * @param {string} parameter.spreadsheetId
 * @param {string} parameter.sheetName
 * @param {string} parameter.emailAddress
 * @param {boolean} parameter.verbose
 * @return Success flag {ok: true/false, message: error message if ok is false}
 * @customfunction
 */
function endCellEdit(parameter) {
  var retVal = {ok: true}
  if (parameter.verbose) logInstrumentation("endCellEdit",parameter)
  var me =              Session.getEffectiveUser().getEmail()
  var ss =              SpreadsheetApp.openById(parameter.spreadsheetId)
  var sht =             ss.getSheetByName(parameter.sheetName)
  var endTime =         (new Date()).getTime() 
  // Remove any ranges that user has rights to edit, or which have expired
  try {
    if(!sht.getProtections(SpreadsheetApp.ProtectionType.SHEET).length) {
      lockSheet({spreadsheetId: parameter.spreadsheetId, sheetName: parameter.sheetName})
    }
    var shtProtect = sht.getProtections(SpreadsheetApp.ProtectionType.SHEET)[0]
    var rngList = sht.getProtections(SpreadsheetApp.ProtectionType.RANGE)
    var unProtectList = shtProtect.getUnprotectedRanges()
    rngList.forEach(function(rng) {
      var myDescription = rng.getDescription()
      var emailAddress, timestamp
      [emailAddress, timestamp] = myDescription.split(" ")
      if(emailAddress == parameter.emailAddress || endTime > parseInt(timestamp)) {
        var r = rng.getRange()
        for(var i=0; i< unProtectList.length; i++) {
          if(unProtectList[i].getA1Notation() == rng.getRange().getA1Notation()) {
            unProtectList.splice(i,1)
            i--
          }
        }
        rng.remove()
      }
    })
    shtProtect.setUnprotectedRanges(unProtectList)
  } catch(e) {
    retVal = {
      ok: true,
      errorMessage: e.errorMessage
    }
  }
  return retVal
}
/**
 * Internal method to log parameters, when verbose flag is true
 */
function logInstrumentation(logReference, parameter) {
  var myString  = "*** " + logReference + " Instrument ***\n"
  myString +=   "activeUser:    " 
  myString += (Session.getActiveUser().getEmail()) ? Session.getActiveUser().getEmail() : "(none)"
  myString += "\neffectiveUser: " + Session.getEffectiveUser().getEmail() + "\n"
  console.log(myString)
  console.log("parameter:     " + JSON.stringify(parameter))
  if(parameter.spreadsheetId) {
    try {
      var x = SpreadsheetApp.openById(parameter.spreadsheetId)
      console.log("Spreadsheet name:  " + x.getName())
      console.log("Spreadsheet owner: " + x.getOwner())
    } catch(e) {
      console.log("Spreadsheet:       Couldn't open spreadsheet")
    }
  }
}
/**
 * Generate an error (for testing / debugging your error handler)
 * @param {Object} parameter Object
 * @param {string} parameter.text
 * @customfunction
 */
function generateError(parameter) {
  logInstrumentation("generateError", parameter)
  var errorType = parseInt(parameter.text)
  var retVal
  switch(errorType) {
    case 2: // generate runtime error
      SpreadsheetApp.openById("")
      break
    case 1: // generate error with helpful info message
      retVal = {ok:false, errorMessage: "This is an error message"}
      break
    case 0: // generate error with no info
    default:
      retVal = {ok:false}
  }
  return retVal
}

