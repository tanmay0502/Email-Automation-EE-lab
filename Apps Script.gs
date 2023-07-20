var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Issued");
var ss2 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Returned");
var templateText = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Template").getRange(1,1).getValue();
var templateText2 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Template").getRange(1,2).getValue();
var templateText3 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Template").getRange(1,4).getValue();

function sendEmails(){
  var lr = ss.getLastRow();
  var lc = ss.getLastColumn();
  var templateText = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Template").getRange(1,1).getValue();
  for(var i = 2; i<=lr; i++){
    if(ss.getRange(i,10).getValue() === "Issued"){

      var currentEmail = ss.getRange(i,4).getValue();
      var currentName = ss.getRange(i,7).getValue();
      var currentComponents = ss.getRange(i,9).getValue()
      // Logger.log(currentEmail);
      var messageBody = templateText.replace("{name}", currentName).replace("{components}",currentComponents);


      Logger.log(messageBody);
      MailApp.sendEmail(currentEmail,
      "Reminder for Components issued from Electrical Engineering Lab",
      messageBody,{ noReply: true})
      Browser.msgBox("Email Sent to " + currentName + ".")
    }

    
  }

    
}

function sendCurrEmail(row,action){
      var currentEmail = ss.getRange(row,4).getValue();
      var currentName = ss.getRange(row,7).getValue();
      var currentComponents = ss.getRange(row,9).getValue();
      var currAction = action;
      if(currAction === "Partially"){
        currAction = "Returned"
        var returnedComponents = ss.getRange(row,12).getValue();
        var messageBody = templateText3.replace("{name}", currentName).replace("{components}",currentComponents).replace("{currAction}",currAction).replace("{returned}",returnedComponents);
        try{
        MailApp.sendEmail(currentEmail,
        "Electrical Engineering Lab Components status",
        messageBody,{ noReply: true})
        }catch(err){
          Browser.msgBox(err)
        }

        var quotaLeft = MailApp.getRemainingDailyQuota();
        Browser.msgBox("You have " + quotaLeft + " left as per daily gmail limit.")
      }else{

        // Browser.msgBox(currentEmail);
        var messageBody = templateText2.replace("{name}", currentName).replace("{components}",currentComponents).replace("{currAction}",currAction);
        // Browser.msgBox(currentEmail);
        try{
        MailApp.sendEmail(currentEmail,
        "Electrical Engineering Lab Components status",
        messageBody,{ noReply: true})
        Browser.msgBox("Email Sent to" + currentName + ".")
        }catch(err){
          Browser.msgBox(err)
        }

        
      }
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Email Actions")
      .addItem("Send Reminder Email", "sendEmails")
      .addToUi();
}

function newIssueEmail(){
  var lr = ss.getLastRow();
  var lc = ss.getLastColumn();
  var templateText = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Template").getRange(1,3).getValue();

      var currentEmail = ss.getRange(lr,4).getValue();
      var currentName = ss.getRange(lr,7).getValue();
      var currentComponents = ss.getRange(lr,9).getValue()
      // Logger.log(currentEmail);
      var messageBody = templateText.replace("{name}", currentName).replace("{components}",currentComponents);

      MailApp.sendEmail(currentEmail,
      "Components issued from Electrical Engineering Lab",
      messageBody,{ noReply: true})
      Browser.msgBox("Email Sent to" + currentName + ".")
    ss.getRange(lr,10).setValue("Issued")
    sortResponses() 

}

function onTheEdit(e) {
  const setValueCol = e.range.getColumn()
  const setValueRow = e.range.getRow()
  const action = ss.getRange(setValueRow, 10).getValue() 
  var lc = ss.getLastColumn();

  if(setValueCol === 10 ){
        row=[]
        for(cell = 1; cell<=lc; cell++){
          row.push(ss.getRange(setValueRow,cell).getValue())
        }
    if(action === "Issued" || action === "Long Term Issued"){
      ss.getRange(setValueRow, 11).setValue("")
    }else if(action === "Partially"){
      timezone = "GMT+5:30"
      var date = Utilities.formatDate(new Date(), "GMT+5:30", "yyyy-MM-dd HH:mm")
    }else{
        sendCurrEmail(setValueRow,action)
        timezone = "GMT+5:30"
        var date = Utilities.formatDate(new Date(), "GMT+5:30", "yyyy-MM-dd HH:mm")
        ss.getRange(setValueRow, 11).setValue(date)

        row=[]
        for(cell = 1; cell<=lc; cell++){
          row.push(ss.getRange(setValueRow,cell).getValue())
        }
        copyRow(row)
        ss.deleteRow(setValueRow)
        
    }

  }else if(setValueCol == 12){
        row=[]
        for(cell = 1; cell<=lc; cell++){
          row.push(ss.getRange(setValueRow,cell).getValue())
        }
        if(ss.getRange(setValueRow, 12).getValue() != ""){
          timezone = "GMT+5:30"
          var date = Utilities.formatDate(new Date(), "GMT+5:30", "yyyy-MM-dd HH:mm")
          ss.getRange(setValueRow, 11).setValue(date)
          sendCurrEmail(setValueRow,action)
      }
  }

}

function copyRow(row) {
  var lastRow = ss2.getLastRow();
  var lc = ss2.getLastColumn();
  for(cell = 1; cell<=lc; cell++){
    ss2.getRange(lastRow+1,cell).setValue(row[cell-1])
  }
  sortResponses()
}
function sortResponses() {
  ss.sort(1, false);
  ss2.sort(1, false);
}

function runDaily(){
  var lr = ss.getLastRow();
  var lc = ss.getLastColumn();
  var templateText = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Template").getRange(2,1).getValue();
  for(var i = 2; i<=lr; i++){
    var dateCell = ss.getRange(i,1)
    if(dateAction(dateCell)){

    if(ss.getRange(i,10).getValue() !== "Returned" && ss.getRange(i,10).getValue() !== "Consumed"){
      var currentEmail = ss.getRange(i,4).getValue();
      var currentName = ss.getRange(i,7).getValue();
      var currentComponents = ss.getRange(i,9).getValue()
      // Logger.log(currentEmail);
      var messageBody = templateText.replace("{name}", currentName).replace("{components}",currentComponents);


      Logger.log(messageBody);
      MailApp.sendEmail(currentEmail,
      "Reminder for EE lab Issued Components",
      messageBody,{ noReply: true})
      Browser.msgBox("Email Sent to" + currentName + ".")
    }
    
  }else{
    Logger.log("none")
  }
  }
}

function dateAction(actionDate) {
  var dateCell = actionDate
  var date = dateCell.getValue();
  var currentDate = new Date();
  var timeDiff = currentDate.getTime() - date.getTime();
  var diffInDays = Math.floor(timeDiff / (1000 * 3600 * 24));
  Logger.log(diffInDays)
  if (diffInDays % 30 === 0 && diffInDays >= 0) {
    return(true)
  } else {
    return(false)
  }
}

