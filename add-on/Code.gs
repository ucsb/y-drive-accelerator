//This is the top level folder in the Team Drive that record folders will be created under
var _ROOTFOLDER = "ORBiT EFiles Folder";
//var _ROOTFOLDER = "Test Folder";
var _FLAGFILENAME = "_addOnFinished.txt";
var _TEMPFILENAME = "_addOnTemp.html";

//This function is called on startup
function buildAddOn(e) {
  // Activate temporary Gmail add-on scopes.
  var accessToken = e.messageMetadata.accessToken;
  GmailApp.setCurrentMessageAccessToken(accessToken);

  var msgId = e.messageMetadata.messageId;
  
  //return buildDisabledCard();
  return buildMainCard(msgId);
}

//Builds the initial card that is displayed to the user
function buildMainCard(msgId) {
  var msg = GmailApp.getMessageById(msgId);
  var subject = getCleanSubject(msg.getSubject());
  var recordNumber = getRecordNumberFromString(subject);  //pulls out the record number from the subject if there is one

  var card = CardService.newCardBuilder();
  card.setHeader(CardService.newCardHeader()
                 .setTitle("Save Email and Attachments"));
  
  var section = CardService.newCardSection();
  section.addWidget(CardService.newTextInput()
                    .setTitle("Record/Subaward Number:")
                    .setFieldName("txtRecordNumber")
                    .setValue(recordNumber));
  section.addWidget(CardService.newTextInput()
                    .setTitle("File Name:")
                    .setFieldName("txtFileName")
                    .setValue(subject));
  section.addWidget(CardService.newKeyValue()
                    .setTopLabel("Selected Email:")
                    .setMultiline(true)
                    .setContent("From " + msg.getFrom() + "on " + msg.getDate().toLocaleString("en-us")));
  
  section.addWidget(buildMessagesWidget("startSelected"));
  section.addWidget(buildAttachmentsWidget("startSelected", "selected"));
  
  var saveAction = CardService.newAction()
    .setFunctionName("saveCallback");
  var button = CardService.newTextButton()
    .setText("Save")
    .setOnClickAction(saveAction);
  section.addWidget(CardService.newButtonSet().addButton(button));
  
  card.addSection(section);
    
  return card.build();
}

//builds a new card with previously set values
function rebuildCard(msgId, recordNumber, filename, selectMessages, attachments) {
  var msg = GmailApp.getMessageById(msgId);
  if(recordNumber == null) {
    recordNumber = "";
  }
  if(filename == null) {
    filename = "";
  }
  var card = CardService.newCardBuilder();
  card.setHeader(CardService.newCardHeader()
                 .setTitle("Save Email and Attachments"));
                 
  var section = CardService.newCardSection();
  section.addWidget(CardService.newTextInput()
                    .setTitle("Record/Subaward Number:")
                    .setFieldName("txtRecordNumber")
                    .setValue(recordNumber.toString()));
  section.addWidget(CardService.newTextInput()
                    .setTitle("File Name:")
                    .setFieldName("txtFileName")
                    .setValue(filename.toString()));
  section.addWidget(CardService.newKeyValue()
                    .setTopLabel("Selected Email:")
                    .setMultiline(true)
                    .setContent("From " + msg.getFrom() + "on " + msg.getDate().toLocaleString("en-us")));
                    
  section.addWidget(buildMessagesWidget(selectMessages));
  section.addWidget(buildAttachmentsWidget(selectMessages, attachments));
  
  var saveAction = CardService.newAction()
    .setFunctionName("saveCallback");
  var button = CardService.newTextButton()
    .setText("Save")
    .setOnClickAction(saveAction);
  section.addWidget(CardService.newButtonSet().addButton(button));
  
  card.addSection(section);
    
  return card.build();
}

//creates the radio buttons for which messages to include
function buildMessagesWidget(selectMessages) {
  var item1Text = "Save thread from selected email";
  var item2Text = "Save thread from most recent email";
  var item3Text = "Only save selected email";
  var item1Value = "startSelected";
  var item2Value = "startMostRecent";
  var item3Value = "onlySelected";
  var item1Selected = false;
  var item2Selected = false;
  var item3Selected = false;
  
  var widget = CardService.newSelectionInput()
              .setType(CardService.SelectionInputType.RADIO_BUTTON)
              .setFieldName("rdoSelectMessages")
              .setOnChangeAction(CardService.newAction().setFunctionName("changeMessageType"));
  
  if(selectMessages == "startSelected") {
    item1Selected = true;
  }
  else if(selectMessages == "startMostRecent") {
    item2Selected = true;
  }
  else {
    item3Selected = true;
  }
  
  widget.addItem(item1Text, item1Value, item1Selected);
  widget.addItem(item2Text, item2Value, item2Selected);
  widget.addItem(item3Text, item3Value, item3Selected);
  
  return widget;
}

//creates the radio buttons for which attachments to include
function buildAttachmentsWidget(selectMessages, attachments) {
  var item1Text = "";
  var item2Text = "";
  var item3Text = "";
  var item1Value = "";
  var item2Value = "";
  var item3Value = "";
  var item1Selected = true;
  var item2Selected = false;
  var item3Selected = false;
  
  var widget = CardService.newSelectionInput()
               .setType(CardService.SelectionInputType.RADIO_BUTTON)
               .setFieldName("rdoAttachments");
  
  if(selectMessages == "startSelected") {
    item1Text = "Selected email's attachments only";
    item1Value = "selected";
    item2Text = "Include all attachments";
    item2Value = "all";
    item3Text = "Do not include attachments";
    item3Value = "none";
    if(attachments == "none") {
      item1Selected = false;
      item3Selected = true;
    }
  }
  else if(selectMessages == "startMostRecent") {
    item1Text = "Include all attachments";
    item1Value = "all";
    item2Text = "Do not include attachments";
    item2Value = "none";
    if(attachments == "none") {
      item1Selected = false;
      item2Selected = true;
    }
  }
  else {
    item1Text = "Include selected email's attachments";
    item1Value = "selected";
    item2Text = "Do not include attachments";
    item2Value = "none";
    if(attachments == "none") {
      item1Selected = false;
      item2Selected = true;
    }
  }
  
  widget.addItem(item1Text, item1Value, item1Selected);
  widget.addItem(item2Text, item2Value, item2Selected);
  if(item3Text != "") {
    widget.addItem(item3Text, item3Value, item3Selected)
  }
  
  return widget;
}

//this is used in version 15
function buildDisabledCard() {
  var card = CardService.newCardBuilder();
  card.setHeader(CardService.newCardHeader()
                 .setTitle("We have temporarily disabled access to the add-on. Please stand by."));
    
  return card.build();
}

//Called when the save button is clicked
function saveCallback(e) {
  var inputs = e.formInputs;
  var recordNumber = inputs["txtRecordNumber"];
  var fileName = inputs["txtFileName"];
  var selectMessages = inputs["rdoSelectMessages"];
  var attachments = inputs["rdoAttachments"];
  var msgId = e.messageMetadata.messageId;
  
  if(isValidFileName(fileName) == false) {
    return notifyError("Invalid File Name. The following characters are not allowed: \ / : * ? \" < > |"); 
  }
  
  if(isValidRecordNumber(recordNumber) == false && isValidSubawardNumber(recordNumber) == false) {
    return notifyError("Invalid Record/Subaward Number"); 
  }
  
  return saveToDrive(recordNumber, fileName, msgId, selectMessages, attachments);
}

//Creates a folder in the Team Drive
//Folder name is the record number
//Creates a temporary html file that is converted to a pdf
//The temporary html file is then trashed
//If the user chose to include attachments, they are saved as well
function saveToDrive(recordNumber, fileName, msgId, selectMessages, includeAttachments) {
  var rootFolderInfo = getFolder(_ROOTFOLDER);
  if(rootFolderInfo.count == 0){
    return notifyError("Root folder not found. Please make sure you can access the Team Drive");
  }
  else if(rootFolderInfo.count == 1) {
    var rootFolderId = rootFolderInfo.id;
    var targetFolderId = DriveApp.getFolderById(rootFolderId).createFolder(recordNumber).getId();
    var targetFolder = DriveApp.getFolderById(targetFolderId);
    var threadId = GmailApp.getMessageById(msgId).getThread().getId();
    var thread = GmailApp.getThreadById(threadId);
    var msgs = thread.getMessages();
    var html = "";
    
    if(selectMessages == "startSelected") {
      //save all messages in the thread previous to and including the selected message
      var reachedSelectedMsg = false;
      for (var m = msgs.length - 1; m >= 0; m--) {
        var msg = msgs[m];
        var msgDate = msg.getDate();  //the date of the current message in the loop - prepended on attachments
        if(msg.getId() == msgId) {
          reachedSelectedMsg = true;
          var msgDateForFile = msgDate;  //the date of the selected message - prepended on the email pdf
        }
        
        if(reachedSelectedMsg == true) {
          html += getMessageHeader(msg);
          
          var attachments = msg.getAttachments();
          if(includeAttachments != "none" && attachments.length > 0) {
            if(includeAttachments == "all" || (includeAttachments == "selected" && msg.getId() == msgId)) {
              var footer = "<strong>Attachments:</strong><ul>";
              for (var a = 0; a < attachments.length; a++) {
                var cleanedAttachment = getCleanAttachmentName(attachments[a].getName());  //remove bad characters from attachment name
                var renamedAttachment = getRenamedFile(cleanedAttachment, msgDate, true);  //prepend date
                targetFolder.createFile(attachments[a]).setName(renamedAttachment);
                footer += "<li>" + attachments[a].getName() + "</li>";
              }
              html += footer + "</ul>";
            }
          }
          html += "<br />";
        }
      }
    }
    else if(selectMessages == "startMostRecent") {
      //save all messages in the thread - most recent message first
      for (var m = msgs.length - 1; m >= 0; m--) {
        var msg = msgs[m];
        var msgDate = msg.getDate();  //the date of the current message in the loop - prepended on attachments
        if(m == (msgs.length - 1)) {
          var msgDateForFile = msgDate;  //the date of the most recent email in the thread - prepended on the email pdf
        }
        html += getMessageHeader(msg);
        
        var attachments = msg.getAttachments();
        if(includeAttachments == "all" && attachments.length > 0) {
          var footer = "<strong>Attachments:</strong><ul>";
          for (var a = 0; a < attachments.length; a++) {
            var cleanedAttachment = getCleanAttachmentName(attachments[a].getName());  //remove bad characters from attachment name
            var renamedAttachment = getRenamedFile(cleanedAttachment, msgDate, true);  //prepend date
            targetFolder.createFile(attachments[a]).setName(renamedAttachment);
            footer += "<li>" + attachments[a].getName() + "</li>";
          }
          html += footer + "</ul>";
        }
        html += "<br />";
      }
    }
    else {
      //only the selected email
      var msg = GmailApp.getMessageById(msgId);
      var msgDate = msg.getDate();
      var msgDateForFile = msgDate;
      html += getMessageHeader(msg);
      
      var attachments = msg.getAttachments();
      if(includeAttachments != "none" && attachments.length > 0) {
        var footer = "<strong>Attachments:</strong><ul>";
        for (var a = 0; a < attachments.length; a++) {
          var cleanedAttachment = getCleanAttachmentName(attachments[a].getName());  //remove bad characters from attachment name
          var renamedAttachment = getRenamedFile(cleanedAttachment, msgDate, true);  //prepend date
          targetFolder.createFile(attachments[a]).setName(renamedAttachment);
          footer += "<li>" + attachments[a].getName() + "</li>";
        }
        html += footer + "</ul>";
      }
      html += "<br />";
    }
      
    var tempFile = targetFolder.createFile(_TEMPFILENAME, html, "text/html");
    var renamedFile = getRenamedFile(fileName, msgDateForFile, false);
    targetFolder.createFile(tempFile.getAs("application/pdf")).setName(renamedFile);
    targetFolder.createFile(_FLAGFILENAME, Session.getActiveUser().getEmail());
  }
  else {
    return notifyError("Duplicate root folder. There must be only one folder in the Google Drive named '" + _ROOTFOLDER + "'");
  }
  
  return notifyInfo("Thread Saved"); 
}

//returns a string with the message header: from, to, etc.
function getMessageHeader(msg) {
  var html = "From: " + getCleanEmailAddress(msg.getFrom()) + "<br />";
  html += "To: " + getCleanEmailAddress(msg.getTo()) + "<br />";
  html += "Cc: " + getCleanEmailAddress(msg.getCc()) + "<br />";
  html += "Date: " + msg.getDate() + "<br />";
  html += "Subject: " + msg.getSubject() + "<br /><hr />"; 
  html += msg.getBody();
  html += "<hr />";
  
  return html;
}

//returns a new card with updated input options
function changeMessageType(e) {
  var inputs = e.formInputs;
  var recordNumber = inputs["txtRecordNumber"];
  var fileName = inputs["txtFileName"];
  var selectMessages = inputs["rdoSelectMessages"];
  var attachments = inputs["rdoAttachments"];
  var msgId = e.messageMetadata.messageId;
  
  return rebuildCard(msgId, recordNumber, fileName, selectMessages, attachments);
}






















