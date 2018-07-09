function isValidRecordNumber(recordNumber) {
  if(isNaN(recordNumber) == true || isInteger(recordNumber) == false) {
    return false; 
  }
  else {
    if(recordNumber > 19840000 && recordNumber < 20500000) {
      return true; 
    }
  }
  return false;
}

function isValidSubawardNumber(subawardNumber) {
  var regExp1 = /^(KK|MC)[0-9]{4,5}$/g;
  var regExp2 = /^(KK|MC)[0-9]{4,5}-[0-9]{1,3}$/g;
  if(regExp1.exec(subawardNumber) || regExp2.exec(subawardNumber)) {
    return true;
  }
  return false;
}

//getTo() and getCc() can return emails surrounded in carats sometimes
//<mcnair@research.ucsb.edu> or mcnair@research.ucsb.edu
//need to remove the carats, they prevent the addresses from showing correctly in the html file
function getCleanEmailAddress(sender) {
  var regex = /\<([^\@]+\@[^\>]+)\>/g;
  var email = sender;  // Default to using the whole string.
  var cleanEmails = "";
  var match;
  do {
    match = regex.exec(sender);
    if (match) {
      if(cleanEmails == "") {
        cleanEmails = match[1];
      }
      else {
        cleanEmails += ", " + match[1];
      }
    }
  } while (match);
  
  if(cleanEmails == "") {
    return sender;
  }
  else {
    return cleanEmails;
  }
}

//removes re: cc: and fwd: from the subject line
function getCleanSubject(subject) {
  var newSubject = subject;
  var first3 = subject.toLowerCase().substring(0, 3);
  var first4 = subject.toLowerCase().substring(0, 4);
  if(first3 == "re:" || first3 == "cc:" || first3 == "fw:") {
    newSubject = newSubject.substring(3)
  }
  else if(first4 == "fwd:") { 
    newSubject = newSubject.substring(4);
  }
  
  //remove leading whitespace
  while(newSubject.substring(0, 1) == " ") {
    newSubject = newSubject.substring(1);
  }
  
  return newSubject.replace(/[\\\\/:*?\"<>|]/g, '_');
}

//removes bad characters from attachment file names
function getCleanAttachmentName(name) {
  return name.replace(/[\\\\/:*?\"<>|]/g, '_');
}

//extracts the record/subaward number from the subject line if there is one
function getRecordNumberFromString(value) {
  var regExp0 = /([2][0][0-9]{6}|[1][9][8-9][0-9]{5})/;
  var regExp1 = /((KK|MC)[0-9]{4,5})/;
  var regExp2 = /((KK|MC)[0-9]{4,5}-[0-9]{1,3})/;
  var match0 = regExp0.exec(value);
  var match1 = regExp1.exec(value);
  var match2 = regExp2.exec(value);
  
  if(match0) {
    return match0[1];
  }
  else if(match1) {
    return match1[1];
  }
  else if(match2) {
    return match2[1];
  }
  
  return "";
}

function isInteger(input) {
  var regExp = /^[0-9]*$/;  
  return regExp.test(input);
}

function isValidFileName(fileName) {
  var regExp1 = /^[^\\/:\*\?"<>\|]+$/; // forbidden characters \ / : * ? " < > |;
  var regExp2 = /^\./; // cannot start with dot (.)
  var regExp3 = /^(nul|prn|con|lpt[0-9]|com[0-9])(\.|$)/i; // forbidden file names

  return regExp1.test(fileName) && !regExp2.test(fileName) && !regExp3.test(fileName);
}
  
//returns an error notification in Gmail
function notifyError(message) {
  return CardService.newActionResponseBuilder()
     .setNotification(CardService.newNotification()
         .setType(CardService.NotificationType.ERROR)
         .setText(message))
     .build();
}

//returns an informational notification in Gmail
function notifyInfo(message) {
  return CardService.newActionResponseBuilder()
     .setNotification(CardService.newNotification()
         .setType(CardService.NotificationType.INFO)
         .setText(message))
     .build();
}

//returns {Date} - Email - {filename}.pdf for the email correspondence
//returns {Date} - {filename} for attachments
//Not all attachments are pdfs, so their file type is included in the filename parameter
function getRenamedFile(filename, msgDate, isAttachment) {
  var dd = msgDate.getDate();
  var mm = msgDate.getMonth() + 1;
  var yyyy = msgDate.getFullYear();
  
  if(dd < 10) {
      dd = '0'+ dd
  } 
  
  if(mm < 10) {
      mm = '0'+ mm
  } 
  msgDate = yyyy + "." + mm + "." + dd;
  
  if(isAttachment) {
    return msgDate + " " + filename;
  }
  else {
    return msgDate + " Email - " + filename + ".pdf";
  }
}

//Returns the number of folders that match the target folder name and one of their IDs
//This is currently only used to identify the root folder. The count should always be exactly 1
function getFolder(folderName) {
  var folders = DriveApp.getFoldersByName(folderName);
  var count = 0;
  var folderId = "";
  while(folders.hasNext()) {
    count++;
    folderId = folders.next().getId();
  }
  
  return new FolderInfo(count, folderId);
}

//returns the number of subfolders that match the target folder name, and one of their IDs
//currently unused
//function getSubfolder(rootFolderId, targetFolderName) {
//  var rootFolder = DriveApp.getFolderById(rootFolderId);
//  var targetFolders = rootFolder.getFoldersByName(targetFolderName);
//  var targetFolderCount = 0;
//  var targetFolderId = "";
//  while(targetFolders.hasNext()) {
//    targetFolderCount++;
//    targetFolderId = targetFolders.next().getId();
//  }
//  
//  return new FolderInfo(targetFolderCount, targetFolderId);
//}

function FolderInfo(count, id) {
  this.count = count;
  this.id = id;
}
