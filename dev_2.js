var filesSheet1 = SpreadsheetApp.getActive().getSheetByName('Statefund Holders');
var range1 = filesSheet1.getDataRange();
var filesData1 = range1.getValues();  
var sec = 5;

var filesSheet3 = SpreadsheetApp.getActive().getSheetByName('Temp Sheet');
var range3 = filesSheet3.getDataRange();
var filesData3 = range3.getValues(); 
var arr = []

// Main function that runs after user uploads file

async function onEdit(e) { 
  Logger.log("Start");
  // try{
  //     var lock = LockService.getScriptLock();
  //     lock.waitLock(30000);    
  // }
  // catch(e){
  //     console.log(e)
  // } 
  try{
    console.log("inside lock")
    var lock = LockService.getDocumentLock();
    lock.waitLock(30000);
  }
  catch(e){
      console.log(e)
  }

  var count = 1
  for (let row = 1; row < filesData3.length; row++)
  { 
    if(filesData3[row][1] ) // check if the temp sheets has any uploaded files
      { 
        Logger.log(filesData3[row][1]) 
        //Utilities.sleep(3*1000);
        var fileInfo = filesData3[row][1];
        var zip_format =fileInfo.split("zip");
        var folderName = await getUniqueID(filesData3, row, filesData1)
        //var parentFolderId = await  getCurrentFolderID(filesData3, row);
        var parentFolderId = "1a5SqwhGFAXIs07o9xqFasQri3ZhSg75B"
        Logger.log(parentFolderId)  
        try{  
                var [newFolder,folderStatus] = await createFolder(parentFolderId, folderName);
                // if (folderStatus == "new"){
                //   arr.push(folderName)
                // }      
                if(zip_format[0] == fileInfo){
                   await filesHelper(newFolder, filesData1, folderName, filesData3, filesSheet1, filesSheet3, row, folderStatus)       
                }
                else{ 
                   await zipFileHelper(newFolder, filesData1, folderName, filesData3, filesSheet1, filesSheet3, row, folderStatus)    
                }
          }
          catch(e){
              console.log("In cache")
              console.log(e)
             
          }  
        
        Logger.log("count "+count)
        // if(count == 20){break}             
      }
      count = count + 1
    }
    console.log(arr) 
    // await delRow(filesSheet3, count,arr);
    await clearrow(filesSheet3, count,arr);
    console.log("logged out") 
    try{
      SpreadsheetApp.flush(); // applies all pending spreadsheet changes
      lock.releaseLock(); 
    }
    catch(e){
      console.log(e)
    }
  };






// function that handles non-zip files
async function filesHelper(newFolder, filesData1, folderName, filesData3, filesSheet1, filesSheet3, row, folderStatus){
    var fileInfo = filesData3[row][1];
    var uniqueKey = filesData3[row][0]
    var newFolderID = newFolder.getId();
    if( folderStatus =="new"){
      await updateSheet1WithLink(newFolder, filesData1, folderName, filesData3, filesSheet1, uniqueKey)
    }
    //await delRow(filesSheet3, row);
    await moveFiles(newFolderID, row, filesData3, folderName)  
}




// function that handles zip files
async function zipFileHelper(newFolder, filesData1, folderName, filesData3, filesSheet1, filesSheet3, row, folderStatus){
    var fileInfo = filesData3[row][1];
    var uniqueKey = filesData3[row][0]
    var newFolderID = newFolder.getId();
    console.log(newFolderID,"newFolderID")
    try{  
         
         
          await unzipFiles(fileInfo, newFolderID);
          if( folderStatus =="new"){await updateSheet1WithLink(newFolder, filesData1, folderName, filesData3, filesSheet1, uniqueKey)}
          //await delRow(filesSheet3, row);
          await moveFiles(newFolderID, row, filesData3,folderName)
            
        }
    catch(e){
          console.log(e)
          if( folderStatus =="new"){await updateSheet1WithLink(newFolder, filesData1, folderName, filesData3, filesSheet1, uniqueKey)}
          //await delRow(filesSheet3, row);
          await moveFiles(newFolderID, row, filesData3,folderName)  
        } 
}


async function checkFilesExit(filename){
  var results;
  var haBDs  = DriveApp.getFilesByName(filename)
  //Does not exist
  if(!haBDs.hasNext()){
  results =  haBDs.hasNext();
  }
  //Does exist
  else{
  results =  haBDs.hasNext();
  }
  Logger.log(results)
  return results;
}


//sends the emails
async function sendEmails(filesData, row, fileURL) {
  Logger.log("In email")
  var name = filesData[row][4];
  var policyNumber = filesData[row][3];
  var folderLink = filesData[row][9]
  var email = filesData[row][1];
  var key = filesData[row][0];
  var message = filesData[row][6];
  var holderEmail = filesData[row][5];
  var policyHolderName = filesData[row][2];
  var body;
  if(message){ var messageBody = "<br><br>Message From PolicyHolder: <b>"+ message +"</b>"}
  else{var messageBody = ""}
  body = "Hello " + email + ",<p>"+  name + " at "+ holderEmail + " has uploaded documents for the policy named above.  See details and link to file upload link below: <p>Confirmation Code: <b>" + key + "</b><p>Policy Number: <b>"+ policyNumber + "</b><p>Link for uploaded documents: " + fileURL + messageBody + "<br><br><br><p style= 'font-family:georgia,garamond,serif;font-size:14px;font-style:italic;'>*** This is a system-generated confirmation e-mail , please do not reply to this message.</p>"
  var subject =  "Records-Upload-Alert: " + policyHolderName + "-" +policyNumber; 	
  var ccList = 	"tarun.kappala@springml.com"
  await MailApp.sendEmail(email , subject, body,{
  htmlBody: body
  //cc:   ccList
});
console.log(subject);
var holderSubject = "State Fund Premium Audit Files Uploaded Successfully"
var holderBody = "Hello "+ name + ",<p>Your documents for the State Fund payroll audit have been uploaded successfully.  We will let you know if any further information is needed. Your confirmation ID is listed below.<p><b>Confirmation ID: " + key + "</b><br><br><br><p style= 'font-family:georgia,garamond,serif;font-size:14px;font-style:italic;'>*** This is a system-generated confirmation e-mail , please do not reply to this message.</p>"
await MailApp.sendEmail(holderEmail, holderSubject, holderBody,{
  htmlBody: holderBody
  //cc: ccList 
});
}


// Gets the unique ID from the sheet
async function getUniqueID(filesData3,row, filesData){
  Logger.log("In State Holders Sheet..getting unique Id")
  var uniqueKey = filesData3[row][0]

  var uniqueID;
  for (let row = 1; row < filesData.length; row++)
  { 
      if(filesData[row][0] == uniqueKey)
    {
        uniqueID = filesData[row][0];
        break
    }
  }
  Logger.log(uniqueID)
  return uniqueID;
}

// Get the time in PDT format
async function formatedDate() {
  var moment = new Date();
  Logger.log(moment)
  if(moment instanceof Date && !isNaN(moment)){
    var updatedTime = Utilities.formatDate(moment, "PST","YYYY-MM-dd hh:mm a z")
    Logger.log(updatedTime);
    return updatedTime
  } else {
    throw 'datetimeString can not be parsed as a JavaScript Date object'
  }
}


// unzips the files
async function unzipFiles(fileInfo, newFolderID){
  console.log("unzipp inside")
  var file = fileInfo.split("/");
  var files = DriveApp.getFilesByName(file[1]);
  var folderName = files.next();
  var count = 0
  var fileBlob = folderName.getBlob();
  var folder = DriveApp.getFolderById(newFolderID);
  fileBlob.setContentType("application/zip");
  var unZippedfile = Utilities.unzip(fileBlob); 
  console.log(unZippedfile.length)
  for(var i = 0; i < unZippedfile.length; i = i+1){
   // if(checkFilesExit(unZippedfile[i]) == true){
     // Logger.log("File Found")
      //}
      //else{
      //Logger.log("File Not Found")
      count = count + 1
      console.log(unZippedfile[i])
      newDriveFile = folder.createFile(unZippedfile[i]);
      //}  

  }   
  console.log(count)        
};




// update the statefund holders sheet with time and folder link
async function updateSheet1WithLink(folderID, filesData, folderName, filesData3, filesSheet, uniqueKey){
  Logger.log("In Sheet1")
  var fileInfo = filesData3[1][1];
  var file = fileInfo.split("/");
  for (let row = 0; row < filesData.length; row++)
  {
    if(row == 0)
    {
      continue;
    }
    else
    {
        if((!filesData[row][8])  && (filesData[row][0] == uniqueKey))
        {  var timestamp = filesData[row][7]
          var convertedTime  = await formatedDate()
           await filesSheet.getRange(row+1, 9).setValue(file[0] + "/" +folderName);
           var fileURL = ('=HYPERLINK("' + folderID.getUrl()  + '")');
           await filesSheet.getRange(row+1, 10).setValue(fileURL);
           if(!timestamp){filesSheet.getRange(row+1, 8).setValue(convertedTime);}
           await sendEmails(filesData,row, folderID.getUrl());
        }
        
    }
  }
};

//deletes the row in temp sheet after moving the file


async function delRow(sheet,count,arr){ 

  console.log(arr)
  var rows = sheet.getDataRange();
  var numRows = rows.getNumRows();
  var values = rows.getValues();
  console.log(numRows,"number of numRows")
  console.log(count,"count")
  
  // var rowsDeleted = 0;
  for (var i = count ; i >= 1 ; i--) {
    var row = values[i];
    if (arr.includes(row[1])) {
        console.log("row deleted" , row[1])
        sheet.deleteRow((parseInt(i)+1)); 
    }
  } 
};


async function clearrow(sheet,count,arr){
  var rows = sheet.getDataRange();
  var rem_rows = 15;
  for (var i = count ; i >= 1 ; i--) {
    range = sheet.getRange(i,1,1,3)
    var f_name = range.getValues()[0][1];
    if (arr.includes(f_name)) {
      console.log("row deleted",f_name)
      if (i > rem_rows){
        sheet.deleteRow((parseInt(i)));  
      }
      else{
        range.clearContent()
      }
    }
  } 
};


// async function delRow(sheet, count){ 
//   var lastRow = sheet.getLastRow();
//   var folderID = "1a5SqwhGFAXIs07o9xqFasQri3ZhSg75B";
//   var parentFolder = DriveApp.getFolderById(folderID);
//   var subFolders = parentFolder.getFolders();


//   if (count < 15) {
//     var status = lastRow 
//   }
//   else{
//     var status = count + 1;
//   }
//   Logger.log("count "+count)

//   for (var i = status; i > 1; i--) {
//         var folderName = filesData3[i-1][0]
//         if (checkFolder(subFolders,folderName) == "present"){
//           sheet.deleteRow(i);
//           Logger.log("row deleted "+ i )
//         }
        
//       }
// };



async function checkFolder(subFolders, folderName){
  Logger.log("Folder id", subFolders.next.getName())
  var doesntExists = true;
  // Check if folder already exists.
  while(subFolders.hasNext()){
    var folder = subFolders.next();
    
    //If the name exists return the id of the folder
    if(folder.getName() == folderName){
      doesntExists = false;
      return "present";
    };
  };
  //If the name doesn't exists, then create a new folder
  if(doesntExists == true){
    //If the file doesn't exists
    return "notpresent";
  };


};

// Get the Id of the current folder for the google sheet.
async function getCurrentFolderID(filesData, row){
  var fileInfo = filesData[row][1];
  var file = fileInfo.split("/");
  try{
  var files = DriveApp.getFilesByName(file[1]);
  
  while (files.hasNext())
  {
    var folderName = files.next();
    var folderId = folderName.getId();
    var parentFolderId = folderName.getParents().next().getId();
  }
  }
  catch(e){
    var parentFolderId = "1a5SqwhGFAXIs07o9xqFasQri3ZhSg75B"
  }
   return parentFolderId;
};


// Moves files from drive to the folder created based on Unique Key
async function moveFiles(targetFolderId, row , filesData, folderName) {
  var fileInfo = filesData[row][1];
  var unikeyId = filesData[row][0]
  var file = fileInfo.split("/");
  var files = DriveApp.getFilesByName(file[1]);  
  var processedFolder = DriveApp.getFolderById(targetFolderId);
  while (files.hasNext())
  {   
    var fileName = files.next();
    var renamedfile = fileName.setName(unikeyId + "_" + fileName.getName());
    renamedfile.moveTo(processedFolder);
    
  }
  arr.push(fileInfo)
};


var parentFolder = DriveApp.getFolderById("1a5SqwhGFAXIs07o9xqFasQri3ZhSg75B");
var subFolders = parentFolder.getFolders();


//Creates a folder based on Unique Key
// async function createFolder(folderID, folderName){
//   Logger.log("Folder id", folderID)
//   var parentFolder = DriveApp.getFolderById(folderID);
//   var subFolders = parentFolder.getFolders();
//   var doesntExists = true;
//   var newFolder = '';
//   var temp_folder = parentFolder.getFoldersByName(folderName)
//   console.log(temp_folder)
//   var num = 0
//   // Check if folder already exists.
//   console.log(folderName,"folderName")
  
//   while(subFolders.hasNext()){
    
//     var folder = subFolders.next();
//     num++
//     //If the name exists return the id of the folder
//     if(folder.getName() == folderName){
//       doesntExists = false;
//       newFolder = folder;
//       console.log(temp_folder.hasNext(),"hasnext")
//       console.log(num)
//       return [newFolder, "old"];
//     };
//   };
//   //If the name doesn't exists, then create a new folder
//   if(doesntExists == true){
//     //If the file doesn't exists
    
//     console.log(temp_folder.hasNext(),"hasnext1")
//     // console.log(temp_folder.next(),"next1")
//     // console.log(temp_folder.getContinuationToken(),"token1 ")    
//     newFolder = parentFolder.createFolder(folderName);
//     return [newFolder,"new"];
//   };


async function createFolder(folderID, folderName){
  Logger.log("Folder id", folderID)
  var parentFolder = DriveApp.getFolderById(folderID);
  var newFolder = '';
  var current_folder = parentFolder.getFoldersByName(folderName)
  console.log(current_folder)
   // Check if folder already exists.
  if(current_folder.hasNext()){
    console.log(current_folder.hasNext(),"hasnext")
    console.log("already exists")
    return [current_folder.next(), "old"];
  }
  else{
    console.log(current_folder.hasNext(),"hasnext1")
    console.log("creating because doesn't exist")
    newFolder = parentFolder.createFolder(folderName);
    return [newFolder,"new"];
  }

};



