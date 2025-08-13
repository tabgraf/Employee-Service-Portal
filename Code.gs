/**
 * ------------------------------------------------------
 * Project        : Employee Service Portal
 * Company        : Tabgraf Technologies LLP
 * Website        : https://tabgraf.com
 * Support Email  : support@tabgraf.com
 * 
 * Description    : An all-in-one employee self-service 
 *                  portal built with Google Apps Script 
 *                  for:
 *                    - Leave Management
 *                    - Asset Management
 *                    - Payroll Processing
 *                    - Payslip Download
 * 
 * © 2025 Tabgraf Technologies LLP. All rights reserved.
 * ------------------------------------------------------
 */

function test(){
//   let kk = insertLeave({
//     "fromDate": "2022-06-27",
//     "toDate": "2022-06-29",
//     "leaveType": "CL/SL",
//     "leaveReason": "hhh"
// });

const diffTime = Math.abs(new Date("2022-06-29 23:59:59") - new Date("2022-06-27 00:00:00"));
  const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24)); 

diffDays = diffDays;

}

function doGet(e) {
 var output = HtmlService.createTemplateFromFile("App").evaluate();
 return output.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function importFile_(fileName){
  return HtmlService.createTemplateFromFile(fileName).getRawContent().replace(/="\{([^}]*)\}"/g,'={$1}');
}

function cloneSheet_(spreadSheetUrl, sheetName){
  var sp = SpreadsheetApp.openByUrl(spreadSheetUrl);
  var sheet = sp.getSheetByName(sheetName);
  let newSheet = sheet.copyTo(sp).setName(new Date().getTime()+"_"+randomStr_(5));
  return newSheet;
}

function sheetToPdf_(spreadsheetUrl, sheetName, range, pdfFileName, folderId){
  let sp = SpreadsheetApp.openByUrl(spreadsheetUrl);
  let sheet = sp.getSheetByName(sheetName);
  var folderToSave = DriveApp.getFolderById(folderId);

  var url = `https://docs.google.com/spreadsheets/d/${sp.getId()}/export?exportFormat=pdf&format=pdf&gridlines=false&size=A4&gid=${sheet.getSheetId()}&range=${range}`;

  var request = {
    "method": "GET",
    "headers":{"Authorization": "Bearer "+ScriptApp.getOAuthToken()},    
    "muteHttpExceptions": true
  };

  var pdf = UrlFetchApp.fetch(url, request);
  pdf = pdf.getBlob().setName(pdfFileName);
  var file = folderToSave.createFile(pdf);
  return file;
}

// function getDirectoryInfo_(){
//   const userEmail = Session.getActiveUser().getEmail();
//   const user = AdminDirectory.Users.get(userEmail);
//   var request = {
//     "method": "GET",
//     "headers":{"Authorization": "Bearer "+ScriptApp.getOAuthToken()},    
//     "muteHttpExceptions": true
//   };
//   var response = UrlFetchApp.fetch(`https://admin.googleapis.com/admin/directory/v1/users/${userEmail}/photos/thumbnail`, request);
//   response = JSON.parse(response.getContentText());
//   let photo = "data:image/png;base64,"+response.photoData.replace(/_/g,'/').replace(/-/g,'+');
//   return {profilePic: photo, name: user.name.fullName};
// }

function updateRow_(sheetName, primaryKey, valuesToUpdate){
  let sp = SpreadsheetApp.getActiveSpreadsheet();
  let allData = sp.getSheetByName(sheetName).getDataRange().getValues();
  let headers = allData[0];
  let headersMap = {};
  headers.forEach((item, i)=>{
    headersMap[item] = i;
  });


  let targetIndex = -1;
  allData.forEach((row, i)=>{
    if(row[headersMap[primaryKey.name]] == primaryKey.value){
      targetIndex = i;
    }
  });

  valuesToUpdate.forEach((item, i)=>{
    if(item.value.includes("eval")){
      allData[targetIndex][headersMap[item.name]] = eval(item.value.replace("{VAL}",allData[targetIndex][headersMap[item.name]]));
    }else{
      allData[targetIndex][headersMap[item.name]] = item.value;
    }
  });

  sp.getSheetByName(sheetName).getRange(targetIndex+1,1,1,allData[0].length).setValues([allData[targetIndex]]);
}

function query_(sheet, query) {
  let sp = SpreadsheetApp.getActiveSpreadsheet();
  let columnIndex = sp.getSheetByName(sheet).getLastColumn();
  let lastColumn = columnToLetter_(columnIndex);
  let formula = `QUERY(${sheet}!A:${lastColumn}, "${query}")`;
  var sheet = sp.insertSheet();
  var r = sheet.getRange(1, 1).setFormula(formula);
  var reply = sheet.getDataRange().getDisplayValues();
  sp.deleteSheet(sheet);
  return reply;
}

function prependRow_(sheetName, rowData) {
  let sp = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = sp.getSheetByName(sheetName);
  sheet.insertRowBefore(2).getRange(2, 1, 1, rowData.length).setValues([rowData]);
}

const objectify2DArray_ = ([keys, ...values]) => 
  values.map(vs => Object.fromEntries(vs.map((v, i) => [keys[i], v])))

const shouldConsiderParticularForPayslip = (key, val) => {
  if(key.includes("Income Tax")){
    return true;
  }else{
    if(val && val.trim() != ""){
      return true;
    }else{
      return false;
    }
  }
}

/*** Line above this are internal functions ***/

function getLeaves(){
  let email = Session.getActiveUser().getEmail();
  let leaves = query_("Leaves",`select A,B,C,D,E,F,G, H where C='${email}'`);
  return {"status": "success", "data": objectify2DArray_(leaves)};
}

function getLeaveBalance(){
  let email = Session.getActiveUser().getEmail();
  let balances = query_("Employees",`select J,K where D='${email}'`);
  return {"status": "success", "data": objectify2DArray_(balances)};
}

function insertLeave(data){
  let email = Session.getActiveUser().getEmail();
  let dataToPretend = [Utilities.getUuid(), formatDate_(new Date(),"DD-MM-YYYY"),email,data.fromDate, data.toDate, data.leaveType, data.leaveReason];
  prependRow_("Leaves", dataToPretend);

  const diffTime = Math.abs(new Date(data.toDate+" 23:59:59") - new Date(data.fromDate+" 00:00:00"));
  const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24)); 

  updateRow_("Employees",{name: "Email", value: email}, [{name: data.leaveType, value: `eval({VAL}-${diffDays})`}]);

  return {"status": "success", data};
}

function getAssets(){
  let email = Session.getActiveUser().getEmail();
  let assets = query_("Assets",`select C,D,E,F,G,H where B='${email}'`);
  return {"status": "success", "data": objectify2DArray_(assets)};
}

function getProfileInfo(){
  let email = Session.getActiveUser().getEmail();
  let info = query_("Employees",`select A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U where D ='${email}'`);
  return {"status": "success", "data": objectify2DArray_(info)};
}

function getSalaries(){
  let email = Session.getActiveUser().getEmail();
  let salaries = query_("Payroll",`select A,B,C,E,F,G,H,I,J,K,L,M where D ='${email}'`);
  return {"status": "success", "data": objectify2DArray_(salaries)};
}

function getParticularName(name){
  return name.replace('[Earnings]','').replace('[Deductions]','').trim();
}

function generatePayslip(month){
  
  let email = Session.getActiveUser().getEmail();

  let info = query_("Employees",`select A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U where D ='${email}'`);
  info = objectify2DArray_(info);

  let salary = query_("Payroll",`select A,B,C,E,F,G,H,I,J,K,L,M,N where D ='${email}' and B='${month}'`);
  salary = objectify2DArray_(salary);

  let newSheet = cloneSheet_(PAYSLIP_TEMPLATE_URL, "Template1");
  let newSheetName = newSheet.getName();
  
  let dataToUpdate = [];
  //try{
  dataToUpdate.push({"range": `${newSheetName}!D10`,"values": [[info[0]['Name']]]});
  dataToUpdate.push({"range": `${newSheetName}!D11`,"values": [[info[0]['Emp No']]]});
  dataToUpdate.push({"range": `${newSheetName}!D12`,"values": [[info[0]['Gender']]]});
  dataToUpdate.push({"range": `${newSheetName}!D13`,"values": [[info[0]['Designation']]]});
  dataToUpdate.push({"range": `${newSheetName}!D14`,"values": [[info[0]['Work Location']]]});
  dataToUpdate.push({"range": `${newSheetName}!K10`,"values": [[info[0]['DOJ']]]});
  dataToUpdate.push({"range": `${newSheetName}!K11`,"values": [[info[0]['Bank Name']]]});
  dataToUpdate.push({"range": `${newSheetName}!K12`,"values": [[info[0]['Bank Account No']]]});
  dataToUpdate.push({"range": `${newSheetName}!K13`,"values": [[info[0]['PAN']]]});
  dataToUpdate.push({"range": `${newSheetName}!B8`,"values": [[salary[0]['Month']]]});
  dataToUpdate.push({"range": `${newSheetName}!R10`,"values": [[salary[0]['Days Paid']]]});
  dataToUpdate.push({"range": `${newSheetName}!R11`,"values": [[salary[0]['LOP']]]});
  dataToUpdate.push({"range": `${newSheetName}!R12`,"values": [[salary[0]['Pay Period']]]});



  let rowIndex = 19;
  for(let key in salary[0]){
    if(key.includes("[Earnings]") && shouldConsiderParticularForPayslip(key, salary[0][key])){
      dataToUpdate.push({"range": `${newSheetName}!B${rowIndex}`,"values": [[getParticularName(key)]]});
      dataToUpdate.push({"range": `${newSheetName}!I${rowIndex}`,"values": [[salary[0][key]]]});
      rowIndex++;
    }
  }

  rowIndex = 19;
  for(let key in salary[0]){
    if(key.includes("[Deductions]") && shouldConsiderParticularForPayslip(key, salary[0][key])){
      dataToUpdate.push({"range": `${newSheetName}!L${rowIndex}`,"values": [[getParticularName(key)]]});
      dataToUpdate.push({"range": `${newSheetName}!S${rowIndex}`,"values": [[salary[0][key]]]});
      rowIndex++;
    }
  }

  var resource = {
    valueInputOption: 'USER_ENTERED',
    data: dataToUpdate,
  };
  
  Sheets.Spreadsheets.Values.batchUpdate(resource, SpreadsheetApp.openByUrl(PAYSLIP_TEMPLATE_URL).getId());
  let netSalary = newSheet.getRange("E33").getValue();
  dataToUpdate.push({"range": `${newSheetName}!E34`,"values": [["Rupees "+ inWords_(netSalary) + "."]]});
  Sheets.Spreadsheets.Values.batchUpdate(resource, SpreadsheetApp.openByUrl(PAYSLIP_TEMPLATE_URL).getId());


  let pdfFile = sheetToPdf_(PAYSLIP_TEMPLATE_URL,newSheetName,"A1:V40",`Payslip_${info[0]['Emp No']}_${salary[0]['Month'].replace(' ','_')}`,PAYSLIP_FOLDER_PATH);

  pdfFile.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);

  updateRow_("Payroll",{name: "Serial Number", value: salary[0]['Serial Number']}, [{name: 'Payslip File ID', value: pdfFile.getId()}]);
  SpreadsheetApp.openByUrl(PAYSLIP_TEMPLATE_URL).deleteSheet(newSheet);
  return {"status": "success", "data": {"pdfFileId": pdfFile.getId()}};
  // }catch(e){
  //   return {"status": "success", "data": {e: e.toString(), line: e.stack, dataToUpdate, month}};
  // }
}

// function sendPayslipToEmail(month){
//   let email = Session.getActiveUser().getEmail();

//   let info = query_("Employees",`select A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U where D ='${email}'`);
//   info = objectify2DArray_(info);

//   let salary = query_("Payroll",`select A,B,C,E,F,G,H,I,J,K,L,M where D ='${email}' and B='${month}'`);
//   salary = objectify2DArray_(salary);

//   MailApp.sendEmail({
//     to: info[0]['Email'],
//     subject: `Your Payslip for ${month}`,
//     htmlBody: `Hello ${info[0]['Nick Name']}, <br/><br/>Please find your Payslip for ${month} in attachment.<br/><br/>Regards,<br/>Administrator`,
//     attachments: [DriveApp.getFileById(salary[0]['Payslip File ID'])]
//   });

// }
