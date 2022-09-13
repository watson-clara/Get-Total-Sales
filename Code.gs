// @ts-nocheck
function main() {
  Logger.log("STARTING loadXLSX()");
  var date = getReports();
  combineToSingleSS(date);
  addSheet();
}




function getReports() {
  // calls specific folder in google drive
  const fldr = DriveApp.getFolderById("1-0WibmmZpk7xqOHvdw_S5UHtL784GgMX");
  Logger.log("got xlsx folder");
  // gets the xlsx files in the specific folder
  const xlsx_files = fldr.getFiles();
  Logger.log("got xlsx files");
  // creates bool to know if it is the first loop and to create file 
  var count = 0;
  // loop throough every xlsx file in folder
  while (xlsx_files.hasNext()) {
    count = count + 1;
    Logger.log("NEW FILE #" + count);
    // sleep
    Utilities.sleep(1000);
    // find the next xlsx file 
    var xlsx = xlsx_files.next();
    // get the xlsx file name and id
    var xlsx_name = xlsx.getName();
    var id = xlsx.getId();
    Logger.log("xlsx file name:  " + xlsx_name);
    // splits the file name to get the name and date
    var name = xlsx_name.split("_")[1];
    var date = xlsx_name.split("_")[2];
    date = date.split(".")[0];
    Logger.log("new SS name:  " + name);
    Logger.log("date:  " + date);
    // copies data in xlsx file to copy to new google sheet file 
    var blob = xlsx.getBlob();
    // sleep
    Utilities.sleep(1000);
    // creates new google sheet file to convert xlsx to 
    var newFile = {
      title: name,
      parents: [{ id: "1SzC0e7TCEMGwdJEgIjo918iQawjM3CYX" }],
      mimeType: MimeType.GOOGLE_SHEETS
    };
    Logger.log("new SS file created");

    // copies the xlsx content to the new file 
    Drive.Files.insert(newFile, blob);
    // deletes the xlsx file 
    Drive.Files.remove(id);
    Logger.log("old files deleted")
    Logger.log("done")
    // sleep 
    Utilities.sleep(1000);
  }
  return date;
}

function sortArray(allsheets) {
  // creates empty array to store sheet names
  var sheetNameArray = [];
  // loop through sheets and add each name to array
  for (var i = 0; i < allsheets.length; i++) {
    sheetNameArray.push(allsheets[i].getName());
  }
  // sort the array 
  sheetNameArray.sort(function (a, b) {
    return a.localeCompare(b);
  });
  Logger.log(sheetNameArray);
  // return sorted array
  return sheetNameArray;
}


function combineToSingleSS(date) {
  // calls specific folder in google drive
  const fldr = DriveApp.getFolderById("1SzC0e7TCEMGwdJEgIjo918iQawjM3CYX");
  Logger.log("got folder");
  // gets the converted google sheet files in the specific folder
  const files = fldr.getFiles();

  Logger.log("got files");
  Logger.log(files);
  // creates new spreadsheet for data in individual xlsx sheets to be uploaded to
  var combinedSS = SpreadsheetApp.create("RAW SALES CONCAT " + date);
  // gets new spreadsheet id 
  var id = combinedSS.getId();
  Logger.log(combinedSS.getName());
  // opens new spreadsheet so it can be edited
  var newSS = SpreadsheetApp.openById(id);
  // keeps track of what file we are on
  var count = 0;
  // loops through GS files in folder to add to new spreadsheet
  while (files.hasNext()) {
    count = count + 1;
    // sleep
    Utilities.sleep(1000);
    // gets next file
    var file = files.next();
    const idDel = file.getId();
    Logger.log(file);
    // gets next file name 
    var file_name = file.getName();
    Logger.log(file_name)
    // opens source spreadsheet
    var sh = SpreadsheetApp.openByUrl(file.getUrl());
    // gets sheet in that spreadsheets
    var target_sh = sh.getSheets()[0];
    // sleep
    Utilities.sleep(1000);
    // copies the sheet from the source spreadsheet to the combined spreadsheeyt
    target_sh.copyTo(newSS).setName(file_name);
    // deletes the old file 
    Drive.Files.remove(idDel);
    Logger.log("old files deleted")

  }
  addSheet(newSS);
}



function addSheet(newSS) {
  // opens new spreadsheet so it can be edited
  var sh = SpreadsheetApp.openById("1JI-1Bp39dgqcdQwzSPhjCw2IAQ7yyb_l0fFKu2fTWY4");
  // gets sheet in that spreadsheets
    var target_sh = sh.getSheets()[0];
    // sleep
    Utilities.sleep(1000);
    // copies the sheet from the source spreadsheet to the combined spreadsheeyt
    target_sh.copyTo(newSS).setName("QUERY");
 
}




