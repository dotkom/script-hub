/*
 *  This script serves to automatically sort committee applications into separate spreadsheets
 *  for the respective committees. To use it, create a form for submitting applications (copy an old one),
 *  connect a spreadsheet to the form, add a "Google Apps-script" extention to the sheet, paste the following
 *  code into a .gs file and run setUpTrigger(). The sheets will now automatically update whenever a new
 *  application is submitted
 */

function updateCommitteesAndGroups() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mainSheet = ss.getSheetByName("Skjemasvar 1");
  var data = mainSheet.getDataRange().getValues();

  var headers = data[0];
  var applicants = data.slice(1);

  var committeeColumns = {
    førstevalg: headers.findIndex((col) =>
      col.toLowerCase().includes("førstevalg")
    ),
    andrevalg: headers.findIndex((col) =>
      col.toLowerCase().includes("andrevalg")
    ),
    tredjevalg: headers.findIndex((col) =>
      col.toLowerCase().includes("tredjevalg")
    ),
    Backlog: headers.findIndex((col) => col.toLowerCase().includes("backlog")),
    FeminIT: headers.findIndex((col) => col.toLowerCase().includes("feminit")),
  };

  processApplicants(applicants, committeeColumns, headers);
}

function processApplicants(applicants, committeeColumns, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  applicants.forEach(function (applicant) {
    var addedCommittees = [];
    for (let keyword in committeeColumns) {
      let col = committeeColumns[keyword];
      let name = applicant[col];

      if (keyword === "Backlog" || keyword === "FeminIT") {
        if (name.includes("ønsker å søke verv")) {
          name = keyword;
        } else {
          continue;
        }
      }

      if (name && !addedCommittees.includes(name)) {
        let sheet = ss.getSheetByName(name);
        if (!sheet) {
          sheet = ss.insertSheet(name);
          sheet.appendRow(
            headers.filter(
              (_, idx) => !Object.values(committeeColumns).includes(idx)
            )
          );
        }

        var emailCol = headers.findIndex((col) =>
          col.toLowerCase().includes("e-postadresse")
        );
        var numRows = sheet.getLastRow() - 1;
        if (numRows >= 1) {
          var existingApplicants = sheet
            .getRange(2, emailCol + 1, numRows, 1)
            .getValues();
          if (!existingApplicants.flat().includes(applicant[emailCol])) {
            sheet.appendRow(
              applicant.filter(
                (_, idx) => !Object.values(committeeColumns).includes(idx)
              )
            );
            addedCommittees.push(name);
          }
        } else {
          sheet.appendRow(
            applicant.filter(
              (_, idx) => !Object.values(committeeColumns).includes(idx)
            )
          );
          addedCommittees.push(name);
        }
      }
    }
  });
}

function setUpTrigger() {
  ScriptApp.newTrigger("updateCommitteesAndGroups")
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onFormSubmit()
    .create();
}
