/*
 *  This script serves to automatically sort committee applications into separate spreadsheets
 *  for the respective committees. To use it, create a form for submitting applications (copy an old one),
 *  connect a spreadsheet to the form, add a "Google Apps-script" extention to the sheet, paste the following
 *  code into a .gs file and run setUpTrigger() ONCE. The sheets will now automatically update whenever a new
 *  application is submitted
 */

function updateCommitteesAndGroups() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var mainSheet = ss.getSheetByName("Skjemasvar 1");
    var data = mainSheet.getDataRange().getValues();

    var headers = data[0];
    var applicants = data.slice(1);

    /* The columns containing committes the applicant applied for */
    var committeeColumns = {
        førstevalg: headers.findIndex((col) =>
            col.toLowerCase().includes("førstevalg"),
        ),
        andrevalg: headers.findIndex((col) =>
            col.toLowerCase().includes("andrevalg"),
        ),
        tredjevalg: headers.findIndex((col) =>
            col.toLowerCase().includes("tredjevalg"),
        ),
        Backlog: headers.findIndex((col) =>
            col.toLowerCase().includes("backlog"),
        ),
        FeminIT: headers.findIndex((col) =>
            col.toLowerCase().includes("feminit"),
        ),
    };

    processApplicants(applicants, committeeColumns, headers);
}

function processApplicants(applicants, committeeColumns, headers) {
    var folder = DriveApp.getFileById(
        SpreadsheetApp.getActiveSpreadsheet().getId(),
    )
        .getParents()
        .next();

    applicants.forEach(function (applicant) {
        var addedCommittees = [];
        for (let keyword in committeeColumns) {
            let col = committeeColumns[keyword];
            let committee = applicant[col];

            /* If Backlog or FeminIT is checked off, register the applicant for that committee */
            if (keyword === "Backlog" || keyword === "FeminIT") {
                if (committee.includes("ønsker å søke verv")) {
                    committee = keyword;
                } else {
                    continue;
                }
            }

            if (committee && !addedCommittees.includes(committee)) {
                let existingFile = folder.getFilesByName(committee).hasNext();
                var ss;
                if (existingFile) {
                    ss = SpreadsheetApp.open(
                        folder.getFilesByName(committee).next(),
                    );
                } else {
                    ss = SpreadsheetApp.create(committee);
                    var fileId = ss.getId();
                    var file = DriveApp.getFileById(fileId);
                    file.moveTo(folder);
                }
                let sheets = ss.getSheets();
                let sheet = sheets[0];

                /* The header filters out which committees an applicant applied for, so the recieving committee can't see the how they're prioritized */
                if (sheet.getLastRow() === 0) {
                    sheet.appendRow(
                        headers.filter(
                            (_, idx) =>
                                !Object.values(committeeColumns).includes(idx),
                        ),
                    );
                }

                var emailCol = headers.findIndex((col) =>
                    col.toLowerCase().includes("e-postadresse"),
                );
                var numRows = sheet.getLastRow() - 1;
                if (numRows >= 1) {
                    var existingApplicants = sheet
                        .getRange(2, emailCol + 1, numRows, 1)
                        .getValues();
                    if (
                        !existingApplicants.flat().includes(applicant[emailCol])
                    ) {
                        sheet.appendRow(
                            applicant.filter(
                                (_, idx) =>
                                    !Object.values(committeeColumns).includes(
                                        idx,
                                    ),
                            ),
                        );
                        addedCommittees.push(committee);
                    }
                } else {
                    sheet.appendRow(
                        applicant.filter(
                            (_, idx) =>
                                !Object.values(committeeColumns).includes(idx),
                        ),
                    );
                    addedCommittees.push(committee);
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
