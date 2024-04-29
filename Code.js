function onEdit(e) {
  var sheet = e.source.getSheetByName('Worksheet');
  var range = e.range;

  // Dropdown is in cell F12
  if (range.getA1Notation() === 'F12') {
    var selectedProfile = range.getValue();

    var profilesSheet = e.source.getSheetByName('Profiles');
    var profilesData = profilesSheet.getDataRange().getValues();

    var valueCells = ['A12', 'A14', 'A15', 'D14', 'D15', 'D16', 'F14', 'F15', 'F16', 'F17', 'F18', 'E20', 'G23', 'G24', 'I23', 'U26', 'Z26', 'U30', 'Z30', 'U42'];

    // Clear existing valueCells' data in Worksheet
    valueCells.forEach(function (cell) {
      sheet.getRange(cell).clearContent();
    });

    // Find the selected profile in the Profiles sheet
    for (var i = 1; i < profilesData.length; i++) {
      if (profilesData[i][0] === selectedProfile) {
        // Spouse Information
        if (profilesData[i][4] === 'Married') {
          sheet.getRange('A12').setValue('Spouse Information:').setFontWeight("bold");
          sheet.getRange('A14').setValue('Profession: ' + profilesData[i][9]);
          sheet.getRange('A15').setValue('Take Home Pay: ' + formatAsDollars(profilesData[i][10]));
          // Spouse take home + profile's take home
        sheet.getRange('E20').setValue(formatAsDollars(profilesData[i][10]) + ' + ' + formatAsDollars(profilesData[i][3]) + ' = ' + (formatAsDollars(profilesData[i][10] + profilesData[i][3])));
        }

        // Gross & monthly pay
        sheet.getRange('D14').setValue('Gross Yearly Salary: ' + formatAsDollars(profilesData[i][1]));
        sheet.getRange('D15').setValue('Gross Monthly Salary: ' + formatAsDollars(profilesData[i][2]));
        sheet.getRange('D16').setValue('Monthly Take Home Pay: ' + formatAsDollars(profilesData[i][3]));

        // Marital status and expenses
        sheet.getRange('F14').setValue('Marital Status: ' + profilesData[i][4]);
        sheet.getRange('F15').setValue('Child(ren): ' + profilesData[i][5]);
        sheet.getRange('F16').setValue('Education Required: ' + profilesData[i][6]);
        sheet.getRange('F17').setValue('Monthly Student Loan Payment: ' + formatAsDollars(profilesData[i][7]));
        sheet.getRange('F18').setValue('Monthly Retirement Contribution: ' + formatAsDollars(profilesData[i][8]));

        // Single profile's take home
        if (profilesData[i][4] === 'Single') {
          sheet.getRange('E20').setValue(formatAsDollars(profilesData[i][3]));
        }
        
        // Other Expenses
        sheet.getRange('G23').setValue(formatAsDollars(profilesData[i][8]));
        sheet.getRange('G24').setValue(formatAsDollars(profilesData[i][7]));
        
        var firstRemainingBalance = (profilesData[i][10] + profilesData[i][3]) - profilesData[i][8];
        sheet.getRange('I23').setValue(formatAsDollars(firstRemainingBalance));

        // W-2
        sheet.getRange('U26').setValue(profilesData[i][11]); // Wages, tips, other compensation
        sheet.getRange('Z26').setValue(profilesData[i][12]); // Federal income tax withheld
        sheet.getRange('U30').setValue(profilesData[i][13]); // Medicare wages and tips
        sheet.getRange('Z30').setValue(profilesData[i][14]); // Medicare tax withheld
        sheet.getRange('U42').setValue(profilesData[i][15]); // Other
        sheet.getRange('Z28').setValue(profilesData[i][16]); // Social security tax withheld
        sheet.getRange('U28').setValue(profilesData[i][13]); // Social security wages

        break;
      }
    }
  }
}

// Function to format a number as dollars
function formatAsDollars(number) {
  return '$' + Number(number).toFixed(2).replace(/(\d)(?=(\d{3})+\.)/g, '$1,');
}
