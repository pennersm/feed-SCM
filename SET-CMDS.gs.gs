///---------------------------------------------------------------------------------
///  "CLI SECTION" sheet : generate the set-commands
/// --------------------------------------------------------------------------------
///===> This is goiung to the sheet and reading the table:
function ProcessTabByRow() {
  const sheetName = 'CLI Section';
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const ui = SpreadsheetApp.getUi();

  if (!sheet) {
    ui.alert(`❌ Sheet "${sheetName}" not found.`);
    return;
  }
  sheet.getRange('D5:D').clearContent();
  const startRow = 5;
  const colB = sheet.getRange(startRow, 2, sheet.getLastRow() - startRow + 1).getValues(); // Col B
  const offset = startRow;

  for (let i = 0; i < colB.length; i++) {
    const value = (colB[i][0] || '').toString().trim();

    if (!value || value.toLowerCase() === 'no') {
      continue; // Skip blank or 'no'
    }

    const rowIndex = i + offset;

    try {
      callRowFunction(rowIndex, value); // set cmd series for this row
    } catch (err) {
      Logger.log(`❌ Error in row ${rowIndex}: ${err.message}`);
    }
  }
}

function appendCommandLines(lines) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CLI Section');
  const colD = sheet.getRange('D5:D').getValues();
  let lastRow = 4;

  for (let i = 0; i < colD.length; i++) {
    if (colD[i][0]) lastRow = i + 5; // row index + offset
  }

  sheet.getRange(lastRow + 1, 4, lines.length, 1).setValues(lines.map(line => [line]));
}

function callRowFunction(rowNumber, inputValue) {
  switch (rowNumber) {
    case 5:
      set_timezone(inputValue);
      break;
    case 6:
      set_region(inputValue);
      break;
    case 7:
      set_dns_1 (inputValue);
      break;
    case 9:
      set_ntp_server_1(inputValue);
      break;
    case 15:
      set_cloud_log_on(inputValue)
      break;
    case 20:
      set_apn_profile(inputValue);
      break;
    case 22:
      set_cellular_if_2_sim;
      break;
    default:
      Logger.log(`ℹ️ No handler defined for row ${rowNumber}`);
  }
}

// sety command generator functions
function set_timezone(val) {
  //const message = `timezone set to: ${val}\n→ Test pattern: APPLY_ROW6_${val}`;
  //SpreadsheetApp.getUi().alert(message)
  lines = [
    `set deviceconfig system timezone ${val}`
  ];
  appendCommandLines(lines);
  Logger.log(lines);
  lines.length = 0 ;
}
var ngfw_region = ''
function set_region(val) {
  //const message = `region set to: ${val}\n→ Test pattern: APPLY_ROW6_${val}`;
  //SpreadsheetApp.getUi().alert(message)
  lines = [
    `set deviceconfig system device-telemetry region ${val}`,
    `set deviceconfig system region ${val}`
  ];
  ngfw_region = val;
  appendCommandLines(lines);
  Logger.log(lines);
  lines.length = 0 ;
}

function set_dns_1(val) {
  //const message = `primary DNS set to: ${val}\n→ Test pattern: APPLY_ROW6_${val}`;
  //SpreadsheetApp.getUi().alert(message)
    lines = [
    `set deviceconfig system dns-setting servers primary ${val}`
  ];
  appendCommandLines(lines);
  Logger.log(lines);
  lines.length = 0 ;
}
function set_ntp_1(val) {
  //const message = `primary NTP set to: ${val}\n→ Test pattern: APPLY_ROW6_${val}`;
  //SpreadsheetApp.getUi().alert(message)
    lines = [
    `set deviceconfig system ntp-servers primary-ntp-server ${val}`
  ];
  appendCommandLines(lines);
  Logger.log(lines);
  lines.length = 0 ;
}
function set_cloud_log_on(val) {
  //const message = `primary NTP set to: ${val}\n→ Test pattern: APPLY_ROW6_${val}`;
  //SpreadsheetApp.getUi().alert(message)
    lines = [
    `set deviceconfig setting logging logging-service-forwarding enable`,
    `set deviceconfig setting logging enhanced-application-logging enable`,   
    `set deviceconfig setting logging logging-service-forwarding logging-service-regions` + ngfw_region
  ];
  appendCommandLines(lines);
  Logger.log(lines);
  lines.length = 0 ;
}
function set_apn_profile(val) {
  const profile = 'apnProfile_' + val;

  const lines = [
    'set network interface cellular cellular1/1 modem radio on',
    'set network interface cellular cellular1/1 modem gps on',
    'set network interface cellular cellular1/1 sim primary slot 1',
    'set network interface cellular cellular1/1 apn-profile ' + profile + ' apn ' + val
  ];
  appendCommandLines(lines);
  Logger.log(lines); // You had Logger.log(result) but 'result' is undefined
  lines.length = 0;
}
