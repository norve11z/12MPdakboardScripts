function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📝 Edit Schedule')
    .addItem('🏈 Football', 'showFBstartupDialog')   // Done
    .addItem('🏀 Men\'s Basketball', 'showMBBstartupDialog') // BigScreen Done, Broadcast
    .addItem('🏀 Women\'s Basketball', 'showWBBstartupDialog')
    .addItem('⚾ Baseball', 'showBSBstartupDialog')
    .addItem('⚾ Softball', 'showSBstartupDialog')
    .addItem('⚾ Soccer', 'showSOCstartupDialog')
    .addItem('⚾ Volleyball', 'showVBstartupDialog')
    .addToUi();
}

//For Styling
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

//utility functions
var toCaps = (str) => (str ? String(str).toUpperCase() : '');
function parseTime(timeStr) {
  const today = new Date();
  const datePart = Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const fullStr = `${datePart} ${timeStr}`;
  const parsed = new Date(fullStr);
  return isNaN(parsed) ? null : parsed;
}

// Wrappers for each sport
function showFBstartupDialog() { showStartupDialog('FB'); }
function showMBBstartupDialog() { showStartupDialog('MBB'); }
function showWBBstartupDialog() { showStartupDialog('WBB'); }
function showBSBstartupDialog() { showStartupDialog('BSB'); }
function showSBstartupDialog()  { showStartupDialog('SB'); }
function showSOCstartupDialog() { showStartupDialog('SOC'); }
function showVBstartupDialog()  { showStartupDialog('VB'); }

// testing changes









