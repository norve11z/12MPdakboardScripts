function handleStartupForm(dateStr, timeStr, team, sport, isBroadcast) {
    Logger.log("Handling Start Up Form");

  var sheetName;
  var teamCord;
  var dateCord;
  var offsets;
  var timeRange;

  switch (sport) {
    case 'FB':
      sheetName = "FB VideoBoard";
      offsets = [315, 300, 285,255, 195, 155 , 135, 0];
      timeRange = "B5:B12";
      teamCord = "B3";
      dateCord = "B4";
    break;

    case 'MBB':

      if(isBroadcast) {
        sheetName = "MBB Broadcast";
      } else {
        sheetName = "MBB BigScreen";
      }
      offsets = [190, 185, 130, 105, 50, 30, 0];
      timeRange = "B5:B11";
      teamCord = "B3";
      dateCord = "B4";
    break;

    case 'WBB':
      if(isBroadcast) {
        sheetName = "MBB Broadcast";
      } else {
        sheetName = "MBB BigScreen";
      }
      offsets = [190, 185, 130, 105, 50, 30, 0];
      timeRange = "B5:B11";
      teamCord = "B3";
      dateCord = "B4";
    break;

    case 'BSB':
      if(isBroadcast) {
        sheetName = "BSB Broadcast";
      } else {
        sheetName = "BSB BigScreen";
      }
      offsets = [190, 185, 130, 105, 50, 30, 0];
      timeRange = "B5:B11";
      teamCord = "B3";
      dateCord = "B4";

    break;

    case 'SFB':
      if(isBroadcast) {
        sheetName = "SFB Broadcast";

      } else {
        sheetName = "SFB BigScreen";
      }
      offsets = [190, 185, 130, 105, 50, 30, 0];
      timeRange = "B5:B11";
      teamCord = "B3";
      dateCord = "B4";
    break;

    case 'SOC':
      if(isBroadcast) {
        sheetName = "SOC Broadcast";
        offsets = [240, 210, 150, 120, 90, 70, 30, 15, 0];
        timeRange = "B5:B13";
      } else {
        sheetName = "SOC BigScreen";
        offsets = [190, 185, 130, 105, 50, 30, 0];
        timeRange = "B5:B11";
      }
      teamCord = "B3";
      dateCord = "B4";
    break;

    case 'VB':
      if(isBroadcast) {
        sheetName = "VB Broadcast";
      } else {
        sheetName = "VB BigScreen";
      }
      offsets = [190, 185, 130, 105, 50, 30, 0];
      timeRange = "B5:B11";
      teamCord = "B3";
      dateCord = "B4";
    break;
  }




  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    SpreadsheetApp.getUi().alert('Sheet named ${sheetName} not found.');
    return;
  }

  // Write opponent info in B3
  if (team) {
    sheet.getRange(teamCord).setValue(`TEXAS A&M vs ${toCaps(team)}`);
  }

  // Format and write date to B4
if (dateStr) {
  const parts = dateStr.split('-'); // ["yyyy", "MM", "dd"]
  const parsedDate = new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2])); // Month is 0-based
  const formattedDate = Utilities.formatDate(parsedDate, Session.getScriptTimeZone(), 'MMMM d, yyyy');
  sheet.getRange(dateCord).setValue(formattedDate);
}


  // Parse tip-off time
  if (timeStr) {
    const tipOffDateTime = parseTime(timeStr);
    if (!tipOffDateTime) {
      SpreadsheetApp.getUi().alert('Invalid Start time.');
      return;
    }
    // Offsets in minutes for rows B5 to B11
    range = sheet.getRange(timeRange);

    for (let i = 0; i < offsets.length; i++) {
      const time = new Date(tipOffDateTime.getTime() - offsets[i] * 60000);
      const timeString = Utilities.formatDate(time, Session.getScriptTimeZone(), 'h:mm a');
      range.getCell(i + 1, 1).setValue(timeString);
    }
  }
}