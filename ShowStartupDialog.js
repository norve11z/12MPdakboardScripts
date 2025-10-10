    ////////////////////////////////////////////////////////////////////////////
   // This was created by Zachary Norvell, a Class of 2026 Computer Engineer // 
  // Intended for 12th Man Production DakBoards outside of Control Rooms    //
 // Was Made in the Summer of 2025 with the Assistance of Summer Whitlock  //
////////////////////////////////////////////////////////////////////////////



var isBroadcast = false;
function showStartupDialog(sport) {
  Logger.log("Showing Startup Dialog");
  if(sport != 'FB') {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      "Press 'Yes' for broadcast, 'No' for bigscreen",
      ui.ButtonSet.YES_NO
    );
    // Broadcast = true, Big Screen = false
    if (response === ui.Button.YES) {
      isBroadcast = true;
    } else if (response === ui.Button.NO) {
      isBroadcast = false;
    } else {
      return; // Dialog closed or cancelled
    }
  }

  productionData = {
    sport: sport || "",
    isBroadcast: isBroadcast || false,
    producer: "",
    producer2: "",
    director: "",
    td: "",
    dakman: "",
    ad: "",
    ap: "",
    xpr: "",
    oots: "",
    toc: "",
    bug: "",
    showControl: "",
    slash: "",
    cam3grip: "",
    cam4grip: "",
    cam5grip: "",
    cam6grip: "",
    dc1: "",
    dc2: "",
    dc3: "",
    dc4: "",
    wx1: "",
    wx2: "",
    wx3: "",
    cameras: []
  };

  var html;
  var sheet;
  var camsCount;
  var startRow;
  var cams = [];

  switch (sport) {
    case 'FB':

      Logger.log("Getting Current Football Values");
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('FB VideoBoard');
      SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sheet);
      camsCount = 7;
      startRow = findRow("GAME PROD");
      for (let i = 0; i < camsCount; i++) { cams.push(sheet.getRange((startRow+16) + i, 3).getValue()); }
      productionData = {
        sport: sport,
        isBroadcast: false,
        producer: sheet.getRange('C' + startRow).getValue(),
        producer2: sheet.getRange('C' + (startRow + 1)).getValue(),
        director: sheet.getRange('C' + (startRow + 2)).getValue(),
        td: sheet.getRange('C' + (startRow + 3)).getValue(),
        dakman: sheet.getRange('C' + (startRow + 4)).getValue(),
        ap: sheet.getRange('C' + (startRow + 5)).getValue(),
        xpr: sheet.getRange('C' + (startRow + 6)).getValue(),
        oots: sheet.getRange('C' + (startRow + 7)).getValue(),
        toc: sheet.getRange('C' + (startRow + 8)).getValue(),
        dc1: sheet.getRange('C' + (startRow + 9)).getValue(),
        dc2: sheet.getRange('C' + (startRow + 10)).getValue(),
        dc3: sheet.getRange('C' + (startRow + 11)).getValue(),
        dc4: sheet.getRange('C' + (startRow + 12)).getValue(),
        wx1: sheet.getRange('C' + (startRow + 13)).getValue(),
        wx2: sheet.getRange('C' + (startRow + 14)).getValue(),
        wx3: sheet.getRange('C' + (startRow + 15)).getValue(),
        cameras: cams
      };
      html = HtmlService.createTemplateFromFile('prompt');
      html.productionData = productionData;
      break;

    case 'MBB':
    case 'WBB':
      if(isBroadcast) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('MBB Broadcast');
        SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sheet);
        camsCount = 6;
        for (let i = 0; i < camsCount; i++) { cams.push(sheet.getRange(24 + i, 3).getValue()); }
        productionData = {
          sport: sport,
          isBroadcast: true,
          producer: sheet.getRange('C13').getValue(),
          director: sheet.getRange('C14').getValue(),
          ad: sheet.getRange('C15').getValue(),
          td: sheet.getRange('C16').getValue(),
          ap: sheet.getRange('C17').getValue(),
          xpr: sheet.getRange('C18').getValue(),
          bug: sheet.getRange('C19').getValue(),
          dc1: sheet.getRange('C20').getValue(),
          dc2: sheet.getRange('C21').getValue(),
          cam3grip: sheet.getRange('C22').getValue(),
          cam4grip: sheet.getRange('C23').getValue(),
          cameras: cams
        };
      } else {
        Logger.log("Getting Current BigScreen MBB Values");
        sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('MBB BigScreen');
        SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sheet);
        camsCount = 1;
        for (let i = 0; i < camsCount; i++) { cams.push(sheet.getRange(23 + i, 3).getValue()); }
        productionData = {
          sport: sport,
          isBroadcast: false,
          producer: sheet.getRange('C13').getValue(),
          director: sheet.getRange('C14').getValue(),
          td: sheet.getRange('C15').getValue(),
          showControl: sheet.getRange('C16').getValue(),
          xpr: sheet.getRange('C17').getValue(),
          wx1: sheet.getRange('C18').getValue(),
          wx2: sheet.getRange('C19').getValue(),
          slash: sheet.getRange('C20').getValue(),
          dc1: sheet.getRange('C21').getValue(),
          dc2: sheet.getRange('C22').getValue(),
          cameras: cams
        };
      }
      html = HtmlService.createTemplateFromFile('prompt');
      html.productionData = productionData;
      break;



    case 'SFB':    
      if(isBroadcast) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('SFB Broadcast');
        SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sheet);
        camsCount = 7;
        for (let i = 0; i < camsCount; i++) { cams.push(sheet.getRange(24 + i, 3).getValue()); }
        productionData = {
          sport: sport,
          isBroadcast: isBroadcast,
          producer: sheet.getRange('C13').getValue(),
          director: sheet.getRange('C14').getValue(),
          ad: sheet.getRange('C15').getValue(),
          ap: sheet.getRange('C16').getValue(),
          td: sheet.getRange('C17').getValue(),
          bug: sheet.getRange('C18').getValue(),
          xpr: sheet.getRange('C19').getValue(),
          toc: sheet.getRange('C20').getValue(),
          dc1: sheet.getRange('C21').getValue(),
          dc2: sheet.getRange('C22').getValue(),
          dc3: sheet.getRange('C23').getValue(),
          cameras: cams
        };
      } else {
        sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('SFB BigScreen');
        SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sheet);
        camsCount = 1;
        for (let i = 0; i < camsCount; i++) { cams.push(sheet.getRange(19 + i, 3).getValue()); }
        productionData = {
          sport: sport,
          isBroadcast: isBroadcast,
          producer: sheet.getRange('C13').getValue(),
          td: sheet.getRange('C14').getValue(),
          xpr: sheet.getRange('C15').getValue(),
          dc1: sheet.getRange('C16').getValue(),
          wx1: sheet.getRange('C17').getValue(),
          wx2: sheet.getRange('C18').getValue(),
        };
      }
      html = HtmlService.createTemplateFromFile('prompt');
      html.productionData = productionData;
      break;


    case 'BSB':
      if(isBroadcast) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('BSB Broadcast');
        SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sheet);
        camsCount = 7;
        for (let i = 0; i < camsCount; i++) { cams.push(sheet.getRange(23 + i, 3).getValue()); }
        productionData = {
          sport: sport,
          isBroadcast: isBroadcast,
          producer: sheet.getRange('C13').getValue(),
          director: sheet.getRange('C14').getValue(),
          ad: sheet.getRange('C15').getValue(),
          ap: sheet.getRange('C16').getValue(),
          td: sheet.getRange('C17').getValue(),
          bug: sheet.getRange('C18').getValue(),
          xpr: sheet.getRange('C19').getValue(),
          dc1: sheet.getRange('C20').getValue(),
          dc2: sheet.getRange('C21').getValue(),
          dc3: sheet.getRange('C22').getValue(),
          cameras: cams
        };
      } else {
        sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('BSB BigScreen');
        SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sheet);
        camsCount = 1;
        for (let i = 0; i < camsCount; i++) { cams.push(sheet.getRange(19 + i, 3).getValue()); }
        productionData = {
          sport: sport,
          isBroadcast: isBroadcast,
          producer: sheet.getRange('C13').getValue(),
          td: sheet.getRange('C14').getValue(),
          xpr: sheet.getRange('C15').getValue(),
          dc1: sheet.getRange('C16').getValue(),
          wx1: sheet.getRange('C17').getValue(),
          wx2: sheet.getRange('C18').getValue(),
        };
      }
      html = HtmlService.createTemplateFromFile('prompt');
      html.productionData = productionData;
      break;

    case 'VB':
      if(isBroadcast) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('VB Broadcast');
        SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sheet);
        camsCount = 7;
        startRow = findRow("PRODUCER");
        for (let i = 0; i < camsCount; i++) { cams.push(sheet.getRange(startRow+11 + i, 3).getValue()); }
        productionData = {
          sport: sport,
          isBroadcast: isBroadcast,
          producer: sheet.getRange('C' + (startRow + 0)).getValue(),
          director: sheet.getRange('C' + (startRow + 1)).getValue(),
          ad: sheet.getRange('C' + (startRow + 2)).getValue(),
          ap: sheet.getRange('C' + (startRow + 3)).getValue(),
          toc: sheet.getRange('C' + (startRow + 4)).getValue(),
          bug: sheet.getRange('C' + (startRow + 5)).getValue(),
          xpr: sheet.getRange('C' + (startRow + 6)).getValue(),
          dc1: sheet.getRange('C' + (startRow + 7)).getValue(),
          dc2: sheet.getRange('C' + (startRow + 8)).getValue(),
          cam5grip: sheet.getRange('C' + (startRow + 9)).getValue(),
          cam6grip: sheet.getRange('C' + (startRow + 10)).getValue(),
          cameras: cams
        };

      } else {
        sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('VB BigScreen');
        SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sheet);
        camsCount = 1;
        for (let i = 0; i < camsCount; i++) { cams.push(sheet.getRange(19 + i, 3).getValue()); }
        productionData = {
          sport: sport,
          isBroadcast: isBroadcast,
          producer: sheet.getRange('C13').getValue(),
          td: sheet.getRange('C14').getValue(),
          xpr: sheet.getRange('C15').getValue(),
          dc1: sheet.getRange('C16').getValue(),
          wx1: sheet.getRange('C17').getValue(),
          wx2: sheet.getRange('C18').getValue(),
        };
      }
      html = HtmlService.createTemplateFromFile('prompt');
      html.productionData = productionData;
      break;


    case 'SOC':
      Logger.log("Getting SOC Data");

      if(isBroadcast) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('SOC Broadcast');
        SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sheet);
        camsCount = 7;
        startRow = findRow("PRODUCER");
        for (let i = 0; i < camsCount; i++) { cams.push(sheet.getRange((startRow + 10) + i, 3).getValue()); }
        productionData.producer = sheet.getRange('C' + (startRow)).getValue();
        productionData.director = sheet.getRange('C' + (startRow + 1)).getValue();
        productionData.td = sheet.getRange('C' + (startRow + 2)).getValue();
        productionData.ad = sheet.getRange('C' + (startRow + 3)).getValue();
        productionData.ap = sheet.getRange('C' + (startRow + 4)).getValue();
        productionData.xpr = sheet.getRange('C' + (startRow + 5)).getValue();
        productionData.bug = sheet.getRange('C' + (startRow + 6)).getValue();
        productionData.toc = sheet.getRange('C' + (startRow + 7)).getValue();
        productionData.dc1 = sheet.getRange('C' + (startRow + 8)).getValue();
        productionData.dc2 = sheet.getRange('C' + (startRow + 9)).getValue();
        productionData.cameras = cams;

      } else {
        Logger.log("Getting SOC BigScreen Data");
        sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('SOC BigScreen');
        SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sheet);
        startRow = findRow("PRODUCER");
        Logger.log('C' + startRow);
        productionData.producer = sheet.getRange('C' + (startRow + 0)).getValue();
        productionData.td = sheet.getRange('C' + (startRow + 1)).getValue();
        productionData.xpr = sheet.getRange('C' + (startRow + 2)).getValue();
        productionData.dc1 = sheet.getRange('C' + (startRow + 3)).getValue();

      }
      html = HtmlService.createTemplateFromFile('prompt');
      html.productionData = productionData;
      Logger.log("Making SOC Prompt");

      break;

    default:
      SpreadsheetApp.getUi().alert("Unknown sport: " + sport);
      break;
  }
  SpreadsheetApp.getUi().showModalDialog(html.evaluate().setWidth(600).setHeight(650), 'Please Enter Game Details');
}
