function showStartupDialog(sport) {
  Logger.log("Showing Startup Dialog");

  if(sport != 'FB') {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      "Press 'Yes' for broadcast, 'No' for bigscreen",
      ui.ButtonSet.YES_NO
    );
    // Broadcast = true, Big Screen = false
    var isBroadcast;
    if (response === ui.Button.YES) {
      isBroadcast = true;
    } else if (response === ui.Button.NO) {
      isBroadcast = false;
    } else {
      return; // Dialog closed or cancelled
    }
  }

  var html;
  var sheet;
  var camsCount;
  var cams = [];

  switch (sport) {
    case 'FB':

      Logger.log("Getting Current Football Values");
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('FB VideoBoard');
      SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sheet);
      camsCount = 7;
      for (let i = 0; i < camsCount; i++) { cams.push(sheet.getRange(29 + i, 3).getValue()); }
      productionData = {
        sport: 'FB',
        isBroadcast: false,
        producer: sheet.getRange('C14').getValue(),
        director: sheet.getRange('C15').getValue(),
        td: sheet.getRange('C16').getValue(),
        dakman: sheet.getRange('C17').getValue(),
        ap: sheet.getRange('C18').getValue(),
        xpr: sheet.getRange('C19').getValue(),
        oots: sheet.getRange('C20').getValue(),
        toc: sheet.getRange('C21').getValue(),
        dc1: sheet.getRange('C22').getValue(),
        dc2: sheet.getRange('C23').getValue(),
        dc3: sheet.getRange('C24').getValue(),
        dc4: sheet.getRange('C25').getValue(),
        wx1: sheet.getRange('C26').getValue(),
        wx2: sheet.getRange('C27').getValue(),
        wx3: sheet.getRange('C28').getValue(),
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
          sport: 'MBB',
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
        html = HtmlService.createTemplateFromFile('prompt');

      } else {

        Logger.log("Getting Current BigScreen MBB Values");
        sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('MBB BigScreen');
        SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sheet);
        camsCount = 1;
        for (let i = 0; i < camsCount; i++) { cams.push(sheet.getRange(23 + i, 3).getValue()); }
        productionData = {
          sport: 'MBB',
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
        html = HtmlService.createTemplateFromFile('prompt');
      }

      html.productionData = productionData;
      break;

    case 'SB':    
    case 'BSB':
    //bsb no toc!
      if(isBroadcast) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('BSB Broadcast');
        camsCount = 3;
        for (let i = 0; i < camsCount; i++) { cams.push(sheet.getRange(23 + i, 3).getValue()); }
        productionData = {
          sport: 'BSB',
          isBroadcast: true,
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
        html = HtmlService.createTemplateFromFile('prompt');

      } else {

        sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('BSB BigScreen');
        camsCount = 3;
        for (let i = 0; i < camsCount; i++) { cams.push(sheet.getRange(23 + i, 3).getValue()); }
        productionData = {
          sport: 'BSB',
          isBroadcast: false,
          producer: sheet.getRange('C13').getValue(),
          td: sheet.getRange('C14').getValue(),
          xpr: sheet.getRange('C15').getValue(),
          dc1: sheet.getRange('C16').getValue(),
          wx1: sheet.getRange('C17').getValue(),
          wx2: sheet.getRange('C18').getValue(),
        };
        html = HtmlService.createTemplateFromFile('prompt');
      }

      html.productionData = productionData;
      break;

    case 'VB':
    case 'SOC':
    //SOC has no grips!
      if(isBroadcast) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('MBB');
        camsCount = 3;
        for (let i = 0; i < camsCount; i++) { cams.push(sheet.getRange(23 + i, 3).getValue()); }
        productionData = {
          sport: 'VB',
          isBroadcast: true,
          producer: sheet.getRange('C13').getValue(),
          director: sheet.getRange('C14').getValue(),
          ad: sheet.getRange('C15').getValue(),
          ap: sheet.getRange('C15').getValue(),
          toc: sheet.getRange('C16').getValue(),
          bug: sheet.getRange('C15').getValue(),
          xpr: sheet.getRange('C17').getValue(),
          dc1: sheet.getRange('C18').getValue(),
          dc2: sheet.getRange('C19').getValue(),
          cam5grip: sheet.getRange('C19').getValue(),
          cam6grip: sheet.getRange('C19').getValue(),
          cameras: cams
        };
        html = HtmlService.createTemplateFromFile('fb_prompt');

      } else {
      //SOC has no wx1, wx2!
        sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('MBB');
        camsCount = 3;
        for (let i = 0; i < camsCount; i++) { cams.push(sheet.getRange(23 + i, 3).getValue()); }
        productionData = {
          sport: 'VB',
          isBroadcast: false,
          producer: sheet.getRange('C13').getValue(),
          td: sheet.getRange('C16').getValue(),
          xpr: sheet.getRange('C17').getValue(),
          dc1: sheet.getRange('C18').getValue(),
          wx1: sheet.getRange('C20').getValue(),
          wx2: sheet.getRange('C21').getValue(),
        };
        html = HtmlService.createTemplateFromFile('fb_prompt');
      }

      html.productionData = productionData;
      break;

    default:
      SpreadsheetApp.getUi().alert("Unknown sport: " + sport);
      break;
  }
  SpreadsheetApp.getUi().showModalDialog(html.evaluate().setWidth(600).setHeight(650), 'Please Enter Game Details');
}
