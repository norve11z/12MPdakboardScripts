function saveProductionData(formData) {
    Logger.log("Saving Production Data");      

  var sheet;
  var cams;
  var writeRow;
  var camName;
  var prodRow;
  var isBroadcast = formData.isBroadcast;


  switch (formData.sport) {
    case 'FB':
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('FB VideoBoard');
      prodRow = findRow("PRODUCER");
      sheet.showRows(prodRow, prodRow + 14);
      sheet.getRange('C' + prodRow).setValue(toCaps(formData.producer));
      sheet.getRange('C' + (prodRow + 1)).setValue(toCaps(formData.director));
      sheet.getRange('C' + (prodRow + 2)).setValue(toCaps(formData.td));
      sheet.getRange('C' + (prodRow + 3)).setValue(toCaps(formData.dakman));
      sheet.getRange('C' + (prodRow + 4)).setValue(toCaps(formData.ap));
      sheet.getRange('C' + (prodRow + 5)).setValue(toCaps(formData.xpr));
      sheet.getRange('C' + (prodRow + 6)).setValue(toCaps(formData.oots));
      sheet.getRange('C' + (prodRow + 7)).setValue(toCaps(formData.toc));
      sheet.getRange('C' + (prodRow + 8)).setValue(toCaps(formData.dc1));
      sheet.getRange('C' + (prodRow + 9)).setValue(toCaps(formData.dc2));
      sheet.getRange('C' + (prodRow + 10)).setValue(toCaps(formData.dc3));
      sheet.getRange('C' + (prodRow + 11)).setValue(toCaps(formData.dc4));
      sheet.getRange('C' + (prodRow + 12)).setValue(toCaps(formData.wx1));
      sheet.getRange('C' + (prodRow + 13)).setValue(toCaps(formData.wx2));
      sheet.getRange('C' + (prodRow + 14)).setValue(toCaps(formData.wx3));
      hideEmptyRows(3, prodRow, prodRow + 14);

      
      // Clear CAM rows 23/25 (labels in B, names in merged C:D)
      sheet.getRange(prodRow+15, 2, 7, 1).clearContent(); // Column B
      sheet.getRange(prodRow+15, 3, 7, 1).clearContent(); // Column C (merged with D)

      // Only write non-empty camera entries
      cams = formData.cameras || [];
      writeRow = prodRow+15;
      for (let i = 0; i < cams.length && writeRow <= writeRow+6; i++) {
        camName = cams[i];
        if (camName && camName.trim() !== '') {
          sheet.getRange(writeRow, 2).setValue(`CAM ${i + 1}`);
          sheet.getRange(writeRow, 3).setValue(toCaps(camName.trim()));
          writeRow++;
        }
      }
      break;

    case 'MBB':
    case 'WBB':

      if(isBroadcast) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('MBB Broadcast');
        sheet.getRange('C13').setValue(toCaps(formData.producer));
        sheet.getRange('C14').setValue(toCaps(formData.director));
        sheet.getRange('C15').setValue(toCaps(formData.ad));
        sheet.getRange('C16').setValue(toCaps(formData.td));
        sheet.getRange('C17').setValue(toCaps(formData.ap));
        sheet.getRange('C18').setValue(toCaps(formData.xpr));
        sheet.getRange('C19').setValue(toCaps(formData.bug));
        sheet.getRange('C20').setValue(toCaps(formData.dc1));
        sheet.getRange('C21').setValue(toCaps(formData.dc2));
        sheet.getRange('C22').setValue(toCaps(formData.cam3grip));
        sheet.getRange('C23').setValue(toCaps(formData.cam4grip));


        // Clear CAM rows 23/25 (labels in B, names in merged C:D)
        sheet.getRange(24, 2, 6, 1).clearContent(); // Column B
        sheet.getRange(24, 3, 6, 1).clearContent(); // Column C (merged with D)

        // Only write non-empty camera entries
        cams = formData.cameras || [];
        writeRow = 24;
        for (let i = 0; i < cams.length && writeRow <= 29; i++) {
          camName = cams[i];
          if (camName && camName.trim() !== '') {
            sheet.getRange(writeRow, 2).setValue(`CAM ${i + 1}`);
            sheet.getRange(writeRow, 3).setValue(toCaps(camName.trim()));
            writeRow++;
          }
        }
      } else {

        sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('MBB BigScreen');
        sheet.getRange('C13').setValue(toCaps(formData.producer));
        sheet.getRange('C14').setValue(toCaps(formData.director));
        sheet.getRange('C15').setValue(toCaps(formData.td));
        sheet.getRange('C16').setValue(toCaps(formData.showControl));
        sheet.getRange('C17').setValue(toCaps(formData.xpr));
        sheet.getRange('C18').setValue(toCaps(formData.wx1));
        sheet.getRange('C19').setValue(toCaps(formData.wx2));
        sheet.getRange('C20').setValue(toCaps(formData.slash));
        sheet.getRange('C21').setValue(toCaps(formData.dc1));
        sheet.getRange('C22').setValue(toCaps(formData.dc2));
        
        // Clear CAM rows 23/25 (labels in B, names in merged C:D)
        sheet.getRange(23, 2, 1, 1).clearContent(); // Column B
        sheet.getRange(23, 3, 1, 1).clearContent(); // Column C (merged with D)

        // Only write non-empty camera entries
        cams = formData.cameras || [];
        writeRow = 23;
        for (let i = 0; i < cams.length && writeRow <= 23; i++) {
          camName = cams[i];
          if (camName && camName.trim() !== '') {
            sheet.getRange(writeRow, 2).setValue(`CAM ${i + 1}`);
            sheet.getRange(writeRow, 3).setValue(toCaps(camName.trim()));
            writeRow++;
          }
        }
      }
      break;

    case 'SFB': 
      if(isBroadcast) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('SFB Broadcast');
        sheet.getRange('C13').setValue(toCaps(formData.producer));
        sheet.getRange('C' + (prodRow + 1)).setValue(toCaps(formData.director));
        sheet.getRange('C' + (prodRow + 2)).setValue(toCaps(formData.ad));
        sheet.getRange('C' + (prodRow + 3)).setValue(toCaps(formData.ap));
        sheet.getRange('C' + (prodRow + 4)).setValue(toCaps(formData.td));
        sheet.getRange('C' + (prodRow + 5)).setValue(toCaps(formData.bug));
        sheet.getRange('C' + (prodRow + 6)).setValue(toCaps(formData.xpr));
        sheet.getRange('C' + (prodRow + 7)).setValue(toCaps(formData.dc1));
        sheet.getRange('C' + (prodRow + 8)).setValue(toCaps(formData.dc2));
        sheet.getRange('C' + (prodRow + 9)).setValue(toCaps(formData.dc3));
        sheet.getRange('C' + (prodRow + 10)).setValue(toCaps(formData.toc));

        // Clear CAM rows 23/25 (labels in B, names in merged C:D)
        sheet.getRange(prodRow+11, 2, 7, 1).clearContent(); // Column B
        sheet.getRange(prodRow+11, 3, 7, 1).clearContent(); // Column C (merged with D)

        // Only write non-empty camera entries
        cams = formData.cameras || [];
        writeRow = prodRow+11;
        for (let i = 0; i < cams.length && writeRow <= prodRow+17; i++) {
          camName = cams[i];
          if (camName && camName.trim() !== '') {
            sheet.getRange(writeRow, 2).setValue(`CAM ${i + 1}`);
            sheet.getRange(writeRow, 3).setValue(toCaps(camName.trim()));
            writeRow++;
          }
        }
      } else {
        sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('SFB BigScreen');
        sheet.getRange('C13').setValue(toCaps(formData.producer));
        sheet.getRange('C14').setValue(toCaps(formData.td));
        sheet.getRange('C15').setValue(toCaps(formData.xpr));
        sheet.getRange('C16').setValue(toCaps(formData.dc1));
        sheet.getRange('C17').setValue(toCaps(formData.wx1));
        sheet.getRange('C18').setValue(toCaps(formData.wx2));

        // Clear CAM rows 23/25 (labels in B, names in merged C:D)
        sheet.getRange(19, 2, 1, 1).clearContent(); // Column B
        sheet.getRange(19, 3, 1, 1).clearContent(); // Column C (merged with D)

        // Only write non-empty camera entries
        cams = formData.cameras || [];
        writeRow = 19;
        for (let i = 0; i < cams.length && writeRow <= 19; i++) {
          camName = cams[i];
          if (camName && camName.trim() !== '') {
            sheet.getRange(writeRow, 2).setValue(`CAM ${i + 1}`);
            sheet.getRange(writeRow, 3).setValue(toCaps(camName.trim()));
            writeRow++;
          }
        }
      }
      break;

    case 'BSB':
      if(isBroadcast) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('BSB Broadcast');
        sheet.getRange('C13').setValue(toCaps(formData.producer));
        sheet.getRange('C14').setValue(toCaps(formData.director));
        sheet.getRange('C15').setValue(toCaps(formData.ad));
        sheet.getRange('C16').setValue(toCaps(formData.ap));
        sheet.getRange('C17').setValue(toCaps(formData.td));
        sheet.getRange('C18').setValue(toCaps(formData.bug));
        sheet.getRange('C19').setValue(toCaps(formData.xpr));
        sheet.getRange('C20').setValue(toCaps(formData.dc1));
        sheet.getRange('C21').setValue(toCaps(formData.dc2));
        sheet.getRange('C22').setValue(toCaps(formData.dc3));

        // Clear CAM rows 23/25 (labels in B, names in merged C:D)
        sheet.getRange(23, 2, 7, 1).clearContent(); // Column B
        sheet.getRange(23, 3, 7, 1).clearContent(); // Column C (merged with D)

        // Only write non-empty camera entries
        cams = formData.cameras || [];
        writeRow = 23;
        for (let i = 0; i < cams.length && writeRow <= 29; i++) {
          camName = cams[i];
          if (camName && camName.trim() !== '') {
            sheet.getRange(writeRow, 2).setValue(`CAM ${i + 1}`);
            sheet.getRange(writeRow, 3).setValue(toCaps(camName.trim()));
            writeRow++;
          }
        }
      } else {
        sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('BSB BigScreen');
        sheet.getRange('C13').setValue(toCaps(formData.producer));
        sheet.getRange('C14').setValue(toCaps(formData.td));
        sheet.getRange('C15').setValue(toCaps(formData.xpr));
        sheet.getRange('C16').setValue(toCaps(formData.dc1));
        sheet.getRange('C17').setValue(toCaps(formData.wx1));
        sheet.getRange('C18').setValue(toCaps(formData.wx2));

        // Clear CAM rows 23/25 (labels in B, names in merged C:D)
        sheet.getRange(19, 2, 1, 1).clearContent(); // Column B
        sheet.getRange(19, 3, 1, 1).clearContent(); // Column C (merged with D)

        // Only write non-empty camera entries
        cams = formData.cameras || [];
        writeRow = 19;
        for (let i = 0; i < cams.length && writeRow <= 19; i++) {
          camName = cams[i];
          if (camName && camName.trim() !== '') {
            sheet.getRange(writeRow, 2).setValue(`CAM ${i + 1}`);
            sheet.getRange(writeRow, 3).setValue(toCaps(camName.trim()));
            writeRow++;
          }
        }
      }
      break;

    case 'VB':
      if(isBroadcast) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('VB Broadcast');
        sheet.getRange('C13').setValue(toCaps(formData.producer));
        sheet.getRange('C14').setValue(toCaps(formData.director));
        sheet.getRange('C15').setValue(toCaps(formData.ad));
        sheet.getRange('C16').setValue(toCaps(formData.ap));
        sheet.getRange('C17').setValue(toCaps(formData.toc));
        sheet.getRange('C18').setValue(toCaps(formData.bug));
        sheet.getRange('C19').setValue(toCaps(formData.xpr));
        sheet.getRange('C20').setValue(toCaps(formData.dc1));
        sheet.getRange('C21').setValue(toCaps(formData.dc2));
        sheet.getRange('C22').setValue(toCaps(formData.cam5grip));
        sheet.getRange('C23').setValue(toCaps(formData.cam6grip));

        // Clear CAM rows 23/25 (labels in B, names in merged C:D)
        sheet.getRange(24, 2, 7, 1).clearContent(); // Column B
        sheet.getRange(24, 3, 7, 1).clearContent(); // Column C (merged with D)

        // Only write non-empty camera entries
        cams = formData.cameras || [];
        writeRow = 24;
        for (let i = 0; i < cams.length && writeRow <= 30; i++) {
          camName = cams[i];
          if (camName && camName.trim() !== '') {
            sheet.getRange(writeRow, 2).setValue(`CAM ${i + 1}`);
            sheet.getRange(writeRow, 3).setValue(toCaps(camName.trim()));
            writeRow++;
          }
        }
        //SOC no wx1, wx2
      } else {
        sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('VB BigScreen');
        sheet.getRange('C13').setValue(toCaps(formData.producer));
        sheet.getRange('C14').setValue(toCaps(formData.td));
        sheet.getRange('C15').setValue(toCaps(formData.xpr));
        sheet.getRange('C16').setValue(toCaps(formData.dc1));
        sheet.getRange('C17').setValue(toCaps(formData.wx1));
        sheet.getRange('C18').setValue(toCaps(formData.wx2));
        

        // Clear CAM rows 23/25 (labels in B, names in merged C:D)
        sheet.getRange(19, 2, 1, 1).clearContent(); // Column B
        sheet.getRange(19, 3, 1, 1).clearContent(); // Column C (merged with D)

        // Only write non-empty camera entries
        cams = formData.cameras || [];
        writeRow = 19;
        for (let i = 0; i < cams.length && writeRow <= 19; i++) {
          camName = cams[i];
          if (camName && camName.trim() !== '') {
            sheet.getRange(writeRow, 2).setValue(`CAM ${i + 1}`);
            sheet.getRange(writeRow, 3).setValue(toCaps(camName.trim()));
            writeRow++;
          }
        }
      }
      break;

    case 'SOC':
    //SOC no grips!
      Logger.log("Saving SOC Data");      
      if(isBroadcast) {
        Logger.log("Saving SOC Broadcast Data");      
        sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('SOC Broadcast');
        prodRow = findRow("PRODUCER");
        sheet.showRows(prodRow, prodRow + 10);
        sheet.getRange('C' + prodRow).setValue(toCaps(formData.producer));
        sheet.getRange('C' + (prodRow + 1)).setValue(toCaps(formData.director));
        sheet.getRange('C' + (prodRow + 2)).setValue(toCaps(formData.td));
        sheet.getRange('C' + (prodRow + 3)).setValue(toCaps(formData.ad));
        sheet.getRange('C' + (prodRow + 4)).setValue(toCaps(formData.ap));
        sheet.getRange('C' + (prodRow + 5)).setValue(toCaps(formData.xpr));
        sheet.getRange('C' + (prodRow + 6)).setValue(toCaps(formData.bug));
        sheet.getRange('C' + (prodRow + 7)).setValue(toCaps(formData.toc));
        sheet.getRange('C' + (prodRow + 8)).setValue(toCaps(formData.dc1));
        sheet.getRange('C' + (prodRow + 9)).setValue(toCaps(formData.dc2));
        hideEmptyRows(3, prodRow, prodRow + 10);



        // Clear CAM rows 23/25 (labels in B, names in merged C:D)
        sheet.getRange((prodRow + 10), 2, 7, 1).clearContent(); // Column B
        sheet.getRange((prodRow + 10), 3, 7, 1).clearContent(); // Column C (merged with D)

        // Only write non-empty camera entries
        cams = formData.cameras || [];
        writeRow = (prodRow + 10);
        for (let i = 0; i < cams.length && writeRow <= (writeRow + 7); i++) {
          camName = cams[i];
          if (camName && camName.trim() !== '') {
            sheet.getRange(writeRow, 2).setValue(`CAM ${i + 1}`);
            sheet.getRange(writeRow, 3).setValue(toCaps(camName.trim()));
            writeRow++;
          }
        }
        //SOC no wx1, wx2
      } else {
        Logger.log("Saving SOC BigScreen Data");      
        Logger.log("Saving SOC Data" + "td:" + formData.td + "xpr: " + formData.xpr + "dc1:" + formData.dc1);      
        sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('SOC BigScreen');
        prodRow = findRow("PRODUCER");
        sheet.getRange('C' + prodRow).setValue(toCaps(formData.producer));
        sheet.getRange('C' + (prodRow + 1)).setValue(toCaps(formData.td));
        sheet.getRange('C' + (prodRow + 2)).setValue(toCaps(formData.xpr));
        sheet.getRange('C' + (prodRow + 3)).setValue(toCaps(formData.dc1));
        }
      break;

    default:
      SpreadsheetApp.getUi().alert("Unknown sport: " + sport);
      break;
  }
}