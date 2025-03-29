function writeSummaryListOnSheet() {
  let error = "";
  let row = { r : 1 }; // must be object to be passed as reference
  let response = null;
  let sheet = SpreadsheetApp.openById(PRIVATE_GSheets_Summary).getActiveSheet();
  if (sheet == null) {
    error = "google sheet not found";
  }

  if (error == "") {
    sheet.clear();

    sheet.getRange(row.r,1).setValue(new Date(Date.now()));
    sheet.getRange(row.r,3).setValue("dernière mise à jour");
    row.r++;
    row.r++;

    sheet.getRange(row.r,1).setValue("dernière modif");
    sheet.getRange(row.r,3).setValue("classe");
    sheet.getRange(row.r,4).setValue("élève et date de naissance");
    row.r++;
    row.r++;
  }

  labels = getBoardLabels();

  if (error == "") {
    response = UrlFetchApp.fetch(
      "https://api.trello.com/1/boards/"+PRIVATE_Trello_idBoard
      +"/lists?key="+PRIVATE_Trello_APIkey
      +"&token="+PRIVATE_Trello_APItoken,
      {
        muteHttpExceptions: true,
      }
    );
    if (response == null) {
      error = "get boards failed";
    }
  }

  if (error != "") {
    // nothing
  } else if (response.getResponseCode() != 200) {
    error = response.getContentText();
  } else {
    JSON.parse ( response.getContentText() ).forEach ( (list) => {
      if ((error == "") && (list.id != PRIVATE_GSheets_Summary_List_Dont_Show)) {
        sheet.getRange(row.r,4)
          .setValue(list.name)
          .setFontWeight('bold')
        ;
        row.r++;
        if (!summarizeCards ( sheet, row, labels, list.id )) {
          error = "error reading list "+list.name;
        }
      }
    } );
  }

  if (error != "") {
    Logger.log ( "error: "+error );
  }
  if (sheet != null) {
    sheet.getRange(row.r,4)
      .setValue ( error==""? "---" : "error: "+error )
      .setFontWeight('normal')
    ;
    row.r++;
  }  
}


function summarizeCards ( sheet, row, labels, listId ) {
  let result = true;
  let stLabels;

  const cards = UrlFetchApp.fetch(
    "https://api.trello.com/1/lists/"+listId
    +"/cards?key="+PRIVATE_Trello_APIkey
    +"&token="+PRIVATE_Trello_APItoken,
    {
      muteHttpExceptions: true,
    }
  );
  if ((cards == null) || (cards.getResponseCode() != 200)) {
    result = false;
  } else {
    JSON.parse ( cards.getContentText() ).forEach ( (card) => {
      sheet.getRange(row.r,1).setValue(new Date ( card.dateLastActivity ));
      stLabels = "";
      card.idLabels.forEach ( (idLabel) => 
        stLabels += getLabelName ( idLabel, labels ) + " " 
      );
      sheet.getRange(row.r,3).setValue ( stLabels.trim() );
      sheet.getRange(row.r,4)
        .setValue(card.name)
        .setFontWeight('normal')
      ;
      row.r++;
    } );
    row.r++;
  }

  return result;
}


function getLabelName ( labelId, labels ) {
  const label = labels.find ( (label) => label.id==labelId );
  if (label != null) {
    return label.name;
  } else {
    return "";
  }
}