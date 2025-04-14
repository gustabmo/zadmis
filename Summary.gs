
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

  if (error == "") {
    sheet.getRange(row.r,4)
      .setValue ( "---")
      .setFontWeight('normal')
    ;
    row.r++;
    row.r++;
    if (!summarizeLabels ( sheet, row, labels )) {
      error = "error listing labels";
    }
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
    JSON.parse ( cards.getContentText() )
      .sort(compareCardsByName)
      .forEach ( (card) => {
        sheet.getRange(row.r,1).setValue(new Date ( card.dateLastActivity ));
        stLabels = "";
        card.idLabels.forEach ( (idLabel) => 
          stLabels += getLabelName ( idLabel, labels ) + " " 
        );
        sheet.getRange(row.r,3).setValue ( stLabels.trim() );
        sheet.getRange(row.r,4)
          .setRichTextValue(
            SpreadsheetApp.newRichTextValue()
              .setText(card.name)
              .setLinkUrl(card.url)
              .setTextStyle(
                SpreadsheetApp.newTextStyle()
                  .setForegroundColor("black")
                  .setUnderline(false)
                  .build()
              )
              .build()
          )
          .setFontWeight('normal')
        ;
        row.r++;
      } )
    ;
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


function compareCardsByName ( a, b ) {
  if (a.name < b.name) return -1;
  if (a.name === b.name) return 0;
  return +1;
}


function compareLabels ( a, b ) {
  let aWords = a.name.split(" ");
  let bWords = b.name.split(" ");
  let aLast = aWords.length>=1 ? aWords[aWords.length-1] : "zzzz";
  let bLast = bWords.length>=1 ? bWords[bWords.length-1] : "zzzz";

  if (aLast < bLast) return +1;
  if (aLast === bLast) return 0;
  return -1;
}


function summarizeLabels ( sheet, row, labels ) {
  labels
    .sort(compareLabels)
    .forEach ( (label) => {
      sheet.getRange(row.r,3)
        .setRichTextValue(
          SpreadsheetApp.newRichTextValue()
            .setText(label.name)
            .setLinkUrl(
              "https://trello.com/b/WLDAl4MM/admissions?filter=label:"
              +encodeURI(label.name)
            )
            .build()
        )
      ;
      row.r++;
    })
  ;
  row.r++;
  return true;
}
