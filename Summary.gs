
function writeSummaryListOnSheet() {

  let error = "";
  let row = { r : 1 }; // must be object to be passed as reference
  let col;
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

    col = 1;
    for (let title of [
      "dernière modif"
      ,null
      ,"classe"
      ,"élève et date de naissance"
      ,"dossier"
      ,"entrée"
      ,"rdv péd"
      ,"stage de"
      ,"stage à"
      ,"ok péd"
      ,"EA"
      ,"ok fin"
      ,"commentaires"
      ,"emails"
      ,"téléphones"
      ,"alarme"
    ]) {
      sheet.getRange(row.r,col++).setValue(title);
    }
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


function getTextField ( line, header ) {
  line = line.trim();
  if ((line.length >= 2) && "*_".includes(line[0]) && "*_".includes(line[line.length-1])) {
    line = line.substr(1,line.length-2);
  }
  if (line.startsWith(header)) {
    return line.substr(header.length).trim();
  } else {
    return "";
  }
}


function ifLineEndsWith ( line, ender ) {
  line = line.trim();
  if ((line.length >= 2) && "*_".includes(line[0]) && "*_".includes(line[line.length-1])) {
    line = line.substr(1,line.length-2);
  }
  if (line.endsWith(ender)) {
    return line.substr(0,line.length-ender.length).trim();
  } else {
    return "";
  }
}


function filterPhone ( st ) {
  let result="";
  let ch;
  for ( ch of st ) {
    if (ch.match(/[\-+0-9]/g)) {
      result += ch;
    }
  }
  return result;
}


function textToDateIfPossible ( st ) {
  st = st.trim();

  if (st == "") return null;

  if (st.match(/^\d\d[.-/]\d\d[.-/]\d\d(\d\d|)$/)) {
    st = st.substring(6)+"-"+st.substring(3,5)+"-"+st.substring(0,2);
    if (st.length == 8) {
      st = "20"+st;
    }
  }

  if (st.match(/^\d\d\d\d\.\d\d\.\d\d$/)) {
    // yyyy.mm.dd is treated as local date, transform it to yyyy-mm-dd to be treated as UTC
    st = st.substring(0,4)+"-"+st.substring(5,7)+"-"+st.substring(8,10);
  }

  let d = new Date(st);

  if (d instanceof Date && !isNaN(d)) {
    return d;
  } else {
    return st;
  }
}


function processEmail ( st ) {
  let parsed;

  if (
    (parsed = /\[.*?\]\(mailto:(.*?) ".*"\)/.exec(st)) 
    &&
    (parsed.length >= 2)
  ) {
    return parsed[1];
  } else {
    return st;
  }
}


function processEntryDate ( st ) {
  if (st=="") return null;

  if (st=="lundi 18 Août 2025 12h00") st = "Rentrée 2025";
  if (st=="lundi 25 Août 2025 12h00") st = "Rentrée 2025";
  if (st=="lundi 1 Sep 2025 12h00") st = "Rentrée 2025";

  if (st=="lundi 17 Août 2026 12h00") st = "Rentrée 2026";

  st = st.replace ( "Rentrée Août ", "Rentrée ");

  st = st.replace ( " 12h00", "" );

  st = st.replace ( "samedi ", "" );
  st = st.replace ( "dimanche ", "" );
  st = st.replace ( "lundi ", "" );
  st = st.replace ( "mardi ", "" );
  st = st.replace ( "mercredi ", "" );
  st = st.replace ( "jeudi ", "" );
  st = st.replace ( "vendredi ", "" );

  st = st.replace ( " Fév ", " Feb " );
  st = st.replace ( " Avr ", " Apr " );
  st = st.replace ( " Mai ", " May " );
  st = st.replace ( " Juin ", " Jun " );
  st = st.replace ( " Juil ", " Jul " );
  st = st.replace ( " Août ", " Aug " );

  return st;
}


function summarizeOneCard ( card, sheet, row, labels ) {
  let col;

  sheet.getRange(row.r,1).setValue(new Date ( card.dateLastActivity ));

  let stLabels = "";
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

  let stEmails = "";
  let stPhones = "";
  let lastIndicatif = "";
  let dateDossier = null;
  let dateEP = "";
  let dateEntree = "";
  let okPedagogique = "";
  let dateEA = "";
  let stageDe = "";
  let stageA = "";
  let okFinancier = "";
  let commentaires = "";
  let temp;
  card.desc.split ( "\n" ).forEach ( (line) => {
    if (temp = processEmail(getTextField(line,"Email :"))) stEmails += (stEmails==""?"":", ") + temp;
    if (temp = getTextField(line,"Indicatif :")) lastIndicatif = filterPhone(temp);
    if (temp = getTextField(line,"Téléphone mobile :")) stPhones = (stPhones + "  "+lastIndicatif+temp).trim();
    if (temp = textToDateIfPossible(getTextField(line,"date-EP :"))) dateEP = temp;
    if (temp = textToDateIfPossible(getTextField(line,"date-dossier :"))) dateDossier = temp;
    if (temp = textToDateIfPossible(getTextField(line,"reçu le :"))) dateDossier = temp;
    if (temp = textToDateIfPossible(ifLineEndsWith(line,"Ecole Rudolf Steiner <info@ersge.ch>"))) dateDossier = temp;
    if (temp = textToDateIfPossible(ifLineEndsWith(line,"Ecole Rudolf Steiner [info@ersge.ch](mailto:info@ersge.ch \"‌\")"))) dateDossier = temp;
    if (temp = textToDateIfPossible(getTextField(line,"ok-pedagogique :"))) okPedagogique = temp;
    if (temp = textToDateIfPossible(getTextField(line,"ok-financier :"))) okFinancier = temp;
    if (temp = textToDateIfPossible(getTextField(line,"date-EA :"))) dateEA = temp;
    if (temp = textToDateIfPossible(getTextField(line,"stage-De :"))) stageDe = temp;
    if (temp = textToDateIfPossible(getTextField(line,"stage-A :"))) stageA = temp;
    if (temp = processEntryDate(getTextField(line,"Date d'entrée souhaitée :"))) dateEntree = temp;
    if (temp = getTextField(line,"commentaires :")) commentaires = temp;
  } )

  col = 5;
  for (let field of [
    dateDossier, dateEntree, dateEP, stageDe, stageA, okPedagogique, dateEA, okFinancier
  ]) {
    sheet.getRange(row.r,col++)
      .setValue ( field )
      .setNumberFormat("DD.MM.YYYY")
      .setHorizontalAlignment('center')
    ;
  }

  sheet.getRange(row.r,col++).setValue ( commentaires );
  sheet.getRange(row.r,col++).setValue ( stEmails );
  sheet.getRange(row.r,col++).setValue ( stPhones );

  if (card.due) {
    sheet.getRange(row.r,col++)
      .setValue ( new Date ( card.due ) )
      .setNumberFormat("DD.MM.YYYY")
      .setHorizontalAlignment('center')
    ;
  }

  row.r++;
}


function summarizeCards ( sheet, row, labels, listId ) {
  let result = true;

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
      .forEach ( (card) => summarizeOneCard ( card, sheet, row, labels ) )
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
