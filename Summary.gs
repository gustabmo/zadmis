
function writeSummaryListOnSheet() {

  let error = "";
  let col;
  let response = null;
  let sheet = SpreadsheetApp.openById(PRIVATE_GSheets_Summary).getActiveSheet();
  let newValues = []; // all the values of all the cells
  if (sheet == null) {
    error = "google sheet not found";
  }

  if (error == "") {
    sheet.clear();
  }

  if (error == "") {
    newValues.push ( [
      new Date(Date.now()),
      null,
      "'<=== dernière mise à jour"
    ] );

    newValues.push ( [] );
  }

  if (error == "") {
    newValues.push ( [
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
      ,null
      ,"situation"
      ,"connu l'école"
    ] );

    newValues.push ( [] );
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
    listNumber = 0;
    JSON.parse ( response.getContentText() ).forEach ( (list) => {
      if ((error == "") && (list.id != PRIVATE_GSheets_Summary_List_Dont_Show)) {
        listNumber++;
        newValues.push ( [
          null,
          null,
          null,
          `'=== ${list.name} ===`
        ] );
        if (!summarizeCards ( 
          newValues,
          labels, 
          list.id, 
          `${String(listNumber).padStart(2,'0')} ${list.name}` 
        )) {
          error = "error reading list "+list.name;
        }
        newValues.push([]);
      }
    } );
  }

  if (error == "") {
    newValues.push ( [null,null,null,"---"] );
    newValues.push ( [] );
    if (!summarizeLabels ( newValues, labels )) {
      error = "error listing labels";
    }
  }

  if (error == "") {
    newValues.push ( [null,null,null,"---"] );
  }

  if (error == "") {
    sheet.getRange ( 1, 1, newValues.length, standardizeLineLengths(newValues) )
      .setValues ( newValues )
    ;
    sheet.getRange("A:A").setNumberFormat("DD.MM.YYYY hh:mm").setHorizontalAlignment('center');
    sheet.getRange("E:L").setNumberFormat("DD.MM.YYYY").setHorizontalAlignment('center');
  }

  if (error != "") {
    Logger.log ( "error: "+error );
    if (sheet != null) {
      sheet.getRange(1,4)
        .setValue ( "error: "+error )
        .setFontWeight('bold')
      ;
    } 
  }
}


function standardizeLineLengths ( newValues ) {
  let max = 0;
  let line;
  for (line of newValues) {
    max = Math.max ( max, line.length );
  }  
  for (line of newValues) {
    while (line.length < max) {
      line.push ( null );
    }
  }  
  return max;
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

  stDateSpecifique = "Date spécifique";
  if (st.startsWith(stDateSpecifique)) {
    let bakst = st;
    st = st.slice ( stDateSpecifique.length );
    console.log ( `@@ ${bakst} -> ${st}` );
  } 

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


function summarizeOneCard ( card, newValues, labels, situation ) {
  let line = [];

  line.push ( new Date ( card.dateLastActivity ) );

  line.push ( card.url );

  let stLabels = "";
  card.idLabels.forEach ( (idLabel) => 
    stLabels += getLabelName ( idLabel, labels ) + " " 
  );
  line.push ( stLabels.trim() );

  line.push ( card.name );

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
  let connulecole = "";
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
    if (
      (temp = getTextField(line,"Comment avez-vous connu l'école ? (plusieurs choix possible) "))
      ||
      (temp = getTextField(line,"Comment avez-vous connu l'école ?"))
      ||
      (temp = getTextField(line,"*Comment avez-vous connu l'école ?*"))
      ||
      (temp = getTextField(line,"_Comment avez-vous connu l'école ?_"))
    ) connulecole = temp;
  } )

  for (let field of [
    dateDossier, dateEntree, dateEP, stageDe, stageA, okPedagogique, dateEA, okFinancier, 
    commentaires, stEmails, stPhones, 
    null,
    situation, connulecole
  ]) {
    line.push ( field );
  }

  newValues.push ( line );
}


function summarizeCards ( newValues, labels, listId, situation ) {
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
      .forEach ( (card) => summarizeOneCard ( card, newValues, labels, situation ) )
    ;
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


function summarizeLabels ( newValues, labels ) {
  labels
    .sort(compareLabels)
    .forEach ( (label) => {
      newValues.push ( [
        null, null,
        label.name,
        `https://trello.com/b/WLDAl4MM/admissions?filter=label:${encodeURI(label.name)}`
      ] );
    })
  ;
  newValues.push ( [] );
  return true;
}
