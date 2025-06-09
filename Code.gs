function zadmisSearchToSend() {
  const labelToSend = GmailApp.getUserLabelByName('zadmis-to-send');
  const labelWasSent = getOrCreateGmailLabel('zadmis-sent-v1');

  const threadsToSend = labelToSend==null ? [] : labelToSend.getThreads(0,10);

  let labelsTrello = null;

  threadsToSend.forEach( (thread) => {
    let threadOk = true;

    msgsToSend = thread.getMessages();
    msgsToSend.forEach ( (message) => {
      const description = 
        "date-EP :\n"
        +"stage-De :\n"
        +"stage-A :\n"
        +"ok-pedagogique :\n"
        +"date-EA :\n"
        +"ok-financier :\n"
        +"\n"
        +"date-dossier : " + message.getDate().toISOString() + "\n"
        +"\n"
        + message.getPlainBody()
      ;
      const info = getAllInfo ( description );
      const cardName = composeCardName ( message, info.prenom, info.nom, info.dobst );
      const year1stclass = getYear1stclass ( info.dobdate );

      if (labelsTrello == null) {
        labelsTrello = getBoardLabels();
      } 

      const cardLabelId = getCardLabel ( labelsTrello, year1stclass );

      const newCard = createNewCard ( cardName, description );

      if (newCard == null) {
        threadOk = false;
      } else {
        addLabelToCard ( newCard.id, cardLabelId );
      }
    } ); // msgsToSend.forEach

    if (threadOk) {
      labelWasSent.addToThread(thread);
      labelToSend.removeFromThread(thread);
    }
  } ); // threadsToSend.forEach
} // zadmisSearchToSend()



function getInfo ( description, header ) {
  match = description.match ( new RegExp("^\\s*"+header+"\\s*:\\s*(.*)$","mi") );
  if ((match != null) && (match.length >= 2)) {
    return match[1]
      .trim()
      .replace(/^\*/,"")
      .replace(/\*$/,"")
      .trim()
    ;
  } else {
    return "";
  }
} // getInfo()



function getAllInfo ( description ) {
  const prenom = getInfo(description,"Pr[eé]nom");
  const nom = getInfo(description,"Nom");
  let dobst = getInfo(description,"Date de naissance");
  const dobdate = getDateFromFormSt(dobst);

  if (dobdate != null) {
    dobst = dobdate.toISOString().substring(0,10);
  }

  return { prenom:prenom, nom:nom, dobst:dobst, dobdate:dobdate };
} // getAllInfo()



function getDateFromFormSt ( st ) {
  const tokens = st.match ( /\s*([^\s]*)\s*/g );
  let day=0;
  let month=0;
  let year=0;
  let newmonth;
  let tok;
  let v;
  let error = false;

  for (i in tokens) if (!error) {
    tok = tokens[i]
      .trim()
      .normalize('NFD').replace(/[\u0300-\u036f]/g,'')  // remove accents
      .toUpperCase()
    ;

    newmonth = 0;
    if (
      (tok === "LUNDI") || (tok === "MARDI") || (tok === "MERCREDI") || (tok === "JEUDI") 
      || 
      (tok === "VENDREDI") || (tok === "SAMEDI") || (tok === "DIMANCHE") 
      ||
      (tok === "MONDAY") || (tok === "TUESDAY") || (tok === "WEDNESDAY") || (tok === "THURSDAY") 
      || 
      (tok === "FRIDAY") || (tok === "SATURDAY") || (tok === "SUNDAY") 
      || 
      (tok === "MON") || (tok === "TUE") || (tok === "WED") || (tok === "THU") 
      || 
      (tok === "FRI") || (tok === "SAT") || (tok === "SUN") 
      ||
      (tok === "12H00") || (tok === "")
    ) {
      // ignore
    } else if ((tok == "JAN") || (tok == "JANVIER") || (tok == "JANUARY")) { newmonth = 1
    } else if ((tok == "FEV") || (tok == "FEVRIER") || (tok == "FEB") || (tok == "FEBRUARY")) { newmonth = 2
    } else if ((tok == "MAR") || (tok == "MARS") || (tok == "MARCH")) { newmonth = 3
    } else if ((tok == "AVR") || (tok == "AVRIL") || (tok == "APR") || (tok == "APRIL")) { newmonth = 4
    } else if ((tok == "MAI") || (tok == "MAY")) { newmonth = 5
    } else if ((tok == "JUIN") || (tok == "JUN") || (tok == "JUNE")) { newmonth = 6
    } else if ((tok == "JUIL") || (tok == "JUILLET") || (tok == "JUL") || (tok == "JULY")) { newmonth = 7
    } else if ((tok == "AOU") || (tok == "AOUT") || (tok == "AUG") || (tok == "AUGUST")) { newmonth = 8
    } else if ((tok == "SEP") || (tok == "SEPTEMBRE") || (tok == "SEPTEMBER")) { newmonth = 9
    } else if ((tok == "OCT") || (tok == "OCTOBRE") || (tok == "OCTOBER")) { newmonth = 10
    } else if ((tok == "NOV") || (tok == "NOVEMBRE") || (tok == "NOVEMBER")) { newmonth = 11
    } else if ((tok == "DEC") || (tok == "DECEMBRE") || (tok == "DECEMBER")) { newmonth = 12
    } else {
      v = Number(tok);
      if ((v != null) && (v>0) && (v==Math.floor(v))) {
        if ((v>=1) && (v<=31)) {
          if (day == 0) {
            day = v;
          } else {
            error = true;
          }
        } else if ((v>=1800) && (v<=4000)) {
          if (year == 0) {
            year = v;
          } else {
            error = true;
          }
        } else {
          error = true;
        }
      } else {
        error = true;
      }
    }

    if ((!error) && (newmonth != 0)) {
      if (month == 0) {
        month = newmonth;
      } else {
        error = true;
      }
    }
  }

  if (error) {
    return null
  } else {
    try {
      return new Date ( Date.UTC ( year,month-1,day ) );
    } catch (e) {
      return null;
    }
  }
} // getDateFromFormSt()



function getBoardLabels() {
  const response = UrlFetchApp.fetch(
    "https://api.trello.com/1/boards/"+PRIVATE_Trello_idBoard
    +"/labels?key="+PRIVATE_Trello_APIkey
    +"&token="+PRIVATE_Trello_APItoken,
    {
      muteHttpExceptions: true,
    }
  );
  if (response.getResponseCode() == 200) {
    return JSON.parse ( response.getContentText() );
  } else {
    Logger.log ( "error reading list of labels" );
    return null;
  }
} // getBoardLabels()



function composeCardName ( message, prenom, nom, dobst ) {
  if ((prenom+nom === "") && (!(dobst instanceof Date) || (dobst < new Date(1970,1,1)))) {
    return message.getSubject();
  } else {
    return (nom.toUpperCase()+" "+prenom+" "+dobst).trim();
  }
} // composeCardName()



function getYear1stclass ( dobdate ) {
  let year1stclass = null;
  if (dobdate != null) {
    year1stclass = dobdate.getUTCFullYear() + 6;
    if ((dobdate.getUTCMonth()+1) >= 8) {
      year1stclass++;
    }
    if (year1stclass < 1980) {
      year1stclass = null;
    }
  }
  return year1stclass;
} // getYear1stclass()



function getCardLabel ( labels, year1stclass ) {
  // do I have a list of labels?
  if (labels == null) {
    return null;
  }

  if ((year1stclass == null) || (year1stclass < 1980)) {
    return null;
  }

  // search among existing labels
  for ( i=0; i<labels.length; i++ ) {
    if (labels[i].name.match(year1stclass)) {
      return labels[i].id;
    }   
  }

  return createCardLabel ( labels, "class xx "+year1stclass );
} // getCardLabel()



function createCardLabel ( labels, nameNewLabel ) {
  const respCreateLabel = UrlFetchApp.fetch(
    "https://api.trello.com/1/labels/",
    {
      muteHttpExceptions: true,
      method: "post",
      payload: {
        name : nameNewLabel,
        color : "yellow",
        idBoard : PRIVATE_Trello_idBoard,
        key: PRIVATE_Trello_APIkey,
        token: PRIVATE_Trello_APItoken
      }
    }
  );

  if (respCreateLabel.getResponseCode() == 200) {
    const newLabel = JSON.parse ( respCreateLabel.getContentText() );
    labels.push ( newLabel );
    Logger.log ( "created label "+nameNewLabel+" "+newLabel.id );
    return newLabel.id;
  } else {
    Logger.log ( "error creating label "+nameNewLabel 
      +" error:"+respCreateLabel.getResponseCode()
      +" "+respCreateLabel.getContentText()
    );
    return null;
  }
} // createCardLabel()



function addLabelToCard ( cardId, labelId ) {
  let result = false;

  if (labelId != null) {
    const respAddLabelToCard = UrlFetchApp.fetch (
      "https://api.trello.com/1/cards/"+cardId+"/idLabels", 
      {
        method: "post",
        muteHttpExceptions: true,
        payload: {
          value : labelId,
          key: PRIVATE_Trello_APIkey,
          token: PRIVATE_Trello_APItoken
        }
      }
    );

    if (respAddLabelToCard.getResponseCode() == 200) {
      result = true;
    } else {
      Logger.log ( "error adding label "+labelId
        +" to card "+cardId
        +" error "+respAddLabelToCard.getResponseCode()
        +" "+respAddLabelToCard.getContentText() 
      );
    }
  }

  return result;
} // addLabelToCard()



function createNewCard ( cardName, description ) {
  const respAddCard = UrlFetchApp.fetch(
    "https://api.trello.com/1/cards", 
    {
      method: "post",
      muteHttpExceptions: true,
      payload: {
        name: cardName,
        desc: description,
        idList: PRIVATE_Trello_idList,
        key: PRIVATE_Trello_APIkey,
        token: PRIVATE_Trello_APItoken
      }
    }
  );
  
  if (respAddCard.getResponseCode() != 200) {
    Logger.log ( "Error creating card "+cardName 
      +" error:"+respAddCard.getResponseCode()
      +" "+respAddCard.getContentText()
    );
    return null;
  } else {
    return JSON.parse ( respAddCard.getContentText() );
  }
} // createNewCard()



function getOrCreateGmailLabel(name) {
  let label = GmailApp.getUserLabelByName(name);
  if (label == null) {
    label = GmailApp.createLabel(name);
  }
  return label;
} // getOrCreateGmailLabel()

