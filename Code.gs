function zadmisSearchToSend() {
  const labelToSend = GmailApp.getUserLabelByName('zadmis-to-send');
  const labelWasSent = GmailApp.getUserLabelByName('zadmis-sent-v1');

  const threadsToSend = labelToSend.getThreads(0,10);

  threadsToSend.forEach( (thread) => {
    msgsToSend = thread.getMessages();
    msgsToSend.forEach ( (message) => {
      var description = parseHtml(message.getBody());
      var cardName;

      var prenom = getInfo(description,"Pr[eé]nom");
      var nom = getInfo(description,"Nom");
      var dobst = getInfo(description,"Date de naissance");

      if (dobst != "") {
        var dob = getDate(dob); 
      }

      if (prenom+nom+dob === "") {
        cardName = message.getSubject();
      } else {
        cardName = nom.toUpperCase()+" "+prenom+" "+dob;
      }

      Logger.log ( "<<<  BODY >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>");
      // Logger.log ( message.getBody());
      Logger.log ( "<<<  DESCRIP >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>");
      // Logger.log ( description);
      Logger.log ( "<<<"+message.getSubject()+">>>");
      Logger.log ( "<"+prenom+">" );
      Logger.log ( "<"+nom+">" );
      Logger.log ( "dob st <"+dob+">" );
      Logger.log ( "dob date <"+new Date(dob)+">" );
      Logger.log ( "<"+cardName+">" );

      if (1==2) UrlFetchApp.fetch(
        "https://api.trello.com/1/cards", 
        {
          method: "post",
          payload: {
            name: cardName,
            desc: description,
            idList: PRIVATE_Trello_idList,
            key: PRIVATE_Trello_APIkey,
            token: PRIVATE_Trello_APItoken
          }
        }
      );
    } );
  } );
}

function parseHtml(html) {
  // source: https://stackoverflow.com/questions/57117272/is-there-a-function-or-example-for-converting-html-string-to-plaintext-without-h
  // by Sizerth
  // very slow as all appsscript's that use google calls, but works very well

  var draftMsg = GmailApp.createDraft('', 'zadmis To be deleted', '', { htmlBody: html });
  var plainText = draftMsg.getMessage().getPlainBody();
  draftMsg.deleteDraft();
  return plainText;
}

function getInfo ( description, header ) {
  var info;
  match = description.match ( new RegExp("^\\s*"+header+"\\s*:\\s*(.*)$","mi") );
  if ((match != null) && (match.length >= 2)) {
    info = match[1].trim();
    return info.trim().replace(/^\*/,"").replace(/\*$/,"").trim();
  } else {
    return "";
  }
}