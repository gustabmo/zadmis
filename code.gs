function zadmisSearchToSend() {
  const labelToSend = GmailApp.getUserLabelByName('zadmis-to-send');
  const labelWasSent = GmailApp.getUserLabelByName('zadmis-was-sent-v1');

  const threadsToSend = labelToSend.getThreads(0,10);

  threadsToSend.forEach( (thread) => {
    msgsToSend = thread.getMessages();
    msgsToSend.forEach ( (message) => {
      var description = parseHtml(message.getBody());
      var cardName = trim(
        description.get
      );

      let nom="";
      let prenom="";
      let dob = "";
      let cardName = "";

      try {
        prenom = description.match(/^\s*Pr[eé]nom\s*:\s*(.*)$/mi)[1].trim();
      } catch (error) {
        prenom = '';
      }

      try {
        nom = description.match(/^\s*Nom\s*:\s*(.*)$/mi)[1].trim();
      } catch (error) {
        nom = '';
      }

      try {
        dob = description.match(/^\s*Date de naissance\s*:\s*(.*)$/mi)[1].trim();
      } catch (error) {
        dob = '';
      }

      if (cardName === "") {
        cardName = message.getSubject();
      }


      Logger.log ( "<<<"+message.getSubject()+">>>");
      Logger.log ( "<"+prenom+">" );
      Logger.log ( "<"+nom+">" );
      Logger.log ( "<"+dob+">" );
      Logger.log ( "<"+cardName+">" );

      if (1==2) UrlFetchApp.fetch(
        "https://api.trello.com/1/cards", 
        {
          method: "post",
          payload: {
            name: cardName,
            desc: description,
            idList: PRIVATE_idList,
            key: PRIVATE_APIkey,
            token: PRIVATE_APItoken
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
