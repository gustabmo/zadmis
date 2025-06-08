# zadmis

by guexel@gmail.com = gustavo.exel@ersge.ch

A Goole AppsScript project to create Trello cards from emails received over gmail, and to summarize the board's cards in a Google Sheet

This implementation names the card with the email subject and uses the plain text of the email body as it's description. Actually, it tries to build the card name from the description in a way that works in my specific use case (admission forms filled on Ecole Rudolf Steiner de Geneve ersge.ch web site) but most probably will fail in any other use case, and revert to the standard subject->name and body->description rule. You can most probably modify the manipulation (function composeCardName) to fit your case.

Setup:

1. from Trello you'll need to get this information :
  - the list id and board id where the cards will be written - easiest way is to open a Trello card in the list, click Share then Export JSON and search for "idlist" and "idboard"
  - the list id of a board list you don't want summarized (if there's none, you can just use "xxx" for this one)
  - the API key and API token - create a Trello power up and things should be straightforward from there
2. on google drive, create a google sheet that will hold a summary of the cards and write down it's id (it's the long string of letters in the URL)
3. on gmail, configure a filter that will put a label "zadmis-to-send" on the emails you want to go into Trello cards
4. on the [AppsScript Google site script.google.com](https://script.google.com) create a new project named "zadmis"
5. get [Code.gs](https://github.com/gustabmo/zadmis/blob/main/Code.gs) from the repository and copy/paste its contents into your appsscript project's Code.gs file; same thing for Summary.gs but you'll have to create this file before copying/pasting the contents
6. create a private.gs file (see [private-template.gs](https://github.com/gustabmo/zadmis/blob/main/private-template.gs) for a template) with the private info you got from Trello and Google on steps 1 and 2
7. on the [AppsScript Google site](https://script.google.com), configure the zadmisSearchToSend() function to be run periodically, I used every 15 minutes; same thing with writeSummaryToSheet(), I used every two hours

When the zadmisSearchToSend() script runs, it will find the emails with a zadmis-to-send label, upload each of them to Trello, and change the label to zadmis-sent-v1 (change the label means remove one and add the other). When the writeSummaryToSheet() runs it will, you guessed, write a summary of the cards to the google sheet file.

Once installed and configured, if you wish to post to Trello some old emails, you can just add the zadmis-to-send label to them and wait for the next iteration.

----
History:
- 2025-03-23 first working version
- 2025-03-29 Summaru.gs with writeSummaryToSheet() function
