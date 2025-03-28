# zadmis

by guexel@gmail.com = gustavo.exel@ersge.ch

A Goole AppsScript project to create Trello cards from emails received over gmail

This implementation names the card with the email subject and uses the plain text of the email body as it's description. Actually, it tries some manipulation that will work in my specific use case (admission forms filled on Ecole Rudolf Steiner de Geneve web site) but most probably this manipulation will fail in any other use case, which will revert to the standard subject->name and body->description rule. You can most probably modify the manipulation to fit your case.

Setup:

1. from Trello you'll need to get this information :
  - the list id and board id - easiest way is to open a Trello card in the list, click Share then Export JSON and search for "idlist" and "idboard"
  - the API key and API token - create a Trello power up and things should be straightforward from there
2. on gmail, configure a filter that will put a label "zadmis-to-send" on the emails you want to go into Trello cards
3. on the [AppsScript Google site script.google.com](https://script.google.com) create a new project named "zadmis"
4. get [Code.gs](https://github.com/gustabmo/zadmis/blob/main/Code.gs) from the repository and copy/paste its contents into your appsscript project's Code.gs file.
5. create a private.gs file (see [private-template.gs](https://github.com/gustabmo/zadmis/blob/main/private-template.gs) for a template) with the private info you got from Trello on the first step
6. on the [AppsScript Google site](https://script.google.com), configure the zadmisSearchToSend() function to be run periodically (I used every 15 minutes)

When the script runs, it will find the emails with a zadmis-to-send label, upload each of them to Trello, and change the label to zadmis-sent-v1 (change the label means remove one and add the other).

Once installed and configured, if you wish to post to Trello some old emails, you can just add the zadmis-to-send label to them and wait for the next iteration.

----
History:
- 2025-03-23 first working version
