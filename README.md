# zadmis - still in development as of 2025-03-20

by guexel@gmail.com = gustavo.exel@ersge.ch

A project to create Trello cards from emails received over gmail

This implementation names the card with the email subject and uses the plain text of the email body as it's description. Actually, it tries some manipulation that will work in my specific use case (admission forms filled on Ecole Rudolf Steiner de Geneve web site) but most probably this manipulation will fail in any other use case, which will revert to the standard subject->name and body->description rule. You can most probably modify the manipulation to fit your case.

Setup:

1. from Trello you'll need to get this information :
  - the list id - easiest way is to open a Trello card in the list, click Share then Export JSON and search for "idlist"
  - the API key and API token - create a Trello power up and things should be straightforward from there
2. on gmail, configure a filter that will put a label "zadmis-to-send" on the emails you want to go into Trello cards
3. on the [AppsScript Google site script.google.com](https://script.google.com) create a new project named "zadmis"
4. get [code.gs](https://github.com/gustabmo/zadmis/blob/main/code.gs) from the repository and copy/paste its contents into your appsscript project's Code.gs file.
5. create a private.gs file (see [private-template.gs](https://github.com/gustabmo/zadmis/blob/main/private-template.gs) for a template) with the private info you got from Trello on the first step
6. on the [AppsScript Google site](https://script.google.com), configure the zadmisSearchToSend() function to be run periodically (I used every 15 minutes)

When the script runs, it will find the emails with a zadmis-to-send label, upload each of them to Trello, and change the label to zadmis-sent-v1 (change the label means remove one and add the other).

Once installed and configured, you can go back to previous emails that you'd like to transform into Trello cards, add the zadmis-to-send label and wait for the next iteration.
