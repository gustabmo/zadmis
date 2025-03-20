# zadmis

by guexel@gmail.com = gustavo.exel@ersge.ch

A project to create Trello cards from emails received over gmail

The implementation that is here works on a particular use case: for admission forms filled on Ecole Rudolf Steiner de Geneve web site. It should be easy to adapt to your particular use case.

Setup:

1. from Trello you'll need to get this information :
  - the list id - easiest way is to open a Trello card in this list, click Share then Export JSON and search for "idlist"
  - the API key and API token - create a Trello power up and things should be straightforward from there
2. on gmail, configure a filter that will put a label "zadmis-to-send" on the emails you want to go into Trello cards; if you use a different label, you can easily modify the code
3. on script.google.com, create a new project named "zadmis"
4. get code.gs from the repository and upload (or copy/paste) into your appsscript project    
5. create a private.gs file (see private-template.gs for a template) with the private info you got from Trello
6. configure the zadmisSearchToSend() function to be run periodically (I used every 15 minutes)
