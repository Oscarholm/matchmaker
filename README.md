##  Matchmaker
 This script was written by Oscar Holm of Numbery, reachable at oscar@numbery.se
Disclaimer: 
 - Please be aware that this solution was put into place as an emergency measure over a few days
 - It was never meant to be a long-term solution and therefore is not constructed with maintainability in mind
 - The purpose of each function is vaguely documented throughout, for full chart please see link below
 
Requirements:
In order to run importData successfully it needs to match the following documents from its root folder:
- 'SAP' : a document containing the invoices to be matched. Should have SAP in the title.
- 'Purchase-Orders' : a document containing the purchase orders. Should have Purchase-Orders in title
- 'Translation' : a document containing translation of store names. Should have 'Translation' in the title

MAKING A NEW VERSION
- Create a new sheet in an empty folder
- Name the first tab 'MAIN'
- Mark and copy all of the content of this script
- Go into Tools -> Script editor, replace function myFunction() {} by pasting
- At the top, change 'Untitled project' into desired title
- Press Debug, make sure it's set to onOpen, allow the permissions required
- Exit out of script editor tab, update the sheet by reloading the page
- In the top menu, 'Match invoice' should have appeared. 
- Test it by pressing Match invoice -> Import data, it will return that no files were found.
- Put a Purchase-order file, name needs to contain purchase order, in the same folder
- Put an invoice file, name needs to contain SAP, in the same folder
- Put the Translation file, name needs to contain Translation in the same folder
- Now run Import data again, to see if successful, press the 3 lines down by the tabs,
  and check so that there are now 6 tabs, 1 visible and 5 hidden.
- Everything should now be up and running!


