# Outlook-Tools-Addin
So far this outlook addin can:
- list all outlook folder objects in a treeview
- search for items in the selected folder node of the treeview and display them in a list
- sort string and datetime columns of the list by clicking on column header

- select the selected folder node in the outlook explorer
- select the selected email-item in the outlook explorer and show it in the outlook preview
- open the selected email item by doubleclicking on a row of the listview

- shift emails into folder by drag and drop on tree node
- display filtered folder contents by doubleclick on a tree node 

- display the attachments of all emails in the resultset in a separate window
- after double-click on a listview row: save an attachment in the windows temp folder, unblock it and open the default application

The application searches by looping through all email items in the folder and subfolders. This search method is slower than the standard OL search, but it also works when the OL search does not function for some reason. This addin could be customized to perform further actions on the selected folders/email items. In particular, it could be useful for customized processing of email attachments. I am open to suggestions what it should be able to do in the future. 

![search form](https://github.com/Jens-Kluge/Outlook-Tools-Addin/blob/master/ol%20searchfolder.png)

![search form and attachment form](https://github.com/Jens-Kluge/Outlook-Tools-Addin/blob/master/OT%20Addin%20Capture.GIF)
