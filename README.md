

> [!CAUTION]
> This ReadMe and repo is still a work-in-process. Its not really ready to look at yet
> <BR>
> You're welcome to browse, but I'm still experimenting with stuff


Lock-Unlock-Sheet
=================

> [!NOTE]
> This is my first GitHub open source repo! **Please :star: if you think it will be useful for people.**
> <BR>

What this repo contains
---
This repo demonstrates a sample technique to programmatically lock/unlock cells in a Google Workspace Sheet.
<BR>
The sample enables you to:
- Create locked forms in Google Sheets with editable form fields
- Capture audit history as people edit the form fields
- Unlock a single cell for editing, and lock the cell after they have finished
- Programatically lock/unlock a cell based on business rules, such as "a user can only edit empty cells" or "person x can override text in any cell" 

The sample is written entirely in Google Apps Script, distributed as open source (under the Apache2 license), you can use + change it as you wish. 
<BR>
<BR>
In this document you'll find a teaser video, try-before-you-buy sample, description of how-it-works, setup and user guide. 

### Here is the teaser video that shows what it does:
<BR>
<a href="http://www.youtube.com/watch?feature=player_embedded&v=YOUTUBE_VIDEO_ID_HERE
" target="_blank"><img src="http://img.youtube.com/vi/YOUTUBE_VIDEO_ID_HERE/0.jpg" 
alt="IMAGE ALT TEXT HERE" width="240" height="180" border="10" /></a>

### Here is the try-before-you-buy sample:
<table>
  <tr>
    <td>TODO WAITING FOR OAUTH VERIFICATION</td>
    <td><-- Click this button to open the lock/unlock sample spreadsheet shown in the teaser video above. Your browser will need to be logged into a Google account</td>
  </tr>
</table>

How it works
---
> [!IMPORTANT]
> If you are the owner of the sheet, you **always** have edit rights to every cell.
> <BR> To see it working, you'll need to use a Google account that has edit rights, but is **not the owner** of the Spreadsheet
> 
The Lock-Unlock-Sheet technique uses two objects running with different permissions:
* **Google Workspace Spreadsheet:** Users interact with a Spreadsheet. The Spreadsheet is shared with Edit permissions, but each Sheet within the Spreadsheet is locked to prevent editing. Because the Spreadsheet itself has edit permissions, users can click an Edit button, but they can't edit cells or change any content in the Sheet until it is unlocked
* **WebApp:** The WebApp runs with the Spreadsheet owner's permission (it has full rights to change any content in the sheet, or unlock cells for editing). When a user clicks an "Edit" button in the locked Spreadsheet, the spreadsheet requests the WebApp to to either set a cell's text (for forms) or unlock a cell for editing (for free editing). This diagram shows the flow for both cases:

```mermaid
%%{init: {"mirrorActors": false} }%%
sequenceDiagram
    actor Spreadsheet
    participant WebApp
    Spreadsheet->>+WebApp: Plz change cell A1 to "Dog"
    WebApp->>-Spreadsheet: OK. A1 changed to "Dog"
    Spreadsheet-->>+WebApp: Plz unlock B2 for Alex to edit
    WebApp-->>-Spreadsheet: OK. B2 Cell unlocked for Alex
```

The WebApp unlocks a cell for the user using a two layers of Spreadsheet protection. The lower protection layer is the sheet-lock. The webapp unlocks ranges within the sheet-lock, then adds an upper-layer range protection to the unlocked cell to ensure only the user requesting the edit is authorized to edit the cell. This means that when a user requests to edit a cell, it is unlocked only for them. The WebApp labels each unlocked range with the editor's email address, and unix timestamp for when the editing session expires. As future cells are unlocked for editing, the WebApp first cleans up expired editing sessions. 

Setup and User Guide
---

TODO

| Step |  | Description |
|:--:|:--:|---|
|  1 | [<img src="res/copy-sheet-button.png" alt="copy sheet" width="170" height="50">](https://www.github.com/)| <-- Click button to copy Spreadsheet + all code to your Google My Drive folder.<br>Your copy will be named "Copy of LockUnlock-Spreadsheet-v1" |
|  2 |                                                  | Choose Google Sheets menu item Extensions | Apps ScriptOpen Apps Script Editor. Set the projects GCP number                     |
|  3 | x                                                | Deploy the app as a WebApp, keep the file number. You'll need it soon|
|  4 |                                                  | <-- Click button to copy the Spreadsheet to your Google My Drive folder, the copy will be named "Copy of LockUnlock-Spreadsheet-v1"|
|  5 | x                                                | Set the project GCP number|
|  6 | x                                                | In the Apps Script code, find the line TODO, and replace the line
|  7 | DOne                                             | OK! YOu're reready
