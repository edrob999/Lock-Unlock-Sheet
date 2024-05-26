> [!CAUTION]
> This ReadMe and repo is still a work-in-process. Its not really ready to look at yet
> <BR>
> You're welcome to browse, but I'm still experimenting with stuff


Lock-Unlock-Sheet
=================

> [!NOTE]
> This is my first GitHub open source repo! **Please :star: if you think it will be useful for people.**
> <BR>

This repo demonstrates how to programatically lock/unlock cells in a Google sheet
---
Welcome! Looking for a way to lock/unlock cells in a Google sheet? In this repo, you'll find source code and ready-to-go spreadsheet that demonstrates a technique for programmatically locking/unlocking cells in a Google Workspace Sheet, so you can:
- Create locked forms in Google Sheets with editable form fields
- Capture audit history as people edit the form fields
- Unlock a single cell for editing, and lock the cell after they have finished
- Programatically lock/unlock a cell based on business rules, such as "a user can only edit empty cells" or "person x can override text in any cell" 

The sample is written entirely in Google Apps Script, distributed as open source (under the Apache2 license). You can use + change it as you wish. 
<BR>
<BR>
In this README you'll find: 
[Teaser video](here-is-the-teaser-video-that-shows-what-the-sample-does), 
[How-it-works](how-works), 
[Quickstart](quickstart), 
[How-to-customize](how-to-customize) and 
[Known issues](known-issues)
<BR>
I hope to include a try-before-you-buy sample, but waiting for Google to finish OAuth verification before I can publish it. 

### Here is the teaser video that shows what the sample does
<BR>
<a href="http://www.youtube.com/watch?feature=player_embedded&v=YOUTUBE_VIDEO_ID_HERE
" target="_blank"><img src="http://img.youtube.com/vi/YOUTUBE_VIDEO_ID_HERE/0.jpg" 
alt="IMAGE ALT TEXT HERE" width="240" height="180" border="10" /></a>
<!--
### Here is the try-before-you-buy sample:
<table>
  <tr>
    <td>TODO WAITING FOR OAUTH VERIFICATION</td>
    <td><-- Click this button to open the lock/unlock sample spreadsheet shown in the teaser video above. Your browser will need to be logged into a Google account</td>
  </tr>
</table>
-->

How it works
---
The Lock-Unlock-Sheet technique uses a dual-permission model:
* **Google Workspace Spreadsheet:** Users open a Spreadsheet, shared with Edit permissions + each Sheet locked to prevent changes. A user clicks a button to initiate editing. The button invokes a method in a WebApp to either change a cell's contents, or unlock the cell for free-editing 
* **WebApp:** The WebApp runs with the Spreadsheet owner's permission (it has full rights to change any content in the sheet, or unlock cells for editing). When user clicks an "Edit" button in the locked Spreadsheet, the WebApp goes to work, and becuase it is running with the owner's permission, it can either update a cell's text (for forms) or unlock a cell for editing (for free editing).

> [!IMPORTANT]
> If you are the owner of the sheet, you **always** have edit rights to every cell, so you won't experience the locking/unlocking
> <BR> To see it working, you'll need to use a Google account that has edit rights, but is **not the owner** of the Spreadsheet
> 

Unlocking a cell for free-editing uses two layers of Spreadsheet protection. The lower layer is the sheet-lock, which locks the sheet for everyone (except you the owner). The upper layer is a range protection unlocking a cell, and ensuring only the user requesting the edit is authorized to edit the cell. This means when a user requests to edit a cell, it is unlocked only for them. The WebApp labels each unlocked range with the editor's email address, and unix timestamp for when the editing session expires. As future cells are unlocked for editing, the WebApp cleans up expired editing sessions. 
<BR>
Sounds complicated? We've tried to make it as simple as possible in the quickstart below

Quickstart
---
In the steps below, you'll copy the code to your Google My Drive folder, and set up the Look-Unlock-Sheet sample
| Step | Description |
|:--:|---|
| [<img src="res/copy-sheet-button.png" alt="copy sheet" width="170" height="50">](https://www.github.com/)| <-- Click button to copy Spreadsheet + code to your Google My Drive folder.<br>After a couple minutes, "Copy of LockUnlock-Spreadsheet-v1" will open in the browser |
|  2 | **Deploy WebApp.** In Google Sheets, open Apps Script editor with menu item Extensions \| Apps Script.<br>Click Deploy \| New Deployment button in the editor, and deploy as a WebApp, executing as "me" (this should already be pre-populated) |
|  3 | **Authorize Access.** As the WebApp is deployed, you'll be asked to authorize the scopes required.<br>In this step, you're authorizing the WebApp permissions (not the Spreadsheet).<br>Wait! before clicking through everything, make sure to copy the Web app URL, you'll need it real soon|
|  4 | **Update WebApp Url.** In Apps Script editor, navigate to the SheetCode.gs file.<br>Update the first line that reads `var LIVE_URL = "TODO"` to use your WebApp Url.<br>The line should look something like: `var LIVE_URL = "https://script.google.com/macros/s/####/exec"` |
|  5 | **Save Your Work.** In Apps Script editor, save the project. You're finished, we're ready to go!|


Test the Look-Unlock-Sheet Sample
---------------------------------
> [!IMPORTANT]
> I know I keep nagging you about this, but its critical to remember.
> <BR>
> You'll need to use a different account that is **not** the owner of the spreadsheet to see the lock/unlock in action.
> <BR>
> OK. Promise I won't mention it again.


Your "Copy of LockUnlock-Spreadsheet-v1" contains two sheets.
<BR>

### Locked cells
Sheet1 behaves like a 'form', all cells are locked, the only cells that can be edited are those with a blue background. Click the edit button next to each cell and enter a value. In this screenshot, we've answered  "Dog" as our favorite animal. Because the cells in the sheet are locked, the Edit button invokes the web app to make the change, using the `set-cell-text` method.
![Sheet that behaves as a locked form](res/test-01-sheet1.png)

### Audit History
Sheet1 also demonstrates how to record an autdit history of changes. After the user has made a change in the form, the spreadsheet invokes the web app to update a scrolling audit history range (which is also locked). This is done using the `set-range-text` method. Also note we've anonymized email adddresses using the custom function `getMaskedEmail(emailAddress)`
![Audit history in a spreadsheet](res/test-02-audithistory.png)

### Free Editing
Sheet2 demonstrates unlocking a cell for free-editing. When a user positions focus on an empty cell, and clicks the Edit button, the cell is unlocked for them to edit for two minutes.<BR>
This is done by invoking the `start-cell-edit` method. The cell remains locked for everyone except the target user (and the spreadsheet owner) for two minutes. Each time a method is called in the web app, the code performs a quick check to see if there are any expired editing sessions, and relocks the cell if the session is expired. This is how the expiry is implemented.
![Free editing in locked sheet](res/test-03-freeedit.png)

How to Customize
----------------
Here are some notes for using the technique for locking/unlocking cells in your own Google Spreadsheet:
> [!NOTE]
> You don't need to re-deploy the WebApp. The same webapp will work with any sheet for which you are the owner, and only needs to be deployed once
> 

Here are the methods the WebApp makes available:<BR>
| Method          | Parameters                                                          | Description
|-----------------|---------------------------------------------------------------------|-------------|
|`get-info`       |                                                                     | return the name and version of the WebApp 
|`set-cell-text`  | `spreadsheetId` `sheetName1` `a1Notation` `text`                    | set the text of a cell (text contians a string)
|`set-range-text` | `spreadsheetId` `sheetName1` `a1Notation` `text`                    | set the text of a range (text contains a 2 dimensional array)
|`lock-sheet`     | `spreadsheetId` `sheetName1`                                        | lock the sheet, removing any open editing session
|`unlock-sheet`   | `spreadsheetId` `sheetName1`                                        | unlock the sheet
|`start-cell-edit`| `spreadsheetId` `sheetName1` `a1Notation` `emailAddress` `editTime` | enable free-editing for a cell
|`end-cell-edit`  | `spreadsheetId` `sheetName1` `emailAddress`                         | finishes free-editing session
|`generate-error` | `text`                                                              | generate an error (to test your error handling). text can be 0,1,2

Here is how you call the WebApp from your Apps Script Spreadsheet code, using the helper method `execCommand`.
In the following example, we unlock the cell A1 in the active sheet for 10 minutes
```
var retVal = execCommand("start-cell-edit", {
  spreadsheetId: SpreadsheetApp.getActive().getId(),
  sheetName: SpreadsheetApp.getActiveSheet().getName(),
  a1Notation: "A1",
  emailAddress: Session.getActiveUser().getEmail(),
  editTime: 10
})
```
We can then check 'retVal' (the returned value) for any errors returned from the WebApp
```
if( !retVal.ok) SpreadsheetApp.getUi().alert(retVal.errorMessage)
```
See the source code for more examples!

Known issues
------------
There are known issues in this sample application. Its not foolproof, and there is no input validation
- [ ] Clicking outside the edit box causes an error (standard Apps Script behavior). Thanks Oliver C
- [ ] `set-cell-text` Doesn't allow an empty string in the sample. Thanks Oliver C
- [ ] `set-cell-text` Using a formula as the `text` parameter, inserts the formula into the cell. 
- [ ] `start-cell-edit` It is possible to extend the editing range, using the cell shortcut menu. Thanks Oliver C
- [ ] It is possible to create a copy of the locked sheet (with no purpose, the copy will also be locked). Thanks Brett G

Let me know if you find anything else, and  **Please :star: if you think this sample will be useful for other people.**



