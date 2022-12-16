# Dialog-API-Bug
A minimalist Outlook web add-in to replicate a bug in the Office dialog API.

## Replication steps:

### Installing this add-in
1. Download the Manifest.xml file from this repository and save it to your local machine.
2. Open Outlook Web Access in a browser, and click the "Get Add-Ins" button to open the "Add-Ins for Outlook" popup window. 
3. Select "My add-ins" from the menu on the left.
4. Click "Add a custom add-in," and select "Add from File" from the dropdown menu.
5. Use the file picker to select the manifest file saved to your local machine in step one.
6. Click "Install"

### Using the add-in to replicate the bug
1. Select an email from your inbox, and click the ellipses on the _MessageReadCommandSurface_ extension point (upper right corner of the item view).
2. Click "Dialog API Bug Demo" from the dropdown menu.
3. This will open a dialog box with two buttons, "Cancel" and "Move Message to Junk"
4. Click "Move Message to Junk" to move the selected email to the "Junk Email" folder.
5. Repeat steps 1 thru 4 on a different email in the inbox. During this iteration, notice that the buttons in the dialog box are non-responsive; the dialog API event handlers have stopped capturing `DialogMessageReceived` events.

See the browser console output for additional information while using this add-in.