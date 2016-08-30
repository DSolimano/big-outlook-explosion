Attribute VB_Name = "MailCleanupModule"
Public Sub cleanInbox()
    Dim messageBoxRetVal As Integer
    MoveAndClearFolderInLoop (Session.GetDefaultFolder(olFolderInbox).Folders("TopLevelExampleFolder"))
    MoveAndClearFolderInLoop (Session.GetDefaultFolder(olFolderInbox).Folders("Application Folder"))
    MoveAndClearFolderInLoop (Session.GetDefaultFolder(olFolderInbox).Folders("Application Folder").Folders("LowLevelFolder"))
    
    messageBoxRetVal = MsgBox("Mail cleanup complete!", vbOKOnly, "Mail Cleanup")
End Sub


'What's going on here is that Exchange doesn't like having too many mail items at once open
'  I was unable to close them to Exchange's satisfaction in MoveAndClearFolderInBatch,
'  but doing it in batches of 100 is fine it seems.
Sub MoveAndClearFolderInLoop(olFolder As folder)

Dim batchCount As Integer
batchCount = (olFolder.Items.count / 100) + 1

Dim i As Long

For i = 0 To batchCount
    Dim myfolder As folder
    Set myfolder = olFolder
    
    MoveAndClearFolderInBatch myfolder
Next i
    
End Sub

Sub MoveAndClearFolderInBatch(olFolder As folder)

Dim olOldItem As MailItem

Dim mailItems() As MailItem

ReDim mailItems(0 To olFolder.Items.count)


'We copy the folder contents to another array as olFolder.Items changes in realtime as you work
'So, without the copy, you only get half the items deleted
'An improvement would be to move the copy of the array up to the caller, so that we only do this once

Dim count As Long
count = 0

For Each olOldItem In olFolder.Items
    Set mailItems(count) = olOldItem
    count = count + 1
Next olOldItem

'Count ends up one too high
count = count - 1

'Batch size is 101, more or less.  Should use the same constant as MoveAndClearFolderInLoop
If count > 101 Then
    count = 101
End If

Dim i As Long
For i = 0 To count

    Dim itemToDelete As MailItem
    Set itemToDelete = mailItems(i)

    itemToDelete.Delete
    Set mailItems(i) = Nothing
    'Try to keep the UI responsive    
    If (i Mod 2 = 0) Then
        DoEvents
    End If
Next
    
End Sub


