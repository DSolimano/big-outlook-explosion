Attribute VB_Name = "MailCleanupModule"
Dim batchSize As Integer

Public Sub cleanInbox()
    batchSize = 100
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
batchCount = (olFolder.Items.count / batchSize) + 1

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

Dim itemCount As Long
itemCount = olFolder.Items.count

If itemCount > batchSize Then
    itemCount = batchSize
End If

ReDim mailItems(0 To itemCount)


'We copy the folder contents to another array as olFolder.Items changes in realtime as you work
'So, without the copy, you only get half the items deleted
'An improvement would be to move the copy of the array up to the caller, so that we only do this once

Dim count As Long
count = 0

For Each olOldItem In olFolder.Items
    If count < itemCount Then
        Set mailItems(count) = olOldItem
        count = count + 1
    End If
    
    If count >= itemCount Then
        Exit For
    End If
    
Next olOldItem

'Count ends up one too high
count = count - 1

If count > batchSize + 1 Then
    count = batchSize + 1
End If

Dim i As Long
For i = 0 To count

    Dim itemToDelete As MailItem
    Set itemToDelete = mailItems(i)

    itemToDelete.Delete
    Set mailItems(i) = Nothing
    
    If (i Mod 10 = 0) Then
        DoEvents
    End If
Next
    
End Sub



