Attribute VB_Name = "Module6"


Sub Extract_EOS()

Dim myDir As String
Dim archiveFilename As String
Dim wkbookCheck As Integer
Dim sheetCount As Integer
Dim wkbTracker As Workbook
Dim shtTracker As Worksheet
Dim wkbArchive As Workbook
Dim shtArchive As Worksheet
Dim strArchiveName As String

' save file before extracting archive
ThisWorkbook.Save

' get current working directory and check for archive file
myDir = CurDir()
archiveFilename = "Picking_Tracker_Archive"
wkbookCheck = checkForWorkbook(archiveFilename)


' make sure archive file is open or saved in the same directory
If wkbookCheck = 1 Then
    MsgBox ("Picking_Tracker_Archive is not open but exists in current " & _
               "directory. Press 'Ctrl+O' and open and re-run macro")
    Exit Sub
ElseIf wkbookCheck = 0 Then
    MsgBox ("Picking_Tracker_Archive is not open and is not in the same " & _
            "active directory. Please find Picking_Tracker_Archive and " & _
            "open. Exiting macro.")
    Exit Sub
End If

' set variables
Set wkbTracker = Workbooks("Picking_Tracker")
Set shtTracker = wkbTracker.Sheets("EOS Summary")
Set wkbArchive = Workbooks(archiveFilename)


'----------EOS Summary----------
' save to end of workbook as values
copyTab wkbTracker, wkbArchive, shtTracker
' set new tab variable and get name
Set shtArchive = wkbArchive.Sheets(wkbArchive.Sheets.Count)
strArchiveName = shtArchive.Name & " W.E " & getWeekEnding(shtArchive.Range("I6"))
' update tab name
updateSheetName wkbArchive, shtArchive, strArchiveName

'----------EOS Employee----------
' change source sheet to EOS tab for employee level
Set shtTracker = wkbTracker.Sheets("EOS")
' save to end of workbook as values
copyTab wkbTracker, wkbArchive, shtTracker
' set new tab variable and get name
Set shtArchive = wkbArchive.Sheets(wkbArchive.Sheets.Count)
strArchiveName = shtArchive.Name & " W.E " & getWeekEnding(shtArchive.Range("I4"))
' update tab name
updateSheetName wkbArchive, shtArchive, strArchiveName

' save archive
wkbArchive.Save



End Sub

Sub updateSheetName(wkBook As Workbook, wkSheet As Worksheet, strName As String)
    ''' check if sheet name already exists. if so, delete old tab
    If checkSheetExists(wkBook, strName) Then
        Application.DisplayAlerts = False ' don't show alert
        wkBook.Sheets(strName).Delete
        Application.DisplayAlerts = True  ' add alerts back in
        wkSheet.Name = strName
    Else
        wkSheet.Name = strName
    End If

End Sub

Function checkSheetExists(wkBook As Workbook, shtName As String) As Boolean
    
    Dim ws As Worksheet
    For Each ws In wkBook.Sheets
        If ws.Name = shtName Then
            checkSheetExists = True
            Exit Function
        End If
    Next
    checkSheetExists = False

End Function


Function checkForWorkbook(wkBook As String) As Integer
''' Check if file is open or exists in current working directory
''' 0 = not open or in directory
''' 1 = not open but in directory
''' 2 = open

FullName = wkBook & ".xlsm"

On Error Resume Next

If Workbooks(wkBook) Is Nothing Then
    If Dir(FullName) <> "" Then
        checkForWorkbook = 1
    Else
        checkForWorkbook = 0
        Exit Function
    End If
Else
    checkForWorkbook = 2
End If

End Function


Function getWeekEnding(rangeDate As Range, Optional underscores As Boolean = True) As String

strDate = CStr(rangeDate.Value)
strDate = Replace(strDate, "/", "_")

getWeekEnding = strDate

End Function


Sub copyTab(sourceWkbook As Workbook, targetWkbook As Workbook, _
                sourceWksheet As Worksheet, _
                Optional pasteValues As Boolean = True)

Dim sheetCount As Integer
Dim targetWksheet As Range


Application.CopyObjectsWithCells = False ' don't copy macro buttons

' get sheet count and copy source sheet over
sheetCount = targetWkbook.Sheets.Count
sourceWksheet.Copy After:=targetWkbook.Sheets(sheetCount)
' set range of new worksheet copied over and copy/paste as values
Set targetWksheet = targetWkbook.Sheets(sheetCount + 1).Cells
targetWksheet.Copy
targetWksheet.PasteSpecial Paste:=xlPasteValues
Application.CutCopyMode = False

targetWksheet.Range("A1").Select

Application.CopyObjectsWithCells = True ' reset setting

End Sub


