Attribute VB_Name = "Module2"

Option Explicit

Function getDistinctLPNs()
''' Use this function to get the string of LPNs showing
''' in the artemis_failed picks. You can then use the
''' string to filter in Apollo DB to get your fail reasons

' set sub variables
Dim failedSheet As Worksheet
Dim lpnList As Range
Dim lpnCell As Range
Dim firstLPN As Range
Dim lastLPN As Range
Dim lpnString As String

Set failedSheet = ThisWorkbook.Sheets("Apollo Fails Picker")
Set lpnList = failedSheet.Range("A2", failedSheet.Range("A2").End(xlDown))
Set firstLPN = failedSheet.Range("A2")
Set lastLPN = failedSheet.Range("A2").End(xlDown)

' test if LPN is beginning or end and append to string accordingly
For Each lpnCell In lpnList
    If lpnCell = firstLPN Then
        lpnString = "('" & lpnCell & "',"
    ElseIf lpnCell = lastLPN Then
        lpnString = lpnString & "'" & lpnCell & "')"
    Else
        lpnString = lpnString & "'" & lpnCell & "',"
    End If
Next lpnCell

getDistinctLPNs = lpnString

End Function
