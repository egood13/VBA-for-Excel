Attribute VB_Name = "Module5"
Option Explicit


Sub qcFailReasonVL()

' set sub variables
Dim failedSheet As Worksheet
Dim vlRange As Range
Set failedSheet = ThisWorkbook.Sheets("artemis_failed_picks")
Set vlRange = failedSheet.Range("N2", failedSheet.Range("B2").End(xlDown).Offset(0, 12))

' set column name
failedSheet.Range("N1").Value = "qc fail reason"

vlRange.Formula = "=VLOOKUP(B2, 'apollo_fail_reasons'!B:H, 7, FALSE)"


End Sub
