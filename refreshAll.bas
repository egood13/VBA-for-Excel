Attribute VB_Name = "Module3"

Option Explicit

Sub refreshAll()

' set pivot table variables
Dim pickCount As PivotTable
Dim loggedHours As PivotTable
Dim estPickHours As PivotTable
Dim failed As PivotTable
Dim roster As QueryTable
Dim apollo_fails As QueryTable
Dim hours As PivotTable

Set pickCount = ThisWorkbook.Sheets("Hourly Pick Count By Employee"). _
                PivotTables("PivotTable2")
Set loggedHours = ThisWorkbook.Sheets("Logged Hours"). _
                PivotTables("PivotTable1")
Set estPickHours = ThisWorkbook.Sheets("Est. Picker Hours"). _
                PivotTables("PivotTable2")
Set failed = ThisWorkbook.Sheets("failed pivot").PivotTables("PivotTable1")
Set roster = ThisWorkbook.Sheets("Picker Names"). _
            ListObjects("Table_ExternalData_12").QueryTable
Set apollo_fails = ThisWorkbook.Sheets("Apollo Fails Picker"). _
            ListObjects("Table_Query_from_Apollo7").QueryTable


' refresh and unfilter pick count and average pick rate tables
pickCount.RefreshTable
pickCount.ClearAllFilters

loggedHours.RefreshTable
loggedHours.ClearAllFilters

estPickHours.RefreshTable
estPickHours.ClearAllFilters

roster.Refresh BackgroundQuery:=False
apollo_fails.Refresh BackgroundQuery:=False

' get LPN's from Apollo Fails Picker, refresh Artemis query
getFailReasons
failed.RefreshTable


End Sub
