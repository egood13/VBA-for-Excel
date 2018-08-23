Attribute VB_Name = "Module4"

Option Explicit

Sub getFailReasons()

' define connection parameters
Dim oConn As New ADODB.Connection
Dim cmd As New ADODB.Command
Dim strUsername As String
Dim strPassword As String
Dim strDatabase As String

' Set connection parameters
strUsername = "user"
strPassword = "p7663c81910e7fc4093125e117ea284b8252091cb0473ea822729c758997e65b3"
strDatabase = "da4vufus9h81aq"

' SQL query
Dim strSQL As String
strSQL = "SELECT u.lpn " & _
    ",u.created_at AT TIME ZONE 'America/New_York' " & _
    ",u.updated_at AT TIME ZONE 'America/New_York' " & _
    ",u.picked_at AT TIME ZONE 'America/New_York' " & _
    ",u.printed_front_at AT TIME ZONE 'America/New_York' " & _
    ",u.printed_back_at AT TIME ZONE 'America/New_York' " & _
    ",u.picker " & _
    ",CASE WHEN u.fulfillment_status = 0 THEN 'unallocated' " & _
          "WHEN u.fulfillment_status = 1 THEN 'allocated' " & _
          "WHEN u.fulfillment_status = 2 THEN 'picked' " & _
          "WHEN u.fulfillment_status = 3 THEN 'printed' " & _
          "WHEN u.fulfillment_status = 4 THEN 'shipped' " & _
          "WHEN u.fulfillment_status = 5 THEN 'failed' " & _
          "WHEN u.fulfillment_status = 6 THEN 'cancelled' " & _
          "WHEN u.fulfillment_status = 7 THEN 'matched' " & _
          "WHEN u.fulfillment_status = 8 THEN 'pick error' " & _
          "Else 'other' END AS fulfillment_status " & _
    ",u.shipment_id" & _
    ",u.gtin " & _
    ",u.wave_id " & _
"FROM units u " & _
"WHERE u.lpn IN "
strSQL = strSQL & getDistinctLPNs() ' filter by LPNs in Apollo fails list

' sub variables
Dim i As Integer
Dim reasonSheet As Worksheet
Set reasonSheet = ThisWorkbook.Sheets("Artemis")

' open connection
oConn.Open ("DSN=Artemis;" & _
            "Database=" & strDatabase & ";" & _
            "UserID=" & strUsername & ";" & _
            "Password=" & strPassword)

' define command type
cmd.CommandType = ADODB.CommandTypeEnum.adCmdText
cmd.ActiveConnection = oConn
cmd.CommandText = strSQL

' get records
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Set rs = cmd.Execute

' clear sheet
reasonSheet.UsedRange.Clear

' paste column headers and records
For i = 0 To rs.Fields.Count - 1
    reasonSheet.Cells(1, i + 1).Value = rs.Fields(i).Name
Next i

reasonSheet.Range("A2").CopyFromRecordset rs

rs.Close
oConn.Close


End Sub
