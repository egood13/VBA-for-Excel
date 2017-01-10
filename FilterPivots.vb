
Option Explicit On

Sub Main()
    ' worksheets used
    Dim employeeTab As Worksheet
    Dim casgafaTab As Worksheet
    Dim bldg1Tab As Worksheet
    Dim bldg1TaxTab As Worksheet
    Dim bldg1AccrualTab As Worksheet

    ' pivot tables
    Dim casgafaPivot1 As PivotTable
    Dim casgafaPivot2 As PivotTable
    Dim bldg1Pivot As PivotTable
    Dim bldg1TaxPivot As PivotTable
    Dim bldg1AccrualPivot As PivotTable

    ' employees to filter
    Dim employees As Range
    Dim empl_str As Object

    ' set worksheets
    employeeTab = Worksheets("worked_hours_empl_filter")
    casgafaTab = Worksheets("casgafa1_worked_hours_JE")
    bldg1Tab = Worksheets("bldg1_worked_hours_JE")
    bldg1TaxTab = Worksheets("bldg1_tax_benefits")
    bldg1AccrualTab = Worksheets("bldg1_accr")

    ' set pivot tables
    casgafaPivot1 = casgafaTab.PivotTables("casgafaPivot1")
    casgafaPivot2 = casgafaTab.PivotTables("casgafaPivot2")
    bldg1Pivot = bldg1Tab.PivotTables("bldg1Pivot")
    bldg1TaxPivot = bldg1TaxTab.PivotTables("bldg1TaxPivot")
    bldg1AccrualPivot = bldg1AccrualTab.PivotTables("bldg1AccrualPivot")

    ' set employee list range
    employees = employeeTab.Range("A1", employeeTab.Range("A1").End(xlDown))

    ' filter pivot tables
    FilterPivot(casgafaPivot1, "employee_ssn", employees, True)
    FilterPivot(casgafaPivot2, "employee_ssn", employees, True)
    FilterPivot(bldg1Pivot, "employee_ssn", employees, False)
    FilterPivot(bldg1TaxPivot, "employee_ssn", employees, False)
    FilterPivot(bldg1AccrualPivot, "employee_ssn", employees, False)

End Sub

Sub FilterPivot(table As PivotTable, field As String, filter As Range, show As Boolean)
    ' Filters pivot table field by "field" given a range as "filter".
    ' Filters to select only values in range if "show" is true and deselects
    ' all values in range if "show" is false.

    Dim value As Object

    table.PivotCache.MissingItemsLimit = xlMissingItemsNone             ' clear cache and set to None
    table.PivotFields(field).ClearAllFilters()                            ' clear filter

    If show Then
        ' note: cannot deselect all. You must select only one then loop again for those
        ' you want visible
        For Each value In table.PivotFields(field).PivotItems
            If value = CStr(filter.Item(1)) Then                        ' test for one value
                value.Visible = True
            Else
                value.Visible = False
            End If
        Next
        ' loop through all values and select those from list
        For Each value In filter
            table.PivotFields(field).PivotItems(CStr(value)).Visible = True
        Next
        ' if show = false, deselect those in list
    Else
        For Each value In filter
            table.PivotFields(field).PivotItems(CStr(value)).Visible = False
        Next
    End If
End Sub
