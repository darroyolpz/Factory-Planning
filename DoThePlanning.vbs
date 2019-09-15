Sub DoThePlanning()

'Optimize Macro Speed
Application.ScreenUpdating = False
Application.EnableEvents = False
Application.Calculation = xlCalculationManual
Application.DisplayAlerts = False

'-------------------------------------------------------------------------'
'Dim newArray As Variant
Dim newArray() As String

'Select both worksheets'
Sheets("AHU").Select False
Sheets("BF").Select False

'Remove the colour of the alerts'
Cells.Select
With Selection.Interior
    .Pattern = xlNone
    .TintAndShade = 0
    .PatternTintAndShade = 0
End With
Sheets("Sales").Select

'Create table objects'
Set tSales = Range("sales_table").ListObject
Set tAHU = Range("ahu_table").ListObject
Set tBF = Range("bf_table").ListObject

'-------------------------------------------------------------------------'
'Clear filters'
tSales.AutoFilter.ShowAllData

'Sort by CO date'
tSales.Sort.SortFields.Add Key:=Range("sales_table[[CO Date]]"), _
SortOn:=sortonvalues, Order:=xlAscending

'-------------------------------------------------------------------------'
'First ask if salesID in ahu_table. If not, create new order (add to ahu_table)'
salesRows = tSales.ListRows.Count

For i = 2 To salesRows
    'Get ID from sales_table and look it up in ahu_table'
    salesID = tSales.ListColumns("ID").DataBodyRange(i)

    'If salesID not in tAHU['ID'] then create a NEW ORDER in tAHU'
    If IsError(Application.Match(salesID, tAHU.ListColumns("ID").Range, 0)) Then
        'Get the last row in ahu table'
        ahuRows = tAHU.ListRows.Count

        'Create an array to copy all the values from sales to ahu_table'
        newArray = Array("W", "Sales order no", "Pos", "Customer name", "CO Line sts", "MO sts", _ 
            "CO Item no", "CO Qty", "Order date", "MO Start", "Assembly hours", "MO Finish date", _
            "CO Date", "Delivery date", "Cost amount", "ID", "PO Data")

        'Add new order to ahu_table'
        For k = 0 To UBound(newArray)
            salesValue = tSales.ListColumns(k).DataBodyRange(i)
            tAHU.ListObject.ListColumns(k).DataBodyRange(ahuRows + 1) = salesValue
        Next k

    'If salesID in tAHU['ID'] then UPDATE & check if there is salesMO'
    Else
        'Get row index of ahu table'
        row_index = Application.Match(salesID, tAHU.ListColumns("ID").Range, 0)

        'UPDATE everything but MO'
        newArray = Array("W", "Sales order no", "Pos", "Customer name", "CO Line sts", "MO sts", _ 
            "CO Item no", "CO Qty", "Order date", "MO Start", "Assembly hours", "MO Finish date", _
            "CO Date", "Delivery date", "Cost amount", "ID", "PO Data")

        'Loop through the array and check sales value is the same as ahu value'
        For k = 0 To UBound(newArray)
            salesValue = tSales.ListColumns(k).DataBodyRange(i)
            ahuValue = tAHU.ListObject.ListColumns(k).DataBodyRange(row_index)
            If salesValue <> ahuValue Then
                tAHU.ListColumns(k).DataBodyRange(row_index) = salesValue
                tAHU.ListColumns(k).DataBodyRange(row_index).Interior.Color = RGB(255, 192, 0) 
            End If
        Next k

        'Check if there is salesMO. Create MO as a formula, not directly from Sales'
        salesMO = tSales.ListColumns("MO no").DataBodyRange(i)
        If salesMO <> "-" Then
            'If there is salesMO, ask if it has ahuMO. If not, update ahuMO and get M3 upload date'
            If IsError(Application.Match(salesMO, tAHU.ListColumns("MO no").Range, 0)) Then
                'Update ahuMO'
                tAHU.ListColumns("MO no").DataBodyRange(row_index) = salesMO

                'M3 upload date'
                tAHU.ListColumns("M3 upload date").DataBodyRange(row_index) = DateTime.Now - 1
            End If

        'If there is no salesMO ask if it was in Plan. If so, then STATUS 90'
        Else
            'Get planValue'
            planValue = Application.Index(tAHU.ListColumns("Was in Plan?").Range, row_index, 1)

            'If it was in Plan, then STATUS 90'
            If planValue = 1 Then
                'STATUS 90'
                tAHU.ListColumns("MO sts").DataBodyRange(row_index) = "90-90"
                tAHU.ListColumns("MO sts").DataBodyRange(row_index).Interior.Color = RGB(146, 208, 80)

                'Get the reported date'
                ahuReported = tAHU.ListColumns("Reported?").DataBodyRange(row_index)
                If ahuReported = 0 Then
                    tAHU.ListColumns("Reported?").DataBodyRange(row_index) = 1
                    tAHU.ListColumns("Report Date").DataBodyRange(row_index) = DateTime.Now - 1
                End If

            End If

        End If

    End If

Next i

'Reset Macro Optimization Settings
Application.ScreenUpdating = True
Application.EnableEvents = True
Application.Calculation = xlCalculationAutomatic
Application.DisplayAlerts = True

End Sub