Sub DoThePlanning()

'Optimize Macro Speed
Application.ScreenUpdating = False
Application.EnableEvents = False
Application.Calculation = xlCalculationManual
Application.DisplayAlerts = False

'-------------------------------------------------------------------------'
' Create ordersArray (columns Sales and in Orders must match)
Dim ordersArray() As Variant

'ordersArray will be the headers of Orders table'
ordersArray = Array("W", "Sales order no", "Pos", "Customer name", "Country", "Type", "CO Line sts", _
"Item name", "Qty", "Order date", "MO Start", "Assembly hours", "MO Finish date", _
"CO Date", "Delivery date", "Cost amount", "Unit", "PO Data")

'Select ORDERS worksheets to remove the colour'
Sheets(3).Select

'Remove the colour of the alerts'
Cells.Select
With Selection.Interior
    .Pattern = xlNone
    .TintAndShade = 0
    .PatternTintAndShade = 0
End With

'Create table objects'
Set tSales = Range("sales_table").ListObject
Set tORDERS = Range("orders_table").ListObject

'-------------------------------------------------------------------------'
'Clear filters'
tSales.AutoFilter.ShowAllData
tSales.Sort.SortFields.Clear

'Sort by CO date'
tSales.Sort.SortFields.Add Key:=Range("sales_table[[CO Date]]"), _
SortOn:=sortonvalues, Order:=xlAscending

'-------------------------------------------------------------------------'
'First ask if salesUnit in orders_table. If not, create new order (add to orders_table)'
salesRows = tSales.ListRows.Count

For i = 1 To salesRows
    'Get Unit from sales_table and look it up in orders_table'
    salesUnit = tSales.ListColumns("Unit").DataBodyRange(i)
    salesType = tSales.ListColumns("Type").DataBodyRange(i)

    If salesType <> "-" Then
        'NEW ORDER CASE: salesUnit not in tORDERS['Unit']. Create a NEW ORDER in tORDERS'
        'It shows an error because is false, so no match so NEW ORDER to be created'
        If IsError(Application.Match(salesUnit, tORDERS.ListColumns("Unit").Range, 0)) Then

            'Get the last row in orders table'
            ordersRows = tORDERS.ListRows.Count

            'Add NEW ORDER to orders_table'
            For Each k In ordersArray
                salesValue = tSales.ListColumns(k).DataBodyRange(i)
                tORDERS.ListColumns(k).DataBodyRange(ordersRows + 1) = salesValue
                tORDERS.ListColumns(k).DataBodyRange(ordersRows + 1).Interior.Color = RGB(255, 192, 0)
            Next k
            
            'Last update column'
            tORDERS.ListColumns("Last update").DataBodyRange(ordersRows + 1) = DateTime.Now - 1
            'Create a conditional that puts in red if there is no MO after 3 days of entering the order'

            '-------------------------------------------------------------------------'
            'Got MO? Compare MOs from tSales and tORDERS'
            k = "MO no"
            salesValue = tSales.ListColumns(k).DataBodyRange(i)
            ordersValue = tORDERS.ListColumns(k).DataBodyRange(ordersRows + 1)

            'Compare MOs and update M3 upload if needed'
            If (salesValue <> ordersValue) And (salesValue <> "") Then
                tORDERS.ListColumns(k).DataBodyRange(ordersRows + 1) = salesValue
                tORDERS.ListColumns(k).DataBodyRange(ordersRows + 1).Interior.Color = RGB(255, 192, 0)

                'M3 upload'
                tORDERS.ListColumns("M3 upload").DataBodyRange(ordersRows + 1) = DateTime.Now - 1

                'Last update'
                tORDERS.ListColumns("Last update").DataBodyRange(ordersRows + 1) = DateTime.Now - 1
            End If
            '-------------------------------------------------------------------------'


        'ALREADY ORDER CASE: If salesUnit in tORDERS['Unit'] then UPDATE & check if there is salesMO'
        Else
            'Get row index that match with salesUnit in orders table'
            row_index = Application.Match(salesUnit, tORDERS.ListColumns("Unit").Range, 0) - 1

            'UPDATE CASE: Loop through the array and check if sales value is the same as orders value'
            For Each k In ordersArray

                salesValue = tSales.ListColumns(k).DataBodyRange(i)
                ordersValue = tORDERS.ListColumns(k).DataBodyRange(row_index)

                'If different, UPDATE values in tORDERS. Don't update blank values'
                If (salesValue <> ordersValue) And (salesValue <> "") Then
                    tORDERS.ListColumns(k).DataBodyRange(row_index) = salesValue
                    tORDERS.ListColumns(k).DataBodyRange(row_index).Interior.Color = RGB(255, 192, 0)

                    'Last update column'
                    tORDERS.ListColumns("Last update").DataBodyRange(row_index) = DateTime.Now - 1
                End If
            Next k

            '-------------------------------------------------------------------------'
            'Got MO? Compare MOs from tSales and tORDERS'
            k = "MO no"
            salesValue = Val(tSales.ListColumns(k).DataBodyRange(i))
            ordersValue = tORDERS.ListColumns(k).DataBodyRange(row_index)

            'Compare MOs and update M3 upload if needed'
            If (salesValue <> ordersValue) And (salesValue <> "") Then
                tORDERS.ListColumns(k).DataBodyRange(row_index) = salesValue
                tORDERS.ListColumns(k).DataBodyRange(row_index).Interior.Color = RGB(255, 192, 0)

                'M3 upload'
                tORDERS.ListColumns("M3 upload").DataBodyRange(row_index) = DateTime.Now - 1

                'Last update'
                tORDERS.ListColumns("Last update").DataBodyRange(row_index) = DateTime.Now - 1
            End If
            '-------------------------------------------------------------------------'

            '-------------------------------------------------------------------------'
            'Status 90'
            k = "MO Sts"
            salesValue = tSales.ListColumns(k).DataBodyRange(i)
            ordersValue = tORDERS.ListColumns(k).DataBodyRange(row_index)

            'Compare status and update if needed'
            If (salesValue = "90-90") And (ordersValue <> salesValue) Then
                'STATUS 90'
                tORDERS.ListColumns(k).DataBodyRange(row_index) = salesValue
                tORDERS.ListColumns(k).DataBodyRange(row_index).Interior.Color = RGB(146, 208, 80)
                
                'Report date column'
                tORDERS.ListColumns("Report date").DataBodyRange(row_index) = DateTime.Now - 1

                'Last update column'
                tORDERS.ListColumns("Last update").DataBodyRange(row_index) = DateTime.Now - 1
            ElseIf (salesValue <> ordersValue) And (salesValue <> "") Then
                tORDERS.ListColumns(k).DataBodyRange(row_index) = salesValue
                tORDERS.ListColumns(k).DataBodyRange(row_index).Interior.Color = RGB(255, 192, 0)

                'Last update'
                tORDERS.ListColumns("Last update").DataBodyRange(row_index) = DateTime.Now - 1
            End If
            '-------------------------------------------------------------------------'

        End If
    End If

Next i

MsgBox "Done bro!"
Sheets(3).Cells(1, 1).Select

'Reset Macro Optimization Settings
Application.ScreenUpdating = True
Application.EnableEvents = True
Application.Calculation = xlCalculationAutomatic
Application.DisplayAlerts = True

End Sub
