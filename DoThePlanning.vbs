Sub DoThePlanning()

'Optimize Macro Speed
Application.ScreenUpdating = False
Application.EnableEvents = False
Application.Calculation = xlCalculationManual
Application.DisplayAlerts = False
'-------------------------------------------------------------------------'

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
'-------------------------------------------------------------------------'

'Clear filters'
Range("sales_table").ListObject.AutoFilter.ShowAllData

'Sort by CO date'
ActiveSheet.ListObjects("sales_table").Sort.SortFields.Add Key:=Range("sales_table[[CO Date]]"), _
SortOn:=sortonvalues, Order:=xlAscending
'-------------------------------------------------------------------------'

'Get matching between Sales and AHU'
salesRows = Range("sales_table").ListObject.ListRows.Count

For i = 1 To salesRows
    'Count rows on AHU data. Check if can be done fewer times'
    ahuRows = Range("ahu_table").ListObject.ListRows.Count
    salesID = Range("sales_table").ListObject.ListColumns("ID").DataBodyRange(i) 'Need to change column name at the file'
    salesType = Range("sales_table").ListObject.ListColumns("Type").DataBodyRange(i)
    ahuCount = 0
    bfCount = 0
    'Code to duplicate in case of a boxfan'
    If salesType = "AHU" Then
        For j = 0 To ahuRows
            ahuID = Range("ahu_table").ListObject.ListColumns("ID").DataBodyRange(j) 'Need to change column name at the file'

            'If ahuCount > ahuRows means that the salesID is a new unit'
            If ahuID <> salesID Then
                ahuCount = ahuCount + 1

            'If salesID found in ahuID'
            ElseIf ahuID = salesID Then
                'Check if the salesID has an MO already'
                salesMO = Range("sales_table").ListObject.ListColumns("MO").DataBodyRange(i)
                'If trimmed and no lenght, there is no MO yet. It is either fresh new CO or status 90-90'
                If Len(Trim(salesMO)) = 0 Then
                    'Check if there is changes in new lines that haven't been in planning'
                    ahuPlan = Range("ahu_table").ListObject.ListColumns("Was in Plan?").DataBodyRange(j)
                    'If hasnt been in plan, is fresh new CO'
                    If ahuPlan = 0 Then
                        'Crate an array to look for the columns'
                        newArray = Array("W", "CO Item no", "CO Qty", "CO Date", "Cost amount")
                        For k = 0 To UBound(newArray)
                            salesValue = Range("sales_table").ListObject.ListColumns(k).DataBodyRange(i)
                            ahuValue = Range("ahu_table").ListObject.ListColumns(k).DataBodyRange(j)
                            If salesValue <> ahuValue Then
                                Range("ahu_table").ListObject.ListColumns(k).DataBodyRange(j) = salesValue
                                Range("ahu_table").ListObject.ListColumns(k).DataBodyRange(j).Interior.Color = RGB(255, 192, 0) 
                            End If
                        Next k
                    ElseIf ahuPlan = 1 Then
                        Range("ahu_table").ListObject.ListColumns("MO sts").DataBodyRange(j) = "90-90"
                        Range("ahu_table").ListObject.ListColumns("MO sts").DataBodyRange(j).Interior.Color = RGB(146, 208, 80)
                        If  = 0 Then

                        End If
                    End If
                ElseIf

                ElseIf

                End If
            End If
        Next j
    End If
Next i


'Reset Macro Optimization Settings
Application.ScreenUpdating = True
Application.EnableEvents = True
Application.Calculation = xlCalculationAutomatic
Application.DisplayAlerts = True

End Sub