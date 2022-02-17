Sub DoThePlanning()

'Optimize Macro Speed
Application.ScreenUpdating = False
Application.EnableEvents = False
Application.Calculation = xlCalculationManual
Application.DisplayAlerts = False
'-------------------------------------------------------------------------'
'Clean colours'
Sheets(3).Select
Cells.Select
With Selection.Interior
    .Pattern = xlNone
    .TintAndShade = 0
    .PatternTintAndShade = 0
End With

Sheets(4).Select
Cells.Select
With Selection.Interior
    .Pattern = xlNone
    .TintAndShade = 0
    .PatternTintAndShade = 0
End With
'-------------------------------------------------------------------------'
'Sort by CO Date @ Sales Sheet'
Sheets(2).Select
LastRowSales = Sheets(2).Cells(Cells.Rows.Count, 2).End(xlUp).Row

'Clear filters'
ActiveWorkbook.Worksheets("Sales").ListObjects("Sales").Sort.SortFields.Clear

'Sort by CO Date'
Sheets("Sales").ListObjects("Sales").Sort.SortFields.Add Key:=Range("Sales[[#All],[CO Date]]"), SortOn:=sortonvalues, Order:=xlAscending, DataOption:=xlSortNormal
Range("Sales[#All]").Select
With Worksheets("Sales").ListObjects("Sales").Sort
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With

Sheets(2).Cells(1, 1).Select
'Filtering'
'Range("Plan[[#Headers],[Type]]").Select
'ActiveSheet.ListObjects("Plan").Range.AutoFilter Field:=19, Criteria1:="AHU"
'----------------------------------------------------------------------------'
'Get line from Sales and look through AHU sheet'
For i = 3 To LastRowSales
    LastRowAHU = Sheets(3).Cells(Cells.Rows.Count, "B").End(xlUp).Row
    LastRowBX = Sheets(4).Cells(Cells.Rows.Count, "B").End(xlUp).Row
    salesLine = Sheets(2).Cells(i, "AL").Value 'Unit column (formula)'
    uType = Sheets(2).Cells(i, "AK").Value 'Type column (formula)'
    ahuCount = 0
    bxCount = 0
    If uType = "AHU" Then
        For j = 3 To LastRowAHU
            ahuLine = Sheets(3).Cells(j, "S").Value 'Unit column'
            
            If ahuLine <> salesLine Then
                ahuCount = ahuCount + 1

            'If found in AHU:'
            ElseIf ahuLine = salesLine Then
                'If today CO number is in Plan, update and track information'
                If Sheets(2).Cells(i, "AM").Value <> "-" Then 'MO column (formula)'
                    'Week'
                    If Sheets(3).Cells(j, "B").Value <> Sheets(2).Cells(i, "AO").Value Then
                        Sheets(3).Cells(j, "B").Value = Sheets(2).Cells(i, "AO").Value
                        Sheets(3).Cells(j, "B").Interior.Color = RGB(255, 192, 0)
                    End If

                    'Country'
                    If Sheets(3).Cells(j, "F").Value <> Sheets(2).Cells(i, "Y").Value Then
                        Sheets(3).Cells(j, "F").Value = Sheets(2).Cells(i, "Y").Value
                        Sheets(3).Cells(j, "F").Interior.Color = RGB(255, 192, 0)
                    End If

                    'CO Line Sts'
                    If Sheets(3).Cells(j, "G").Value <> Sheets(2).Cells(i, "AR").Value Then
                        Sheets(3).Cells(j, "G").Value = Sheets(2).Cells(i, "AR").Value
                        Sheets(3).Cells(j, "G").Interior.Color = RGB(255, 192, 0)
                    End If

                    'MO Sts'
                    If Sheets(3).Cells(j, "H").Value <> Sheets(2).Cells(i, "AS").Value Then
                        Sheets(3).Cells(j, "H").Value = Sheets(2).Cells(i, "AS").Value
                        Sheets(3).Cells(j, "H").Interior.Color = RGB(255, 192, 0)
                    End If
                    
                    'MO no'
                    If Sheets(3).Cells(j, "I").Value <> Sheets(2).Cells(i, "AM").Value Then
                        Sheets(3).Cells(j, "I").Value = Sheets(2).Cells(i, "AM").Value
                        'Sheets(3).Cells(j, "I").Interior.Color = RGB(255, 192, 0)'
                    End If
                    
                    'CO Item no'
                    If Sheets(3).Cells(j, "J").Value <> Sheets(2).Cells(i, "W").Value Then
                        Sheets(3).Cells(j, "J").Value = Sheets(2).Cells(i, "W").Value
                        Sheets(3).Cells(j, "J").Interior.Color = RGB(255, 192, 0)
                    End If

                    'CO Qty'
                    If Sheets(3).Cells(j, "K").Value <> Sheets(2).Cells(i, "AF").Value Then
                        Sheets(3).Cells(j, "K").Value = Sheets(2).Cells(i, "AF").Value
                        Sheets(3).Cells(j, "K").Interior.Color = RGB(255, 192, 0)
                    End If

                    'Order date'
                    If Sheets(3).Cells(j, "L").Value <> Sheets(2).Cells(i, "AU").Value Then
                        Sheets(3).Cells(j, "L").Value = Sheets(2).Cells(i, "AU").Value
                        Sheets(3).Cells(j, "L").Interior.Color = RGB(255, 192, 0)
                    End If

                    'MO Start'
                    If Sheets(3).Cells(j, "M").Value <> Sheets(2).Cells(i, "AV").Value Then
                        Sheets(3).Cells(j, "M").Value = Sheets(2).Cells(i, "AV").Value
                        Sheets(3).Cells(j, "M").Interior.Color = RGB(255, 192, 0)
                    End If

                    'Assembly hours'
                    If Sheets(3).Cells(j, "N").Value <> Sheets(2).Cells(i, "AW").Value Then
                        Sheets(3).Cells(j, "N").Value = Sheets(2).Cells(i, "AW").Value
                        Sheets(3).Cells(j, "N").Interior.Color = RGB(255, 192, 0)
                    End If
                    
                    'MO Finish date'
                    If Sheets(3).Cells(j, "O").Value <> Sheets(2).Cells(i, "AX").Value Then
                        Sheets(3).Cells(j, "O").Value = Sheets(2).Cells(i, "AX").Value
                        Sheets(3).Cells(j, "O").Interior.Color = RGB(255, 192, 0)
                    End If
                    
                    'CO Date'
                    If Sheets(3).Cells(j, "P").Value <> Sheets(2).Cells(i, "AN").Value Then
                        Sheets(3).Cells(j, "P").Value = Sheets(2).Cells(i, "AN").Value
                        Sheets(3).Cells(j, "P").Interior.Color = RGB(255, 192, 0)
                    End If
                    
                    'Delivery Date'
                    If Sheets(3).Cells(j, "Q").Value <> Sheets(2).Cells(i, "AY").Value Then
                        Sheets(3).Cells(j, "Q").Value = Sheets(2).Cells(i, "AY").Value
                        Sheets(3).Cells(j, "Q").Interior.Color = RGB(255, 192, 0)
                    End If

                    'Cost amount'
                    If Sheets(3).Cells(j, "R").Value <> Sheets(2).Cells(i, "AH").Value Then
                        Sheets(3).Cells(j, "R").Value = Sheets(2).Cells(i, "AH").Value
                        Sheets(3).Cells(j, "R").Interior.Color = RGB(255, 192, 0)
                    End If

                    'PO Data'
                    If Sheets(3).Cells(j, "T").Value <> Sheets(2).Cells(i, "AP").Value Then
                        Sheets(3).Cells(j, "T").Value = Sheets(2).Cells(i, "AP").Value
                        Sheets(3).Cells(j, "T").Interior.Color = RGB(255, 192, 0)
                    End If

                    Sheets(3).Cells(j, "V").Value = 1 'Was in plan? column'
                    
                'If today CO number isn't in Plan:'
                ElseIf Sheets(2).Cells(i, "AM").Value = "-" Then
                    'Check if there is changes in new lines that haven't been in planning'
                    If Sheets(3).Cells(j, "V").Value = 0 Then 'Was in plan? column'
                        'Week'
                        If Sheets(3).Cells(j, "B").Value <> Sheets(2).Cells(i, "AO").Value Then
                            Sheets(3).Cells(j, "B").Value = Sheets(2).Cells(i, "AO").Value
                            Sheets(3).Cells(j, "B").Interior.Color = RGB(255, 192, 0)
                        End If

                        'Country'
                        If Sheets(3).Cells(j, "F").Value <> Sheets(2).Cells(i, "Y").Value Then
                            Sheets(3).Cells(j, "F").Value = Sheets(2).Cells(i, "Y").Value
                            Sheets(3).Cells(j, "F").Interior.Color = RGB(255, 192, 0)
                        End If
                        
                        'CO Item no'
                        If Sheets(3).Cells(j, "J").Value <> Sheets(2).Cells(i, "W").Value Then
                            Sheets(3).Cells(j, "J").Value = Sheets(2).Cells(i, "W").Value
                            Sheets(3).Cells(j, "J").Interior.Color = RGB(255, 192, 0)
                        End If
    
                        'CO Qty'
                        If Sheets(3).Cells(j, "K").Value <> Sheets(2).Cells(i, "AF").Value Then
                            Sheets(3).Cells(j, "K").Value = Sheets(2).Cells(i, "AF").Value
                            Sheets(3).Cells(j, "K").Interior.Color = RGB(255, 192, 0)
                        End If
                        
                        'CO Date'
                        If Sheets(3).Cells(j, "P").Value <> Sheets(2).Cells(i, "AN").Value Then
                            Sheets(3).Cells(j, "P").Value = Sheets(2).Cells(i, "AN").Value
                            Sheets(3).Cells(j, "P").Interior.Color = RGB(255, 192, 0)
                        End If
    
                        'Cost amount'
                        If Sheets(3).Cells(j, "R").Value <> Sheets(2).Cells(i, "AH").Value Then
                            Sheets(3).Cells(j, "R").Value = Sheets(2).Cells(i, "AH").Value
                            Sheets(3).Cells(j, "R").Interior.Color = RGB(255, 192, 0)
                        End If
                    'If it has been already in Plan, go for 90-90 status'
                    ElseIf Sheets(3).Cells(j, "V").Value = 1 Then 'Was in plan? column'
                        Sheets(3).Cells(j, "H").Value = "90-90" 'MO sts column'
                        Sheets(3).Cells(j, "H").Interior.Color = RGB(146, 208, 80)
                        If Sheets(3).Cells(j, "X").Value = 0 Then 'Bol column'
                            Sheets(3).Cells(j, "W").Value = DateTime.Now - 1 'Report Date column'
                            Sheets(3).Cells(j, "X").Value = 1 'Bol column'
                        End If
                    End If
                End If
            End If
            
            'Add the line from Sales that we've been looking through AHU with no success'
            If ahuCount >= (LastRowAHU - 2) Then
                'Week'
                Sheets(3).Cells(LastRowAHU + 1, "B").Value = Sheets(2).Cells(i, "AO").Value
                Sheets(3).Cells(LastRowAHU + 1, "B").Interior.Color = RGB(255, 192, 0)
            
                'Sales order no'
                Sheets(3).Cells(LastRowAHU + 1, "C").Value = Sheets(2).Cells(i, "AQ").Value
                Sheets(3).Cells(LastRowAHU + 1, "C").Interior.Color = RGB(255, 192, 0)
                
                'Pos'
                Sheets(3).Cells(LastRowAHU + 1, "D").Value = Sheets(2).Cells(i, "K").Value
                Sheets(3).Cells(LastRowAHU + 1, "D").Interior.Color = RGB(255, 192, 0)
                
                'Customer name'
                Sheets(3).Cells(LastRowAHU + 1, "E").Value = Sheets(2).Cells(i, "D").Value

                'Country'
                Sheets(3).Cells(LastRowAHU + 1, "F").Value = Sheets(2).Cells(i, "Y").Value
                
                'CO Line sts'
                Sheets(3).Cells(LastRowAHU + 1, "G").Value = Sheets(2).Cells(i, "AR").Value
                
                'MO Sts'
                Sheets(3).Cells(LastRowAHU + 1, "H").Value = Sheets(2).Cells(i, "AS").Value
                
                'MO no'
                Sheets(3).Cells(LastRowAHU + 1, "I").Value = Sheets(2).Cells(i, "AT").Value

                'CO Item no'
                Sheets(3).Cells(LastRowAHU + 1, "J").Value = Sheets(2).Cells(i, "W").Value
                
                'CO Qty'
                Sheets(3).Cells(LastRowAHU + 1, "K").Value = Sheets(2).Cells(i, "AF").Value
                
                'Order date'
                Sheets(3).Cells(LastRowAHU + 1, "L").Value = Sheets(2).Cells(i, "AU").Value
                
                'MO Start'
                Sheets(3).Cells(LastRowAHU + 1, "M").Value = Sheets(2).Cells(i, "AV").Value

                'Assembly hours'
                Sheets(3).Cells(LastRowAHU + 1, "N").Value = Sheets(2).Cells(i, "AW").Value
                
                'MO Finish date'
                Sheets(3).Cells(LastRowAHU + 1, "O").Value = Sheets(2).Cells(i, "AX").Value
                
                'CO Date'
                Sheets(3).Cells(LastRowAHU + 1, "P").Value = Sheets(2).Cells(i, "AN").Value
                
                'Delivery date'
                Sheets(3).Cells(LastRowAHU + 1, "Q").Value = Sheets(2).Cells(i, "AY").Value
                
                'Cost amount'
                Sheets(3).Cells(LastRowAHU + 1, "R").Value = Sheets(2).Cells(i, "AH").Value
                
                'Unit'
                Sheets(3).Cells(LastRowAHU + 1, "S").Value = Sheets(2).Cells(i, "AL").Value
                Sheets(3).Cells(LastRowAHU + 1, "S").Interior.Color = RGB(255, 192, 0)
                
                'PO Data'
                Sheets(3).Cells(LastRowAHU + 1, "T").Value = Sheets(2).Cells(i, "AP").Value

                'Plan?'
                Sheets(3).Cells(LastRowAHU + 1, "V").Value = Sheets(2).Cells(i, "AZ").Value
            End If
        Next j
    ElseIf uType = "BOXFAN" Then
        For j = 3 To LastRowBX
            bxLine = Sheets(4).Cells(j, "S").Value 'Unit column'
            
            If bxLine <> salesLine Then
                bxCount = bxCount + 1
            'If found in BOX:'
            ElseIf bxLine = salesLine Then
                'If today CO number is in Plan, update and track information'
                If Sheets(2).Cells(i, "AM").Value <> "-" Then 'MO column'
                    'Week'
                    If Sheets(4).Cells(j, "B").Value <> Sheets(2).Cells(i, "AO").Value Then
                        Sheets(4).Cells(j, "B").Value = Sheets(2).Cells(i, "AO").Value
                        Sheets(4).Cells(j, "B").Interior.Color = RGB(255, 192, 0)
                    End If

                    'Country'
                    If Sheets(4).Cells(j, "F").Value <> Sheets(2).Cells(i, "Y").Value Then
                        Sheets(4).Cells(j, "F").Value = Sheets(2).Cells(i, "Y").Value
                        Sheets(4).Cells(j, "F").Interior.Color = RGB(255, 192, 0)
                    End If

                    'CO Line Sts'
                    If Sheets(4).Cells(j, "G").Value <> Sheets(2).Cells(i, "AR").Value Then
                        Sheets(4).Cells(j, "G").Value = Sheets(2).Cells(i, "AR").Value
                        Sheets(4).Cells(j, "G").Interior.Color = RGB(255, 192, 0)
                    End If

                    'MO Sts'
                    If Sheets(4).Cells(j, "H").Value <> Sheets(2).Cells(i, "AS").Value Then
                        Sheets(4).Cells(j, "H").Value = Sheets(2).Cells(i, "AS").Value
                        Sheets(4).Cells(j, "H").Interior.Color = RGB(255, 192, 0)
                    End If
                    
                    'MO no'
                    If Sheets(4).Cells(j, "I").Value <> Sheets(2).Cells(i, "AM").Value Then
                        Sheets(4).Cells(j, "I").Value = Sheets(2).Cells(i, "AM").Value
                        'Sheets(4).Cells(j, "I").Interior.Color = RGB(255, 192, 0)'
                    End If
                    
                    'CO Item no'
                    If Sheets(4).Cells(j, "J").Value <> Sheets(2).Cells(i, "W").Value Then
                        Sheets(4).Cells(j, "J").Value = Sheets(2).Cells(i, "W").Value
                        Sheets(4).Cells(j, "J").Interior.Color = RGB(255, 192, 0)
                    End If

                    'CO Qty'
                    If Sheets(4).Cells(j, "K").Value <> Sheets(2).Cells(i, "AF").Value Then
                        Sheets(4).Cells(j, "K").Value = Sheets(2).Cells(i, "AF").Value
                        Sheets(4).Cells(j, "K").Interior.Color = RGB(255, 192, 0)
                    End If

                    'Order date'
                    If Sheets(4).Cells(j, "L").Value <> Sheets(2).Cells(i, "AU").Value Then
                        Sheets(4).Cells(j, "L").Value = Sheets(2).Cells(i, "AU").Value
                        Sheets(4).Cells(j, "L").Interior.Color = RGB(255, 192, 0)
                    End If

                    'MO Start'
                    If Sheets(4).Cells(j, "M").Value <> Sheets(2).Cells(i, "AV").Value Then
                        Sheets(4).Cells(j, "M").Value = Sheets(2).Cells(i, "AV").Value
                        Sheets(4).Cells(j, "M").Interior.Color = RGB(255, 192, 0)
                    End If

                    'Assembly hours'
                    If Sheets(4).Cells(j, "N").Value <> Sheets(2).Cells(i, "AW").Value Then
                        Sheets(4).Cells(j, "N").Value = Sheets(2).Cells(i, "AW").Value
                        Sheets(4).Cells(j, "N").Interior.Color = RGB(255, 192, 0)
                    End If
                    
                    'MO Finish date'
                    If Sheets(4).Cells(j, "O").Value <> Sheets(2).Cells(i, "AX").Value Then
                        Sheets(4).Cells(j, "O").Value = Sheets(2).Cells(i, "AX").Value
                        Sheets(4).Cells(j, "O").Interior.Color = RGB(255, 192, 0)
                    End If
                    
                    'CO Date'
                    If Sheets(4).Cells(j, "P").Value <> Sheets(2).Cells(i, "AN").Value Then
                        Sheets(4).Cells(j, "P").Value = Sheets(2).Cells(i, "AN").Value
                        Sheets(4).Cells(j, "P").Interior.Color = RGB(255, 192, 0)
                    End If
                    
                    'Delivery Date'
                    If Sheets(4).Cells(j, "Q").Value <> Sheets(2).Cells(i, "AY").Value Then
                        Sheets(4).Cells(j, "Q").Value = Sheets(2).Cells(i, "AY").Value
                        Sheets(4).Cells(j, "Q").Interior.Color = RGB(255, 192, 0)
                    End If

                    'Cost amount'
                    If Sheets(4).Cells(j, "R").Value <> Sheets(2).Cells(i, "AH").Value Then
                        Sheets(4).Cells(j, "R").Value = Sheets(2).Cells(i, "AH").Value
                        Sheets(4).Cells(j, "R").Interior.Color = RGB(255, 192, 0)
                    End If

                    'PO Data'
                    If Sheets(4).Cells(j, "T").Value <> Sheets(2).Cells(i, "AP").Value Then
                        Sheets(4).Cells(j, "T").Value = Sheets(2).Cells(i, "AP").Value
                        Sheets(4).Cells(j, "T").Interior.Color = RGB(255, 192, 0)
                    End If

                    Sheets(4).Cells(j, "V").Value = 1 'Was in plan? column'
                    
                'If today CO number isn't in Plan:'
                ElseIf Sheets(2).Cells(i, "AM").Value = "-" Then 'MO column'
                    If Sheets(4).Cells(j, "V").Value = 0 Then 'Was in plan? column'
                        'Week'
                        If Sheets(4).Cells(j, "B").Value <> Sheets(2).Cells(i, "AO").Value Then
                            Sheets(4).Cells(j, "B").Value = Sheets(2).Cells(i, "AO").Value
                            Sheets(4).Cells(j, "B").Interior.Color = RGB(255, 192, 0)
                        End If

                        'Country'
                        If Sheets(4).Cells(j, "F").Value <> Sheets(2).Cells(i, "Y").Value Then
                            Sheets(4).Cells(j, "F").Value = Sheets(2).Cells(i, "Y").Value
                            Sheets(4).Cells(j, "F").Interior.Color = RGB(255, 192, 0)
                        End If
                        
                        'CO Item no'
                        If Sheets(4).Cells(j, "J").Value <> Sheets(2).Cells(i, "W").Value Then
                            Sheets(4).Cells(j, "J").Value = Sheets(2).Cells(i, "W").Value
                            Sheets(4).Cells(j, "J").Interior.Color = RGB(255, 192, 0)
                        End If
    
                        'CO Qty'
                        If Sheets(4).Cells(j, "K").Value <> Sheets(2).Cells(i, "AF").Value Then
                            Sheets(4).Cells(j, "K").Value = Sheets(2).Cells(i, "AF").Value
                            Sheets(4).Cells(j, "K").Interior.Color = RGB(255, 192, 0)
                        End If
                        
                        'CO Date'
                        If Sheets(4).Cells(j, "P").Value <> Sheets(2).Cells(i, "AN").Value Then
                            Sheets(4).Cells(j, "P").Value = Sheets(2).Cells(i, "AN").Value
                            Sheets(4).Cells(j, "P").Interior.Color = RGB(255, 192, 0)
                        End If
    
                        'Cost amount'
                        If Sheets(4).Cells(j, "R").Value <> Sheets(2).Cells(i, "AH").Value Then
                            Sheets(4).Cells(j, "R").Value = Sheets(2).Cells(i, "AH").Value
                            Sheets(4).Cells(j, "R").Interior.Color = RGB(255, 192, 0)
                        End If
                    'If it has been already in Plan, go for 90-90 status'
                    ElseIf Sheets(4).Cells(j, "V").Value = 1 Then 'Was in plan? column'
                        Sheets(4).Cells(j, "H").Value = "90-90" 'MO sts column'
                        Sheets(4).Cells(j, "H").Interior.Color = RGB(146, 208, 80) 'MO sts column'
                        If Sheets(4).Cells(j, "X").Value = 0 Then 'Bol column'
                            Sheets(4).Cells(j, "W").Value = DateTime.Now - 1 'Report Date column'
                            Sheets(4).Cells(j, "X").Value = 1 'Bol column'
                        End If
                    End If
                End If
            End If
            
            'Add the line from Sales that we've been looking through BX with no success'
            If bxCount >= (LastRowBX - 2) Then
                'Week'
                Sheets(4).Cells(LastRowBX + 1, "B").Value = Sheets(2).Cells(i, "AO").Value
                Sheets(4).Cells(LastRowBX + 1, "B").Interior.Color = RGB(255, 192, 0)
            
                'Sales order no'
                Sheets(4).Cells(LastRowBX + 1, "C").Value = Sheets(2).Cells(i, "AQ").Value
                Sheets(4).Cells(LastRowBX + 1, "C").Interior.Color = RGB(255, 192, 0)
                
                'Pos'
                Sheets(4).Cells(LastRowBX + 1, "D").Value = Sheets(2).Cells(i, "K").Value
                Sheets(4).Cells(LastRowBX + 1, "D").Interior.Color = RGB(255, 192, 0)
                
                'Customer name'
                Sheets(4).Cells(LastRowBX + 1, "E").Value = Sheets(2).Cells(i, "D").Value

                'Country'
                Sheets(4).Cells(LastRowBX + 1, "F").Value = Sheets(2).Cells(i, "Y").Value
                
                'CO Line sts'
                Sheets(4).Cells(LastRowBX + 1, "G").Value = Sheets(2).Cells(i, "AR").Value
                
                'MO Sts'
                Sheets(4).Cells(LastRowBX + 1, "H").Value = Sheets(2).Cells(i, "AS").Value
                
                'MO no'
                Sheets(4).Cells(LastRowBX + 1, "I").Value = Sheets(2).Cells(i, "AT").Value

                'CO Item no'
                Sheets(4).Cells(LastRowBX + 1, "J").Value = Sheets(2).Cells(i, "W").Value
                
                'CO Qty'
                Sheets(4).Cells(LastRowBX + 1, "K").Value = Sheets(2).Cells(i, "AF").Value
                
                'Order date'
                Sheets(4).Cells(LastRowBX + 1, "L").Value = Sheets(2).Cells(i, "AU").Value
                
                'MO Start'
                Sheets(4).Cells(LastRowBX + 1, "M").Value = Sheets(2).Cells(i, "AV").Value

                'Assembly hours'
                Sheets(4).Cells(LastRowBX + 1, "N").Value = Sheets(2).Cells(i, "AW").Value
                
                'MO Finish date'
                Sheets(4).Cells(LastRowBX + 1, "O").Value = Sheets(2).Cells(i, "AX").Value
                
                'CO Date'
                Sheets(4).Cells(LastRowBX + 1, "P").Value = Sheets(2).Cells(i, "AN").Value
                
                'Delivery date'
                Sheets(4).Cells(LastRowBX + 1, "Q").Value = Sheets(2).Cells(i, "AY").Value
                
                'Cost amount'
                Sheets(4).Cells(LastRowBX + 1, "R").Value = Sheets(2).Cells(i, "AH").Value
                
                'Unit'
                Sheets(4).Cells(LastRowBX + 1, "S").Value = Sheets(2).Cells(i, "AL").Value
                Sheets(4).Cells(LastRowBX + 1, "S").Interior.Color = RGB(255, 192, 0)
                
                'PO Data'
                Sheets(4).Cells(LastRowBX + 1, "T").Value = Sheets(2).Cells(i, "AP").Value

                'Plan?'
                Sheets(4).Cells(LastRowBX + 1, "V").Value = Sheets(2).Cells(i, "AZ").Value
            End If
        Next j
    End If
Next i

'Go to AHU sheet and select first cell'
Sheets(3).Select
Cells(1, 1).Select

'Call DV_Alerts'

MsgBox ("Done bro!")
'Reset Macro Optimization Settings
Application.ScreenUpdating = True
Application.EnableEvents = True
Application.Calculation = xlCalculationAutomatic
Application.DisplayAlerts = True
Exit Sub

ErrorMessage:
    If LastLookUp <> 2 Then
    'MsgBox ("Watch out! Item " & Item & " not found in inventory!")
    Sheets(3).Select
    'NextEmpty = Cells(Cells.Rows.Count, "B").End(xlUp).Row + 1
    'Cells(NextEmpty, "B") = Item
    'Cells(NextEmpty, "C") = "No aparece en inventario!"
    'Cells(NextEmpty, "D") = "No aparece en inventario!"
    Resume Next
    End If

'Reset Macro Optimization Settings
Application.ScreenUpdating = True
Application.EnableEvents = True
Application.Calculation = xlCalculationAutomatic
Application.DisplayAlerts = True

End Sub