

Private Sub PickFileToImport_DropButtonClick()
    PickFileToImport.List = Array("Arrow Excel File")
End Sub

Private Sub PickFileToImport_Change()
    If PickFileToImport = "Arrow Excel File" Then
        ArrowDataCopy
    End If
    PickFileToImport = "Pick File To Import"
End Sub





Sub AddOneRow()
'
'   Rows have to be added by buttons to make sure data is properly set
'
    Rows("31:31").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrBelow
    Rows("32:32").Select
    Selection.Copy
    Rows("31:31").Select
    ActiveSheet.Paste
    
    ActiveSheet.Range("F31:E31").Value = "'"
    ActiveSheet.Range("F31:F31").Value = "'"
    ActiveSheet.Range("I31:I31").Value = "'"
    ActiveSheet.Range("Q31:Q31").Value = "'"
    ActiveSheet.Range("S31:S31").Value = "'"
    ActiveSheet.Range("P31:T31").Value = "'"
    ActiveSheet.Range("K31:K31").Value = "'"
    
    Application.CutCopyMode = False
    Range("D21").Select
'======================================================================================================================================================'
'======================================================================================================================================================'
'======================================================================================================================================================'
End Sub
Sub AddTenRows()
'
'   Rows have to be added by buttons to make sure data is properly set
'
    Rows("31:40").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Rows("41:41").Select
    Selection.Copy
    Rows("31:40").Select
    ActiveSheet.Paste
    
    ActiveSheet.Range("F31:E40").Value = "'"
    ActiveSheet.Range("I31:I40").Value = "'"
    ActiveSheet.Range("Q31:Q40").Value = "'"
    ActiveSheet.Range("S31:S40").Value = "'"
    ActiveSheet.Range("P31:T40").Value = "'"
    ActiveSheet.Range("K31:K40").Value = "'"
    ActiveSheet.Range("H31:H40").Value = "'"
    
    Application.CutCopyMode = False
    Range("D19").Select
'======================================================================================================================================================'
'======================================================================================================================================================'
'======================================================================================================================================================'
End Sub

Sub ValidateData()
'
'
'
    Dim length_counter As Integer
    length_counter = 30

    ' Count length of table
    For Each i In Worksheets(1).Range("F31:F10000")
        If i.Value <> "" Then
            length_counter = length_counter + 1
        Else
            Exit For
        End If
    Next i
    
    If length_counter > 30 Then
        ' Make sure data in optional rows are set properly
        If Worksheets(1).Range("C10").Value = "" Then
            Worksheets(1).Range("C10").Value = "'"
        End If
        If Worksheets(1).Range("C12").Value = "" Then
            Worksheets(1).Range("C12").Value = "'"
        End If
        If Worksheets(1).Range("C18").Value = "" Then
            Worksheets(1).Range("C18").Value = "'"
        End If
        If Worksheets(1).Range("C19").Value = "" Then
            Worksheets(1).Range("C19").Value = "'"
        End If
        If Worksheets(1).Range("C20").Value = "" Then
            Worksheets(1).Range("C20").Value = "'"
        End If
        If Worksheets(1).Range("C21").Value = "" Then
            Worksheets(1).Range("C21").Value = "'"
        End If
        If Worksheets(1).Range("C26").Value = "" Then
            Worksheets(1).Range("C26").Value = "'"
        End If

        ' Make sure data in hidden columns are set properly
        For Each i In Worksheets(1).Range("G31:G" & CStr(length_counter))
            i.Value = 5
        Next i
        For Each i In Worksheets(1).Range("J31:K" & CStr(length_counter))
            i.Value = "'"
        Next i
        For Each i In Worksheets(1).Range("P31:P" & CStr(length_counter))
            i.Value = "'"
        Next i
        For Each i In Worksheets(1).Range("R31:R" & CStr(length_counter))
            i.Value = "'"
        Next i
        For Each i In Worksheets(1).Range("L31:N" & CStr(length_counter))
            i.Value = 0
        Next i
        For Each i In Worksheets(1).Range("O31:O" & CStr(length_counter))
            i.Value = 1
        Next i

        Worksheets(1).Range(CStr(length_counter + 1) & ":" & CStr(length_counter + 100)).Delete Shift:=xlDown

        MsgBox "Data validation complete."
    Else
        MsgBox "Insert at least one RIM number."
    End If

End Sub

Sub ArrowDataCopy()
'======================================================================================================================================================'
'======================================== MACRO TO CONVERT ARROW CONFIRMATION EXCEL INTO XML =========================================================='
'======================================== ONE XML FOR EACH ORDER ======================================================================================'
'======================================== BY PIOTR KOCHANY 20-06-2017 ================================================================================='
'======================================================================================================================================================'
    ' Open selected file and set both workbooks to variables
    Dim wbRimaster As Workbook, wbArrow As Workbook
    Dim ws As Worksheet
    Dim vFile As Variant

    'Set source workbook
    Set wbRimaster = ActiveWorkbook
    'Open the target workbook
    vFile = Application.GetOpenFilename("(.xls),*.xls,(.xlsx),*.xlsx,(.xlsm),*.xlsm", 1, "Select Arrow Order Confirmation", , False)
        
    'if the user didn't select a file, exit sub
    If TypeName(vFile) = "Boolean" Then Exit Sub
    Workbooks.Open vFile
    'Set targetworkbook
    Set wbArrow = ActiveWorkbook
'======================================================================================================================================================'
    ' Before anything make sure a correct file is opened
    ' A1 cell for Arrow excel always contain same sentence
    If Trim(Replace(CStr(wbArrow.Worksheets(1).Range("A1").Value), Chr$(160), Chr$(32))) = "ARROW EUROPE Reporting : BACKLOG" Then
        wbArrow.Worksheets(1).Range("$A$2:$N$10000").AutoFilter Field:=1
        wbArrow.Worksheets(1).Range("$A$2:$N$10000").AutoFilter Field:=2
        wbArrow.Worksheets(1).Range("$A$2:$N$10000").AutoFilter Field:=3
        wbArrow.Worksheets(1).Range("$A$2:$N$10000").AutoFilter Field:=4
        wbArrow.Worksheets(1).Range("$A$2:$N$10000").AutoFilter Field:=5
        wbArrow.Worksheets(1).Range("$A$2:$N$10000").AutoFilter Field:=6
        wbArrow.Worksheets(1).Range("$A$2:$N$10000").AutoFilter Field:=7
        wbArrow.Worksheets(1).Range("$A$2:$N$10000").AutoFilter Field:=8
        wbArrow.Worksheets(1).Range("$A$2:$N$10000").AutoFilter Field:=9
        wbArrow.Worksheets(1).Range("$A$2:$N$10000").AutoFilter Field:=10
        wbArrow.Worksheets(1).Range("$A$2:$N$10000").AutoFilter Field:=11
        wbArrow.Worksheets(1).Range("$A$2:$N$10000").AutoFilter Field:=12
        wbArrow.Worksheets(1).Range("$A$2:$N$10000").AutoFilter Field:=13
        wbArrow.Worksheets(1).Range("$A$2:$N$10000").AutoFilter Field:=14
'======================================================================================================================================================'
        ' Clear "null" rows in Arrow excel file, they are not confirmed'
        Dim clear_null_values As Integer
        clear_null_values = 3
        Dim clear_null_string As String
'======================================================================================================================================================'
        ' Remember to clear non-break spaces generated by Arrow system for some unknown reason. '
        ' It is done below with Replace function. '
        For Each i In wbArrow.Worksheets(1).Range("$J$3:$J$10000")
            clear_null_string = CStr(i.Value)
            clear_null_string = Replace(clear_null_string, Chr$(160), Chr$(32))
            clear_null_string = Trim(clear_null_string)
            If clear_null_string = "null" Then
                wbArrow.Worksheets(1).Range(CStr(clear_null_values) & ":" & CStr(clear_null_values)).ClearContents
                clear_null_values = clear_null_values + 1
            Else
                clear_null_values = clear_null_values + 1
            End If
        Next i

        wbArrow.Worksheets(1).AutoFilter.Sort.SortFields.Clear
        wbArrow.Worksheets(1).AutoFilter.Sort.SortFields.Add Key:=Range("C2:C10000"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        With wbArrow.Worksheets(1).AutoFilter.Sort
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
'======================================================================================================================================================'
        Dim nr_of_orders As Integer
        nr_of_orders = 2
'======================================================================================================================================================'
        'Count number of orders'
        For Each Order In wbArrow.Worksheets(1).Range("C3:C10000")
            If Order.Value <> "" Then
                nr_of_orders = nr_of_orders + 1
            End If
            If Order.Value = "" Then
                Exit For
            End If
        Next Order
'======================================================================================================================================================'
        Dim end_of_orders As String
        end_of_orders = "C" + CStr(nr_of_orders)
        Dim next_order_count As Integer
        next_order_count = 4
        Dim next_order As String
        next_order = "C" + CStr(next_order_count)
        Dim current_order_count As Integer
        current_order_count = 3
        Dim current_order As String
        current_order = "C" + CStr(current_order_count)
        Dim ranges_of_orders(1 To 10000, 1 To 2) As String
        Dim single_range As Integer
        single_range = 1
'======================================================================================================================================================'
        'Get endings of orders'
        For Each Order In wbArrow.Worksheets(1).Range("C3:" & end_of_orders)
            If Order = wbArrow.Worksheets(1).Range(next_order).Value Then
                next_order_count = next_order_count + 1
                current_order_count = current_order_count + 1
                next_order = "C" + CStr(next_order_count)
                current_order = "C" + CStr(current_order_count)
            Else
                next_order_count = next_order_count + 1
                
                ranges_of_orders(single_range, 1) = "C" + CStr(current_order_count)
                
                single_range = single_range + 1
                current_order_count = current_order_count + 1

                next_order = "C" + CStr(next_order_count)
                current_order = "C" + CStr(current_order_count)
            End If
        Next Order

        single_range = 1
        current_order_count = 3
        next_order_count = 2
        current_order = "C" + CStr(current_order_count)
        next_order = "C" + CStr(next_order_count)

        'Get beginnings of orders'
        For Each Order In wbArrow.Worksheets(1).Range("C3:" & end_of_orders)
            If Order = wbArrow.Worksheets(1).Range(next_order).Value Then
                next_order_count = next_order_count + 1
                current_order_count = current_order_count + 1
                next_order = "C" + CStr(next_order_count)
                current_order = "C" + CStr(current_order_count)
            Else
                next_order_count = next_order_count + 1
                
                ranges_of_orders(single_range, 2) = "C" + CStr(current_order_count)
                
                single_range = single_range + 1
                current_order_count = current_order_count + 1

                next_order = "C" + CStr(next_order_count)
                current_order = "C" + CStr(current_order_count)
            End If
        Next Order
'======================================================================================================================================================'
        Dim combined_count As Integer
        combined_count = 1
        Dim combined_ranges_of_orders(1 To 10000) As String
        Dim every_second_row As Integer
        every_second_row = 0
        Dim row_counter As Integer
        Dim column_counter As Integer
'======================================================================================================================================================'
        'Combine two level array into one'
        For row_counter = 1 To 10000
            For column_counter = 1 To 2
                combined_ranges_of_orders(combined_count) = ranges_of_orders(row_counter, column_counter) + combined_ranges_of_orders(combined_count)
            Next column_counter
            combined_count = combined_count + 1
        Next row_counter
'======================================================================================================================================================'
        Dim clear_ranges_of_orders As Variant
        Dim clear_counter As Integer
        clear_counter = 1
'======================================================================================================================================================'
        'Clear combined_ranges_of_orders array from empty fields'
        ReDim clear_ranges_of_orders(LBound(combined_ranges_of_orders) To UBound(combined_ranges_of_orders))
        For i = LBound(combined_ranges_of_orders) To UBound(combined_ranges_of_orders)
            If combined_ranges_of_orders(i) <> "" Then
                j = j + 1
                clear_ranges_of_orders(j) = combined_ranges_of_orders(i)
            End If
        Next i
        ReDim Preserve clear_ranges_of_orders(LBound(combined_ranges_of_orders) To j)
'======================================================================================================================================================'
        Dim left_piece As String
        Dim right_piece As String
        Dim array_counter As Integer
        array_counter = 1
        Dim array_size_meter As Integer
        array_size_meter = 0
        ReDim order_numbers(1 To 1) As String
        Dim preserve_order_nr As Integer
        preserve_order_nr = 1
'======================================================================================================================================================'
        'Weld together ranges for orders'
        For Each i In clear_ranges_of_orders

            If Len(i) = 4 Then
                ReDim Preserve order_numbers(1 To preserve_order_nr)
                left_piece = Left(i, 2)
                right_piece = Right(i, 2)
                order_numbers(preserve_order_nr) = left_piece
                preserve_order_nr = preserve_order_nr + 1
            ElseIf Len(i) = 5 Then
                ReDim Preserve order_numbers(1 To preserve_order_nr)
                left_piece = Left(i, 2)
                right_piece = Right(i, 3)
                order_numbers(preserve_order_nr) = left_piece
                preserve_order_nr = preserve_order_nr + 1
            ElseIf Len(i) = 6 Then
                ReDim Preserve order_numbers(1 To preserve_order_nr)
                left_piece = Left(i, 3)
                right_piece = Right(i, 3)
                order_numbers(preserve_order_nr) = left_piece
                preserve_order_nr = preserve_order_nr + 1
            ElseIf Len(i) = 7 Then
                ReDim Preserve order_numbers(1 To preserve_order_nr)
                left_piece = Left(i, 3)
                right_piece = Right(i, 4)
                order_numbers(preserve_order_nr) = left_piece
                preserve_order_nr = preserve_order_nr + 1
            ElseIf Len(i) = 8 Then
                ReDim Preserve order_numbers(1 To preserve_order_nr)
                left_piece = Left(i, 4)
                right_piece = Right(i, 4)
                order_numbers(preserve_order_nr) = left_piece
                preserve_order_nr = preserve_order_nr + 1
            ElseIf Len(i) = 9 Then
                ReDim Preserve order_numbers(1 To preserve_order_nr)
                left_piece = Left(i, 4)
                right_piece = Right(i, 5)
                order_numbers(preserve_order_nr) = left_piece
                preserve_order_nr = preserve_order_nr + 1
            ElseIf Len(i) = 10 Then
                ReDim Preserve order_numbers(1 To preserve_order_nr)
                left_piece = Left(i, 5)
                right_piece = Right(i, 5)
                order_numbers(preserve_order_nr) = left_piece
                preserve_order_nr = preserve_order_nr + 1
            ElseIf Len(i) = 11 Then
                ReDim Preserve order_numbers(1 To preserve_order_nr)
                left_piece = Left(i, 5)
                right_piece = Right(i, 6)
                order_numbers(preserve_order_nr) = left_piece
                preserve_order_nr = preserve_order_nr + 1
            ElseIf Len(i) = 12 Then
                ReDim Preserve order_numbers(1 To preserve_order_nr)
                left_piece = Left(i, 6)
                right_piece = Right(i, 6)
                order_numbers(preserve_order_nr) = left_piece
                preserve_order_nr = preserve_order_nr + 1
            Else
                MsgBox "Too many rows, please contact file designer"
            End If

            clear_ranges_of_orders(array_counter) = left_piece + ":" + right_piece
            array_counter = array_counter + 1
            array_size_meter = array_size_meter + 1

        Next i
'======================================================================================================================================================'
        Dim row_number As String
        Dim column_number As String
        ReDim combined_pos_range(1 To 1) As String
        ReDim combined_rim_number_range(1 To 1) As String
        ReDim combined_confirmed_date_range(1 To 1) As String
        ReDim combined_confirmed_quantity_range(1 To 1) As String
        ReDim combined_confirmed_price_range(1 To 1) As String
        Dim preserve_counter As Integer
        preserve_counter = 1
        array_counter = 1
'======================================================================================================================================================'
        'Set ranges for RIMs'
        For Each i In clear_ranges_of_orders
            ReDim Preserve combined_pos_range(1 To preserve_counter)
            ReDim Preserve combined_rim_number_range(1 To preserve_counter)
            ReDim Preserve combined_confirmed_date_range(1 To preserve_counter)
            ReDim Preserve combined_confirmed_quantity_range(1 To preserve_counter)
            ReDim Preserve combined_confirmed_price_range(1 To preserve_counter)
            preserve_counter = preserve_counter + 1

            If Len(i) = 5 Then
                row_number = Left(i, 2)
                row_number = Right(row_number, 1)
                column_number = Right(i, 1)
            ElseIf Len(i) = 6 Then
                row_number = Left(i, 2)
                row_number = Right(row_number, 1)
                column_number = Right(i, 2)
            ElseIf Len(i) = 7 Then
                row_number = Left(i, 3)
                row_number = Right(row_number, 2)
                column_number = Right(i, 2)
            ElseIf Len(i) = 8 Then
                row_number = Left(i, 3)
                row_number = Right(row_number, 2)
                column_number = Right(i, 3)
            ElseIf Len(i) = 9 Then
                row_number = Left(i, 4)
                row_number = Right(row_number, 3)
                column_number = Right(i, 3)
            ElseIf Len(i) = 10 Then
                row_number = Left(i, 4)
                row_number = Right(row_number, 4)
                column_number = Right(i, 4)
            ElseIf Len(i) = 11 Then
                row_number = Left(i, 5)
                row_number = Right(row_number, 4)
                column_number = Right(i, 4)
            ElseIf Len(i) = 12 Then
                row_number = Left(i, 5)
                row_number = Right(row_number, 5)
                column_number = Right(i, 5)
            ElseIf Len(i) = 13 Then
                row_number = Left(i, 6)
                row_number = Right(row_number, 5)
                column_number = Right(i, 5)
            End If
            
            combined_pos_range(array_counter) = "N" + row_number + ":" + "N" + column_number
            combined_rim_number_range(array_counter) = "D" + row_number + ":" + "D" + column_number
            combined_confirmed_date_range(array_counter) = "J" + row_number + ":" + "J" + column_number
            combined_confirmed_quantity_range(array_counter) = "G" + row_number + ":" + "G" + column_number
            combined_confirmed_price_range(array_counter) = "M" + row_number + ":" + "M" + column_number
            array_counter = array_counter + 1

        Next i
'======================================================================================================================================================'
        Dim order_counter As Integer
        order_counter = 0
        Dim xml_nr As Integer
        xml_nr = 1
'======================================================================================================================================================'
        'Count number of orders
        For Each Order In order_numbers
            order_counter = order_counter + 1
        Next Order
'======================================================================================================================================================'
        ' Declarations for variables used in XML copy loop
        Dim how_many_rows_to_add As Integer
        Dim conf_rows_counter As Integer
        Dim conf_pos_range As String
        Dim conf_rim_range As String
        Dim conf_date_range As String
        Dim conf_quantity_range As String
        Dim conf_price_range As String
        Dim add_row_counter As Integer
        Dim file_name As String
        Dim file_path As String
'======================================================================================================================================================'
        ' Copy data from single order, save as XML, clear and move to next order
        For xml_nr = 1 To order_counter
            ' Copy order number and paste to Rimaster Excel
            wbRimaster.Worksheets(1).Range("C10").Value = Trim(Replace(CStr(wbArrow.Worksheets(1).Range(order_numbers(xml_nr)).Value), Chr$(160), Chr$(32)))

            ' Check if lines need to be added. One line is always there by default, that's why it's -1
            how_many_rows_to_add = -1
            For Each i In wbArrow.Worksheets(1).Range(combined_rim_number_range(xml_nr))
                how_many_rows_to_add = how_many_rows_to_add + 1
            Next i
            ' If lines need to be added, add them
            If how_many_rows_to_add > 0 Then
                For add_row_counter = 1 To how_many_rows_to_add
                    wbRimaster.Worksheets(1).Rows("31:31").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrBelow
                Next add_row_counter
            End If

            conf_rows_counter = 31
            conf_pos_range = "E" + CStr(conf_rows_counter)
            conf_rim_range = "F" + CStr(conf_rows_counter)
            conf_date_range = "Q" + CStr(conf_rows_counter)
            conf_quantity_range = "S" + CStr(conf_rows_counter)
            conf_price_range = "T" + CStr(conf_rows_counter)
            
            ' Copy all necessary data from Arrow Excel to Rimaster Order Confirmation
            ' Copy order lines, trim white spaces
            For Each i In wbArrow.Worksheets(1).Range(combined_pos_range(xml_nr))
                wbRimaster.Worksheets(1).Range(conf_pos_range).Value = Trim(Replace(CStr(i.Value), Chr$(160), Chr$(32)))
                conf_rows_counter = conf_rows_counter + 1
                conf_pos_range = "E" + CStr(conf_rows_counter)
            Next i
            conf_rows_counter = 31
            ' Copy RIM numbers
            For Each i In wbArrow.Worksheets(1).Range(combined_rim_number_range(xml_nr))
                wbRimaster.Worksheets(1).Range(conf_rim_range).Value = Trim(Replace(CStr(i.Value), Chr$(160), Chr$(32)))
                conf_rows_counter = conf_rows_counter + 1
                conf_rim_range = "F" + CStr(conf_rows_counter)
            Next i
            conf_rows_counter = 31
            ' Copy confirmed date
            For Each i In wbArrow.Worksheets(1).Range(combined_confirmed_date_range(xml_nr))
                wbRimaster.Worksheets(1).Range(conf_date_range).Value = Trim(Replace(CStr(i.Value), Chr$(160), Chr$(32)))
                conf_rows_counter = conf_rows_counter + 1
                conf_date_range = "Q" + CStr(conf_rows_counter)
            Next i
            conf_rows_counter = 31
            ' Copy confirmed quantity
            For Each i In wbArrow.Worksheets(1).Range(combined_confirmed_quantity_range(xml_nr))
                wbRimaster.Worksheets(1).Range(conf_quantity_range).Value = Trim(Replace(CStr(i.Value), Chr$(160), Chr$(32)))
                conf_rows_counter = conf_rows_counter + 1
                conf_quantity_range = "S" + CStr(conf_rows_counter)
            Next i
            conf_rows_counter = 31
            ' Copy confirmed price
            For Each i In wbArrow.Worksheets(1).Range(combined_confirmed_price_range(xml_nr))
                wbRimaster.Worksheets(1).Range(conf_price_range).Value = Trim(Replace(CStr(i.Value), Chr$(160), Chr$(32)))
                conf_rows_counter = conf_rows_counter + 1
                conf_price_range = "T" + CStr(conf_rows_counter)
            Next i

            conf_rows_counter = conf_rows_counter - 1
            conf_price_range = "T" + CStr(conf_rows_counter)

            ' Make sure optional data have no empty cells
            If wbRimaster.Worksheets(1).Range("C12").Value = "" Then
                wbRimaster.Worksheets(1).Range("C12").Value = "'"
            End If
            If wbRimaster.Worksheets(1).Range("C18").Value = "" Then
                wbRimaster.Worksheets(1).Range("C18").Value = "'"
            End If
            If wbRimaster.Worksheets(1).Range("C19").Value = "" Then
                wbRimaster.Worksheets(1).Range("C19").Value = "'"
            End If
            If wbRimaster.Worksheets(1).Range("C20").Value = "" Then
                wbRimaster.Worksheets(1).Range("C20").Value = "'"
            End If
            If wbRimaster.Worksheets(1).Range("C21").Value = "" Then
                wbRimaster.Worksheets(1).Range("C21").Value = "'"
            End If
            If wbRimaster.Worksheets(1).Range("C26").Value = "" Then
                wbRimaster.Worksheets(1).Range("C26").Value = "'"
            End If

            ' Make sure data in hidden columns are set properly
            For Each i In wbRimaster.Worksheets(1).Range("G31:G" & CStr(conf_rows_counter))
                i.Value = 5
            Next i
            For Each i In wbRimaster.Worksheets(1).Range("J31:K" & CStr(conf_rows_counter))
                i.Value = "'"
            Next i
            For Each i In wbRimaster.Worksheets(1).Range("P31:P" & CStr(conf_rows_counter))
                i.Value = "'"
            Next i
            For Each i In wbRimaster.Worksheets(1).Range("R31:R" & CStr(conf_rows_counter))
                i.Value = "'"
            Next i
            For Each i In wbRimaster.Worksheets(1).Range("L31:N" & CStr(conf_rows_counter))
                i.Value = 0
            Next i
            For Each i In wbRimaster.Worksheets(1).Range("O31:O" & CStr(conf_rows_counter))
                i.Value = 1
            Next i

            ' Declare order number to be used in name
            file_name = CStr(wbRimaster.Worksheets(1).Range("C10").Value)

            ' Check if folder on desktop exist if not create it
            If Dir(Environ("USERPROFILE") & "\Desktop\Arrow XML\", vbDirectory) = "" Then
                MkDir Environ("USERPROFILE") & "\Desktop\Arrow XML\"
            End If

            ' Declare file path
            file_path = Environ("USERPROFILE") & "\Desktop\Arrow XML\XML Order nr" & file_name & ".xml"
            
            ' Check if file exists, if so delete and save new
            If Dir(file_path, vbDirectory) <> "" Then
                ' Remove readonly
                SetAttr file_path, vbNormal
                ' Delete file
                Kill file_path
                ' Save new
                wbRimaster.XmlMaps("ORDRSP419_mapa").Export Url:=file_path
            Else
                wbRimaster.XmlMaps("ORDRSP419_mapa").Export Url:=file_path
            End If

            ' Clear extra fields, leave only one and clear values
            conf_rows_counter = conf_rows_counter + 1
            wbRimaster.Worksheets(1).Range("32:" & CStr(conf_rows_counter)).Delete Shift:=xlUp
            wbRimaster.Worksheets(1).Range("E31").Value = "'"
            wbRimaster.Worksheets(1).Range("F31").Value = "'"
            wbRimaster.Worksheets(1).Range("H31").Value = "'"
            wbRimaster.Worksheets(1).Range("I31").Value = "'"
            wbRimaster.Worksheets(1).Range("Q31").Value = "'"
            wbRimaster.Worksheets(1).Range("S31").Value = "'"
            wbRimaster.Worksheets(1).Range("T31").Value = "'"
            wbRimaster.Worksheets(1).Range("C10").Value = "'"
            wbRimaster.Worksheets(1).Range("C12").Value = "'"
            wbRimaster.Worksheets(1).Range("C18:C21").Value = "'"
            wbRimaster.Worksheets(1).Range("C26").Value = "'"

        Next xml_nr
        ' After all files are saved close Arrow excel
        wbArrow.Close SaveChanges:=False
    Else
        ' If other file is picked show error msg and close the file
        MsgBox "Incorrect file, please try again with Arrow Excel."
        wbArrow.Close SaveChanges:=False
    End If

End Sub
