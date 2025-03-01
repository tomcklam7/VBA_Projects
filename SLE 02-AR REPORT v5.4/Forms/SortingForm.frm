Private Sub UserForm_Initialize()
    ' Define the column options
    Dim columnOptions As Variant
    columnOptions = Array("A", "D", "E", "F", "G", "H", "I", "J", "M", "P", "Q", "R", "S", "Z", "AB", "AC", "AD", "AF", "AG")
    
    ' Populate all comboboxes with the same options
    Dim i As Long
    For i = 0 To UBound(columnOptions)
        ComboBox1.AddItem columnOptions(i)
        ComboBox2.AddItem columnOptions(i)
        ComboBox3.AddItem columnOptions(i)
    Next i
    
    ' Set default radio button selections to Ascending
    OptionButton1_Asc = True
    OptionButton2_Asc = True
    OptionButton3_Asc = True
End Sub

Private Sub Button_ApplySort_Click()
    ' Validate selections
    If ComboBox1.Value = "" Then
        MsgBox "Please select a value for the first sorting level!", vbExclamation
        Exit Sub
    End If
    
    If ComboBox3.Value <> "" And ComboBox2.Value = "" Then
        MsgBox "Please select a value for the second level before selecting the third level!", vbExclamation
        Exit Sub
    End If
    
    ' Get the last row
    Dim ws As Worksheet
    Set ws = ActiveSheet
    Dim lastRow As Long
    lastRow = ws.Range("A:A").Find("ENDOFROW").Row - 1
    
    ' Unprotect sheet
    ws.Unprotect Password:="" ' Replace with your password
    
    ' Create sort object
    With ws.Sort
        .SortFields.Clear
        
        ' First level sort
        .SortFields.Add Key:=Range(ComboBox1.Value & "8:" & ComboBox1.Value & lastRow), _
            SortOn:=xlSortOnValues, _
            Order:=IIf(OptionButton1_Asc.Value, xlAscending, xlDescending), _
            DataOption:=xlSortNormal
            
        ' Second level sort (if selected)
        If ComboBox2.Value <> "" Then
            .SortFields.Add Key:=Range(ComboBox2.Value & "8:" & ComboBox2.Value & lastRow), _
                SortOn:=xlSortOnValues, _
                Order:=IIf(OptionButton2_Asc.Value, xlAscending, xlDescending), _
                DataOption:=xlSortNormal
        End If
        
        ' Third level sort (if selected)
        If ComboBox3.Value <> "" Then
            .SortFields.Add Key:=Range(ComboBox3.Value & "8:" & ComboBox3.Value & lastRow), _
                SortOn:=xlSortOnValues, _
                Order:=IIf(OptionButton3_Asc.Value, xlAscending, xlDescending), _
                DataOption:=xlSortNormal
        End If
        
        ' Set sort range and execute
        .SetRange Range("A8:AG" & lastRow)
        .Header = xlNo
        .Apply
    End With
    
    ' Protect sheet
    ws.Protect Password:="" ' Replace with your password
    
    ' Close the form
    Unload Me
End Sub


===========================================================

Begin UserForm
    Caption = "Multi-Level Sort"
    Width = 400
    Height = 200
    
    ' First Level
    ComboBox1: Left = 20, Top = 20, Width = 120
    Frame1: Left = 150, Top = 10, Width = 200, Height = 35, Caption = "Sort Order"
        OptionButton1_Asc: Left = 10, Top = 15, Caption = "Ascending"
        OptionButton1_Desc: Left = 100, Top = 15, Caption = "Descending"
    
    ' Second Level
    ComboBox2: Left = 20, Top = 60, Width = 120
    Frame2: Left = 150, Top = 50, Width = 200, Height = 35, Caption = "Sort Order"
        OptionButton2_Asc: Left = 10, Top = 15, Caption = "Ascending"
        OptionButton2_Desc: Left = 100, Top = 15, Caption = "Descending"
    
    ' Third Level
    ComboBox3: Left = 20, Top = 100, Width = 120
    Frame3: Left = 150, Top = 90, Width = 200, Height = 35, Caption = "Sort Order"
        OptionButton3_Asc: Left = 10, Top = 15, Caption = "Ascending"
        OptionButton3_Desc: Left = 100, Top = 15, Caption = "Descending"
    
    ' Apply Button
    CommandButton1: Left = 150, Top = 140, Caption = "Apply Sort"
End UserForm

===========================================================