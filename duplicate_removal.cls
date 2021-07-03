VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub remove_continous_value_Click()
    Dim DataRange As Range
    
    Dim m As Integer
    Dim n As Integer
    Dim x As Double
    
    Set DataRange = Selection
    m = DataRange.Rows.count
    n = DataRange.Columns.count
    x = 0
    
    For j = 1 To n
        If VarType(DataRange.Cells(2, j)) = vbDate Or VarType(DataRange.Cells(2, j)) = vbInteger _
        Or VarType(DataRange.Cells(2, j)) = vbLong Or VarType(DataRange.Cells(2, j)) = vbCurrency _
        Or VarType(DataRange.Cells(2, j)) = vbLongLong Or VarType(DataRange.Cells(2, j)) = vbDouble _
        And DataRange.Cells(2, j).EntireColumn.Hidden = False Then

        
        
            
            x = 0
            For i = 2 To m - 1
                If Not DataRange.Cells(i, j) + 1 = DataRange.Cells(i + 1, j) Then
                    x = 1
                End If
            Next i
            
            If x = 0 Then
                x = Range("B2").Value
                Range("B" & x).Value = DataRange.Cells(2, j).Value
                Range("A" & x).Value = DataRange.Cells(1, j).Value
                Range("A" & x).Interior.ColorIndex = 20
                Range("B2").Cells(1, 1).Value = x + 1
                DataRange.Columns(j).EntireColumn.Hidden = True
            End If
        End If
    Next j
    
    
    x = Range("B1").Cells(1, 1).Value
    
    For i = 1 To x
        Cells.Rows(i).EntireRow.Hidden = False
    Next i
            
            
            
    
End Sub

Private Sub hide_duplicate_rows_Click()
    Dim DataRange As Range
    
    Dim m As Integer
    Dim n As Integer
    Dim x As Double
    
    Set DataRange = Selection
    m = DataRange.Rows.count
    n = DataRange.Columns.count
    x = 0

    For i = 1 To m - 1

        x = 0
        
        For j = 1 To n

            If Not StrComp(DataRange.Cells(i, j), DataRange.Cells(i + 1, j), vbBinaryCompare) = 0 Then
        
                x = 1

            End If

        Next j
        If x = 0 Then
            DataRange.Rows(i).EntireRow.Hidden = True
        End If

    Next i
    
    x = Range("B1").Cells(1, 1).Value
    
    For i = 1 To x
        Cells.Rows(i).EntireRow.Hidden = False
    Next i

End Sub

Private Sub say_helloworld_Click()
    Dim DataRange As Range
    
    Dim m As Integer
    Dim n As Integer
    Dim x As Double
    
    Set DataRange = Selection
    m = DataRange.Rows.count
    n = DataRange.Columns.count
    x = 0
    
    
    For i = 1 To n


            If (Application.WorksheetFunction.CountIf(DataRange.Columns(i), DataRange.Cells(2, i))) = m - 1 _
            And m > 1 _
            And DataRange.Cells(2, i).EntireColumn.Hidden = False Then
                DataRange.Columns(i).Hidden = True
                x = Range("B2").Value
                Range("B" & x).Value = DataRange.Cells(2, i).Value
                Range("A" & x).Value = DataRange.Cells(1, i).Value
                Range("B2").Cells(1, 1).Value = x + 1
                
            End If
            
        
    Next i
    
    x = Range("B2").Cells(1, 1).Value
    
    For i = 1 To x
        Cells.Rows(i).EntireRow.Hidden = False
    Next i
    
    'MsgBox x
End Sub

Private Sub sheet_reset_Click()
    Dim x As Integer
    x = Range("B1").Cells(1, 1).Value
    Range("A:B").Value = ""
    Range("1:2").Value = ""
    Range("A1").Value = "Starts From"
    Range("B1").Value = x
    Range("A2").Value = "Count"
    Range("B2").Value = x + 1
    Cells.Columns.EntireColumn.Hidden = False
    Cells.Rows.EntireRow.Hidden = False
    Cells.Interior.ColorIndex = xlColorIndexNone
    ActiveWindow.DisplayGridlines = True
    ActiveWindow.DisplayGridlines = True
    ActiveWindow.GridlineColor = RGB(0, 0, 0)
End Sub
