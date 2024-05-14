Attribute VB_Name = "Module1"
Option Explicit

Sub ①作成()
Dim strIn As String, price, i As Long, AlRow As Long, ElRow As Long
    Columns("A:A").ColumnWidth = 25.42
    Range("A5:A6").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("B5:B6").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("E2:F2").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("E3:F3").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("E4:F4").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("H2:I2").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("H3:I3").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("H4:I4").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("D1:I1").Select
    Range("I1").Activate
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("C5:G5").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("H5:I5").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Rows("1:1").RowHeight = 36
    Rows("2:2").RowHeight = 27
    Rows("3:3").RowHeight = 27
    Rows("4:4").RowHeight = 27
    
    strIn = InputBox("商品名を入力してください")
    price = InputBox("金額を入力して下さい")
    
    'MsgBox strIn & "、" & price & "円です"
    
    ActiveSheet.Name = strIn
    
    Cells(2, 4) = "価格(税抜)"
    Cells(2, 7) = "価格(税込)"
    Cells(3, 4) = "季節"
    Cells(3, 7) = "種類"
    Cells(4, 4) = "原価率"
    Cells(4, 7) = "原価"
    Cells(5, 1) = "材料名"
    Cells(5, 2) = "取引業者"
    Cells(5, 3) = "仕入れ容量"
    Cells(6, 3) = "容量"
    Cells(6, 4) = "単位"
    Cells(6, 5) = "歩留率"
    Cells(6, 6) = "単価"
    Cells(6, 7) = "単価(歩留)"
    Cells(6, 8) = "容量"
    Cells(6, 9) = "単価"
    Cells(5, 8) = "1皿当たりの使用料"
    Cells(1, 4) = strIn
    Cells(2, 5) = price
    Cells(2, 8) = price * 1.1
    'Sheets("仕入れ").Cells(1, 7) = price
    
    AlRow = Sheets("仕入れ").Cells(Rows.Count, "A").End(xlUp).Row  'A列の最終行を取得
    ElRow = Sheets("仕入れ").Cells(Rows.Count, "E").End(xlUp).Row
    
    With Range("A2").EntireColumn.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, _
        Operator:=xlBetween, Formula1:="=仕入れ!$A$2:$A$" & AlRow
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = "リストに登録されていません。"
        .IMEMode = xlIMEModeNoControl
        .ShowInput = True
        .ShowError = True
    End With
    
    Witsh Sheets("仕入れ").Range("C2").EntireColumn.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, _
        Operator:=xlBetween, Formula1:="=仕入れ!$H$2:$H$" & ElRow
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = "リストに登録されていません。"
        .IMEMode = xlIMEModeNoControl
        .ShowInput = True
        .ShowError = True
    End With
    
    With Range("D7").EntireColumn.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, _
        Operator:=xlBetween, Formula1:="=仕入れ!$H$2:$H$" & ElRow
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = "リストに登録されていません。"
        .IMEMode = xlIMEModeNoControl
        .ShowInput = True
        .ShowError = True
    End With

End Sub
