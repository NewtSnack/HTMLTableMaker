Attribute VB_Name = "HTMLTableMaker"
'DARRON CHARLES 2022
Private Sub Start() 'Paste Table from HTML into A1 and Start Here
    TableSortAlgo
    HTMLTableMaker
    Cells(1, 1).Select
End Sub
Private Sub HTMLTableMaker()
'makes a X column table in alphabetical order going across columns left to right, column count determined by first row, no other rows should have a column count larger that the first row
Dim tableColumns As Integer
Dim rowHtml, empName, empNumber, HTML As String
tableColumns = range("A1", range("A1").End(xlToRight)).Count
currRow = 0
With Application.ActiveSheet
    .Cells(1, 1).Select
    While Not checkEmpty(.range("A" & currRow + 1 & ":D" & currRow + 1), tableColumns)
        currRow = currRow + 1
        HTML = HTML & "<tr> <!--" & currRow & "-->" & vbNewLine
        For currCol = 1 To tableColumns
            .Cells(currRow, currCol).Select
            If Not IsEmpty(Selection.Value) Then
                'HTML here change to what ever you need
                empNumber = Right(Selection.Value, 3)
                empName = Left(Selection.Value, Len(Selection.Value) - 3)
                rowHtml = vbTab & "<td><a class=""names"">" & empName & "</a><a class=""numbers"">" & empNumber & "</a></td>" & vbNewLine
                HTML = HTML + rowHtml
            End If
        Next
        HTML = HTML & "</tr>" & vbNewLine
    Wend
    .Cells(currRow + 1, tableColumns + 1).Value = HTML
End With
MsgBox "When copy pasting the HTML, copy the literal value of the cell or else there will be double quotes.", vbInformation
End Sub
Function checkEmpty(range As range, tblcol As Integer) As Boolean
Dim i As Integer
i = 0
For Each cell In range
    If IsEmpty(cell) Then
        i = i + 1
    End If
Next
If i = tblcol Then
    checkEmpty = True
Else
    checkEmpty = False
End If
End Function
Private Sub TableSortAlgo()
'takes an usorted table, tranforms it into a list, sorts, then transforms data back into a sorted table, alphabetically down columns and to the right.
'columns lengths are a function of number of items and the number of columns desired.
Dim xlrow As Integer
Dim cellcount As Integer
'List
tableColumns = range("A1", range("A1").End(xlToRight)).Count

With Application.ActiveSheet
    For col = 2 To tableColumns
        .Cells(1, col).Select
        collength = Columns(col).Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
        .range(.Cells(1, col), .Cells(collength, col)).Select
        xlrow = range(Cells(1, 1), Cells(ActiveSheet.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row, 1)).Count
        Selection.Cut range("A" & xlrow + 1)
    Next
End With
'Sort
cellcount = range("A:A").Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
range(Cells(1, 1), Cells(cellcount, 1)).Select
ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Add2 Key:=range("A1:A49") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Sheet1").Sort
       .SetRange range("A1:A" & cellcount)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
'Table Algo
cellcount = range("A:A").Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
range(Cells(1, 1), Cells(cellcount, 1)).Select
cellcount = Selection.Count
colheight = WorksheetFunction.Ceiling(cellcount / tableColumns, 1)

For j = 0 To WorksheetFunction.Ceiling(cellcount / colheight, 1)
    range("A" & colheight * j + 1, "A" & colheight * (j + 1)).Select
    Selection.Cut Cells(1, 1 + j)
    Next
End Sub


