Attribute VB_Name = "Module4"
Sub PostClosedPrice()
    Dim DataString As String
    Dim xmlhttp As Object
    Dim i As Integer
    Dim baseDt As String
    Dim dataSetId As String
    Dim StartingPoint As String
    Dim dataId As String
    Dim closePric As String
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Market Data")
    Dim targetDate As Date
    ' Retrieve the base date and data set ID from the worksheet
    targetDate = Sheets("Market Data").Range("A2").Value
    
    baseDt = Format(targetDate, "yyyymmdd")
    dataSetId = Sheets("Market Data").Range("O2").Value
    StartingPoint = Sheets("Market Data").Range("P2").Value
    
    Dim Table1Point As Range
    Set Table1Point = Sheets("Market Data").Range(StartingPoint).Offset(3, 0)
    
    ' Find the cell that contains "FX" after "Equity" table
    Dim fxRow As Range
    Set fxRow = ws.Range(Table1Point.Offset(1, 0), ws.Cells(ws.Rows.Count, Table1Point.Column)).Find(What:="FX", LookIn:=xlValues, LookAt:=xlWhole)
    Debug.Print fxRow.Value
    Debug.Print Table1Point.Value
    Debug.Print fxRow.Row
    Debug.Print Table1Point.Row
    ' Initialize the DataString
    DataString = ""
    
    For i = Table1Point.Row + 1 To fxRow.Row - 2
        If Len(DataString) > 0 Then
            DataString = DataString & "&"
        End If
        
        dataId = ws.Cells(i, Table1Point.Column).Value
        closePric = ws.Cells(i, Table1Point.Column + 1).Value
        'Construct the String
        DataString = DataString & "BASE_DT=" & baseDt & _
                     "&DATA_SET_ID=" & dataSetId & _
                     "&DATA_ID=" & dataId & _
                     "&CLOSE_PRIC=" & closePric & _
                     "&PGM_ID=TEST" & _
                     "&WRKR_ID=HS" & _
                     "&WORK_TRIP=0.0.0.0"
    Next i
    Debug.Print DataString
    
    
    
End Sub
