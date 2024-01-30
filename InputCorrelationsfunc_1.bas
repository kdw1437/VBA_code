Attribute VB_Name = "InputCorrelationsfunc"
'Correlation값을 칼럼에 맞춰 다이나믹하게 넣어주는 코드
Sub InputCorrelations()
    ' Variables to hold the HTTP request and response data
    Dim httpRequest As Object
    Dim jsonResponse As Object
    Dim JsonString As String
    
    ' Assuming you have a worksheet variable set to the target sheet
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Market Data") ' Change to your actual sheet name

    ' Retrieve the date value from cell A2 and format it as 'yyyymmdd'
    Dim targetDate As Date
    targetDate = ws.Range("A2").Value
    Dim dateParameter As String
    dateParameter = Format(targetDate, "yyyymmdd")

    ' Construct the full URL with the formatted date parameter
    Dim baseURL As String
    Dim url As String
    baseURL = "http://localhost:8080/val/v1/Correlations/official?basedt="
    url = baseURL & dateParameter
    
    ' Create the HTTP request
    Set httpRequest = CreateObject("MSXML2.XMLHTTP")
    With httpRequest
        .Open "GET", url, False
        .Send
        JsonString = .responseText
    End With
    
    Dim rowIndex As Integer
    Dim ColumnIndex As Integer
    
    Dim lastContiguousColumn As Integer
    lastContiguousColumn = 3 ' Start from column 3
    
    Dim equityRow As Integer
    equityRow = ws.Columns(1).Find(What:="Equity", LookIn:=xlValues, LookAt:=xlPart).Row
    
    ' Starting row for writing data is 4 rows below 'Equity'
    Dim startRow As Integer
    startRow = equityRow + 4
    Dim columnnameRow As Integer
    columnnameRow = equityRow + 3
    
    While Not IsEmpty(ws.Cells(columnnameRow, lastContiguousColumn + 1))
        lastContiguousColumn = lastContiguousColumn + 1
    Wend
    
    Dim lastContiguousRow As Integer
    lastContiguousRow = startRow
    
    While Not IsEmpty(ws.Cells(lastContiguousRow + 1, 1))
        lastContiguousRow = lastContiguousRow + 1
    Wend
    'When I dont' know beforehand how many columns contain data.
    For ColumnIndex = 3 To lastContiguousColumn
        Dim headerValue As String
        headerValue = ws.Cells(columnnameRow, ColumnIndex).Value
        
        For rowIndex = startRow To lastContiguousRow
            If ws.Cells(rowIndex, 1).Value = headerValue Then
                ws.Cells(rowIndex, ColumnIndex).Value = 1
            End If
        Next rowIndex
    Next ColumnIndex

    ' Parse the JSON response
    Set jsonResponse = JsonConverter.ParseJson(JsonString)

    ' Extract the correlation data from the JSON response
    Dim selCorrelation As Collection
    Set selCorrelation = jsonResponse("selCorrelation")
    
    ' Update the worksheet with correlations for 'Equity'
    Call UpdateCellsWithCorrelation(ws, selCorrelation, columnnameRow, startRow, lastContiguousRow, lastContiguousColumn, 3)

    ' Define the start row and column name row based on 'FX'
    Dim FXRow As Integer
    FXRow = ws.Columns(1).Find(What:="FX", LookIn:=xlValues, LookAt:=xlWhole).Row
    
    Dim FXmarker As Integer
    FXmarker = FXRow + 4
    
    Dim ColumnNameRow2 As Integer
    ColumnNameRow2 = FXRow + 3

    Dim LastContiguousRow2 As Integer
    LastContiguousRow2 = FXmarker
    
    Dim ColumnIndex2 As Integer
    ColumnIndex2 = 4
    
    While Not IsEmpty(ws.Cells(LastContiguousRow2 + 1, 1))
        LastContiguousRow2 = LastContiguousRow2 + 1
    Wend
    
    Dim LastContiguousColumn2 As Integer
    LastContiguousColumn2 = ColumnIndex2
    
    While Not IsEmpty(ws.Cells(ColumnNameRow2, LastContiguousColumn2 + 1))
        LastContiguousColumn2 = LastContiguousColumn2 + 1
    Wend

    ' Update the worksheet with correlations for 'FX'
    Call UpdateCellsWithCorrelation(ws, selCorrelation, ColumnNameRow2, FXmarker, LastContiguousRow2, LastContiguousColumn2, 4)
End Sub


Function UpdateCellsWithCorrelation(ws As Worksheet, selCorrelation As Collection, _
                                    ColumnNameRow2 As Integer, FXmarker As Integer, _
                                    LastContiguousRow2 As Integer, LastContiguousColumn2 As Integer, ColumnIndex As Integer) As Boolean
    Dim ColumnIndex2 As Integer
    Dim RowIndex2 As Integer

    For ColumnIndex2 = ColumnIndex To LastContiguousColumn2
        Dim hheader2 As String
        hheader2 = ws.Cells(ColumnNameRow2, ColumnIndex2).Value
        For RowIndex2 = FXmarker To LastContiguousRow2
            Dim vheader2 As String
            vheader2 = ws.Cells(RowIndex2, 1).Value
            For i = 1 To selCorrelation.Count
                Dim data2 As Variant
                data2 = selCorrelation(i)("data")
                 
                Dim dataParts2 As Variant
                dataParts2 = Split(data2, "|")
                                 
                If (vheader2 = dataParts2(4) And hheader2 = dataParts2(5)) Or _
                   (vheader2 = dataParts2(5) And hheader2 = dataParts2(4)) Then
                    ws.Cells(RowIndex2, ColumnIndex2).Value = dataParts2(3)
                End If
            Next i
        Next RowIndex2
    Next ColumnIndex2
    
    UpdateCellsWithCorrelation = True
End Function

