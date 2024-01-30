Attribute VB_Name = "InputCorrelations"
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
    Dim columnIndex As Integer
    
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
    
    Dim LastContiguousRow As Integer
    LastContiguousRow = startRow
    
    While Not IsEmpty(ws.Cells(LastContiguousRow + 1, 1))
        LastContiguousRow = LastContiguousRow + 1
    Wend
    'When I dont' know beforehand how many columns contain data.
    For columnIndex = 3 To lastContiguousColumn
        Dim headerValue As String
        headerValue = ws.Cells(columnnameRow, columnIndex).Value
        
        For rowIndex = startRow To LastContiguousRow
            If ws.Cells(rowIndex, 1).Value = headerValue Then
                ws.Cells(rowIndex, columnIndex).Value = 1
            End If
        Next rowIndex
    Next columnIndex
    
    Set jsonResponse = JsonConverter.ParseJson(JsonString)
    ' ... [earlier code remains the same]

    ' Extract the data_get_1 array from the JSON response
    Dim selCorrelation As Collection
    Set selCorrelation = jsonResponse("selCorrelation")
    
    For columnIndex = 3 To lastContiguousColumn '수정해야됨. End(xlToLeft)는 끝까지 갔다가 돌아오는 거여서 Dynamic table 만들 시, 사용하면 안됨.
        
        headerValue = ws.Cells(columnnameRow, columnIndex).Value
        
        For rowIndex = startRow To LastContiguousRow
            Dim cellValue As String
            cellValue = ws.Cells(rowIndex, 1).Value
            For i = 1 To selCorrelation.Count
                Dim data As Variant
                data = selCorrelation(i)("data")
                 
                Dim dataParts As Variant
                dataParts = Split(data, "|")
                                 
                If (cellValue = dataParts(4) And headerValue = dataParts(5)) Or _
                (cellValue = dataParts(5) And headerValue = dataParts(4)) Then
                    ws.Cells(rowIndex, columnIndex).Value = dataParts(3)
                End If

            Next i
            
        Next rowIndex
    Next columnIndex
    
    Dim FXRow As Integer
    FXRow = ws.Columns(1).Find(What:="FX", LookIn:=xlValues, LookAt:=xlWhole).Row
    
    Dim FXmarker As Integer
    FXmarker = FXRow + 4
    
    Dim ColumnNameRow2 As Integer
    ColumnNameRow2 = FXRow + 3
    
    Dim ColumnIndex2 As Integer
    ColumnIndex2 = 4
    
    Dim LastContiguousRow2 As Integer
    LastContiguousRow2 = FXmarker
    
    While Not IsEmpty(ws.Cells(LastContiguousRow2 + 1, 1))
        LastContiguousRow2 = LastContiguousRow2 + 1
    Wend
    
    Dim LastContiguousColumn2 As Integer
    LastContiguousColumn2 = ColumnIndex2
    
    While Not IsEmpty(ws.Cells(ColumnNameRow2, LastContiguousColumn2 + 1))
        LastContiguousColumn2 = LastContiguousColumn2 + 1
    Wend
    
    Dim RowIndex2 As Integer
    
    For ColumnIndex2 = 4 To LastContiguousColumn2
        Dim hheader2 As String
        hheader2 = ws.Cells(ColumnNameRow2, ColumnIndex2).Value
        For RowIndex2 = FXmarker To LastContiguousRow2
            Dim vheader2 As String
            vheader2 = ws.Cells(RowIndex2, 1)
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
    
End Sub
