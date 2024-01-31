Attribute VB_Name = "Functions"
Public Function GetJsonResponse(url As String) As Object
    ' Variables to hold the HTTP request and response data
    Dim httpRequest As Object
    Dim JsonString As String
    Dim jsonResponse As Object

    ' Create the HTTP request
    Set httpRequest = CreateObject("MSXML2.XMLHTTP")
    With httpRequest
        .Open "GET", url, False
        .Send
        JsonString = .responseText
    End With

    ' Parse the JSON response
    Set jsonResponse = JsonConverter.ParseJson(JsonString)

    ' Return the parsed JSON response
    Set GetJsonResponse = jsonResponse
End Function

'셀의 범위를 변수로 받아서 (header포함), header의 값(vertical, horizontal)과 지표가 일치하는 경우 해당 corr값을 넣어주는 함수입니다.
Function UpdateCellsWithCorrelation(ws As Worksheet, selCorrelation As Collection, _
                                    ColumnNameRow As Integer, FXmarker As Integer, _
                                    LastContiguousRow As Integer, LastContiguousColumn As Integer, ColumnIndex As Integer) As Boolean
    Dim ColumnIndex2 As Integer
    Dim RowIndex2 As Integer

    For ColumnIndex2 = ColumnIndex To LastContiguousColumn
        Dim hheader2 As String
        hheader2 = ws.Cells(ColumnNameRow, ColumnIndex2).Value
        For RowIndex2 = FXmarker To LastContiguousRow
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


