Attribute VB_Name = "Module2"
Sub UpdateClosePrice()
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
    baseURL = "http://localhost:8080/val/Get_data_1?basedt="
    url = baseURL & dateParameter & "&datasetid=official"

    
    ' Create the HTTP request
    Set httpRequest = CreateObject("MSXML2.XMLHTTP")
    With httpRequest
        .Open "GET", url, False
        .Send
        JsonString = .responseText
    End With
    
    Set jsonResponse = JsonConverter.ParseJson(JsonString)
    ' ... [earlier code remains the same]

    ' Extract the data_get_1 array from the JSON response
    Dim dataGet1 As Collection
    Set dataGet1 = jsonResponse("data_get_1")
    
    ' Find the row with 'Equity' in column A
    Dim equityRow As Integer
    equityRow = ws.Columns(1).Find(What:="Equity", LookIn:=xlValues, LookAt:=xlPart).Row
    
    ' Starting row for writing data is 4 rows below 'Equity'
    Dim startRow As Integer
    startRow = equityRow + 4
    Dim columnnameRow As Integer
    columnnameRow = equityRow + 3
    
    ws.Cells(columnnameRow, 1).Value = "Code"
    ws.Cells(columnnameRow, 2).Value = "ClosedPrice"
    ' Loop over each item in the data_get_1 array
    Dim item As Dictionary
    Dim currentRow As Integer
    currentRow = startRow
    
    For Each item In dataGet1
        ' Write dataId to column A and closePric to column B
        ws.Cells(currentRow, 1).Value = item("dataId")
        ws.Cells(currentRow, 2).Value = item("closePric")
    
        ' Increment row counter
        currentRow = currentRow + 1
    Next item
    'row에 가로로 이동하면서 dataId key의 value값을 넣어준다.
    
    Dim currentColumn As Integer
    currentColumn = 3
    For Each item In dataGet1
        ws.Cells(columnnameRow, currentColumn).Value = item("dataId")
        currentColumn = currentColumn + 1
    Next item
    ' ... [rest of your code]
    
        ' Compare and fill in 1 if the index matches
    Dim rowIndex As Integer
    Dim columnIndex As Integer
    
    Dim lastContiguousColumn As Integer
    lastContiguousColumn = 3 ' Start from column 3
    
    ' Check if there is data in the next column. If so, move one column to the right.
    While Not IsEmpty(ws.Cells(columnnameRow, lastContiguousColumn + 1))
        lastContiguousColumn = lastContiguousColumn + 1
    Wend
    
    Dim lastContiguousRow As Integer
    lastContiguousRow = startRow
    
    While Not IsEmpty(ws.Cells(lastContiguousRow + 1, 1))
        lastContiguousRow = lastContiguousRow + 1
    Wend
    'When I dont' know beforehand how many columns contain data.
    For columnIndex = 3 To lastContiguousColumn
        Dim headerValue As String
        headerValue = ws.Cells(columnnameRow, columnIndex).Value
        
        For rowIndex = startRow To lastContiguousRow
            If ws.Cells(rowIndex, 1).Value = headerValue Then
                ws.Cells(rowIndex, columnIndex).Value = 1
            End If
        Next rowIndex
    Next columnIndex
    '여기서 부터 추가 코드 작성 (Corrleation matrix 넣어주기)
    Dim dataGet2 As Collection
    Set dataGet2 = jsonResponse("data_get_2")
    

    For columnIndex = 3 To lastContiguousColumn '수정해야됨. End(xlToLeft)는 끝까지 갔다가 돌아오는 거여서 Dynamic table 만들 시, 사용하면 안됨.
        
        headerValue = ws.Cells(columnnameRow, columnIndex).Value
        
        For rowIndex = startRow To lastContiguousRow
            Dim cellValue As String
            cellValue = ws.Cells(rowIndex, 1).Value
            
            For Each item In dataGet2
                If (cellValue = item("th01DataId") And headerValue = item("th02DataId")) Or _
               (cellValue = item("th02DataId") And headerValue = item("th01DataId")) Then
                    ws.Cells(rowIndex, columnIndex).Value = item("crltCfcn")
                End If
            Next item
        Next rowIndex
    Next columnIndex
    '증감된 currentRow variable을 이용해서, 다음 값들을 넣어줌.
    
    If ws.Cells(currentRow + 1, 1).Value <> "FX" Then
        ws.Cells(currentRow + 1, 1).Value = "FX"
        ws.Cells(currentRow + 4, 1).Value = "Code"
        ws.Cells(currentRow + 4, 2).Value = "기준환율"
        ws.Cells(currentRow + 4, 3).Value = "Mar환율"
    End If

    
    currentColumn = 4
    For Each item In dataGet1
        ws.Cells(currentRow + 4, currentColumn).Value = item("dataId")
        currentColumn = currentColumn + 1
    Next item
    
    Dim uniqueFXIds As Object
    Set uniqueFXIds = CreateObject("Scripting.Dictionary")

    ' Iterate over each item in dataGet2 collection
    Dim item2 As Object
    For Each item2 In dataGet2
        ' Check both th01DataId and th02DataId for the substring "FX"
        If InStr(item2("th01DataId"), "FX") > 0 Then
            ' Add to the Dictionary if not already present
            If Not uniqueFXIds.Exists(item2("th01DataId")) Then
                uniqueFXIds.Add item2("th01DataId"), item2("th01DataId")
            End If
        End If

        If InStr(item2("th02DataId"), "FX") > 0 Then
            ' Add to the Dictionary if not already present
            If Not uniqueFXIds.Exists(item2("th02DataId")) Then
                uniqueFXIds.Add item2("th02DataId"), item2("th02DataId")
            End If
        End If
    Next item2

    Dim item3 As Variant
    currentRow2 = currentRow + 5
    For Each item3 In uniqueFXIds
        ws.Cells(currentRow2, 1).Value = item3
        currentRow2 = currentRow2 + 1
    Next item3
End Sub

'이거 작동하는 코드 correlation 전까지
'기초자산 간 correlation값 까지 작동


