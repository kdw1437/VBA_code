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
    'row�� ���η� �̵��ϸ鼭 dataId key�� value���� �־��ش�.
    
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
    'When I dont' know beforehand how many columns contain data.
    For columnIndex = 3 To ws.Cells(columnnameRow, Columns.Count).End(xlToLeft).Column
        Dim headerValue As String
        headerValue = ws.Cells(columnnameRow, columnIndex).Value
        
        For rowIndex = startRow To ws.Cells(Rows.Count, 1).End(xlUp).Row
            If ws.Cells(rowIndex, 1).Value = headerValue Then
                ws.Cells(rowIndex, columnIndex).Value = 1
            End If
        Next rowIndex
    Next columnIndex
    '���⼭ ���� �߰� �ڵ� �ۼ� (Corrleation matrix �־��ֱ�)
    Dim dataGet2 As Collection
    Set dataGet2 = jsonResponse("data_get_2")
    
    For columnIndex = 3 To ws.Cells(columnnameRow, Columns.Count).End(xlToLeft).Column
        
        headerValue = ws.Cells(columnnameRow, columnIndex).Value
        
        For rowIndex = startRow To ws.Cells(Rows.Count, 1).End(xlUp).Row
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


End Sub

'�̰� �۵��ϴ� �ڵ� correlation ������
'�����ڻ� �� correlation�� ���� �۵�

