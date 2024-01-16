Attribute VB_Name = "Module3"
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
    Dim codeCol As Integer
    equityRow = ws.Columns(1).Find(What:="Equity", LookIn:=xlValues, LookAt:=xlPart).Row
    codeCol = equityRow + 1 ' Assuming 'code' is right below 'Equity'
    
    ' Write headers 'code' and 'ClosedPrice'
    ws.Cells(codeCol, 1).Value = "code"
    ws.Cells(codeCol, 2).Value = "ClosedPrice"

    ' Starting row for writing data is 4 rows below 'Equity' + 1 for the header
    Dim startRow As Integer
    startRow = equityRow + 5

    ' Loop over each item in the data_get_1 array
    Dim item As Dictionary
    Dim currentRow As Integer
    currentRow = startRow

    For Each item In dataGet1
        ' Write dataId to column A and closePric to column B
        ws.Cells(currentRow, 3).Value = item("dataId") ' 2 columns right from 'code'
        ws.Cells(currentRow, 2).Value = item("closePric")

        ' Increment row counter
        currentRow = currentRow + 1
    Next item
End Sub

