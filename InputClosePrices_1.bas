Attribute VB_Name = "InputClosePrices"
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
    baseURL = "http://localhost:8080/val/v1/ClosePrices/official?basedt="
    url = baseURL & dateParameter
    
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
    Dim selClosePrice As Collection
    Set selClosePrice = jsonResponse("selClosePrice")
    
    ' Find the row with 'Equity' in column A
    Dim equityRow As Integer
    equityRow = ws.Columns(1).Find(What:="Equity", LookIn:=xlValues, LookAt:=xlPart).Row

    Dim startRow As Integer
    startRow = equityRow + 4
    
    Dim codeCell As Range
    Dim codeValue As String
    Dim i As Integer
    
    ' Loop through each code in the worksheet starting from startRow
    For Each codeCell In ws.Range("A" & startRow & ":A" & ws.Rows.Count).Cells 'ws.Rows.Count: 마지막 row까지
        ' Stop if the cell is empty
        If IsEmpty(codeCell.Value) Then Exit For '비어있는 cell을 만나면 loop가 끝난다.
        
        ' Get the code value from the cell
        codeValue = codeCell.Value
        
        ' Loop through each item in the selClosePrice collection
        For i = 1 To selClosePrice.Count
            Dim data As Variant
            data = selClosePrice(i)("data") 'JSON object의 key인 data의 value값을 data에 저장
            
            ' Split the data string by '|'
            Dim dataParts As Variant
            dataParts = Split(data, "|")
            
            ' Check if the first part of the data (DATA_ID) matches the codeValue
            If dataParts(0) = codeValue Then
                ' If it matches, update the closed price in the next column
                codeCell.Offset(0, 1).Value = dataParts(3)
                Exit For 'Inner loop 밖으로 나오기 위해서 필요하다. codeValue에 해당하는 값을 찾았다.
            End If
        Next i
    Next codeCell
    
    MsgBox "Update complete!"
End Sub
    
    
