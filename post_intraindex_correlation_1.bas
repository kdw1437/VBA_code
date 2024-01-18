Attribute VB_Name = "Module5"
Sub PostCorrelation1()

    
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
    targetDate = Sheets("Market Data").Range("A2").value
    
    baseDt = Format(targetDate, "yyyymmdd")
    dataSetId = Sheets("Market Data").Range("O2").value
    StartingPoint = Sheets("Market Data").Range("P2").value
    
    Dim Table1Point As Range
    Set Table1Point = Sheets("Market Data").Range(StartingPoint).Offset(3, 0)
    
    ' Find the cell that contains "FX" after "Equity" table
    Dim fxRow As Range
    Set fxRow = ws.Range(Table1Point.Offset(1, 0), ws.Cells(ws.Rows.Count, Table1Point.Column)).Find(What:="FX", LookIn:=xlValues, LookAt:=xlWhole)
    Debug.Print fxRow.value
    Debug.Print Table1Point.value
    Debug.Print fxRow.Row
    Debug.Print Table1Point.Row
    
    Dim indexArray() As Variant
    
    Dim arraySize As Integer
    arraySize = fxRow.Row - Table1Point.Row - 2
    
    'Resize the array to the desired size
    ReDim indexArray(1 To arraySize)
    
    'Loop through the array to populate it
    For i = 1 To arraySize
        indexArray(i) = ws.Cells(Table1Point.Row + i, Table1Point.Column).value
    Next i
    
    Dim j As Long
    Dim k As Long
    Dim combined_name As String
    Dim valueofcorrelation As Double
    'Dim correlationRow As Integer
    'Dim correlationColumn As Integer
    Dim DataString As String
    ' Initialize the DataString
    DataString = ""
    
    'j가 가로로 진행, k가 세로로 진행. 고로 특정 j에서 k값 하나씩 correlation 붙여 넣도록 하기.
    'j값이 예를 들어서 2이면, k값이 1, 2인경우는 생략 가능. j와 k값이 같은 경우에도 생략 가능
    'j값이 3이면, k값이 1, 2, 3인 경우는 생략 가능.
    For j = LBound(indexArray) To UBound(indexArray)
        For k = LBound(indexArray) To UBound(indexArray)
            If Not (j = k Or j > k) Then
                combined_name = indexArray(j) & ":" & indexArray(k)
                valueofcorrelation = ws.Cells(Table1Point.Row + k, Table1Point.Column + j + 1).value
                                ' Construct the string
                If Len(DataString) > 0 Then
                    DataString = DataString & "&"
                End If
                DataString = DataString & "BASE_DT=" & baseDt & _
                             "&DATA_SET_ID=" & dataSetId & _
                             "&DATA_ID=" & combined_name & _
                             "&CRLT_CFCN_MATX_ID=CORR" & _
                             "&TH01_DATA_ID=" & indexArray(j) & _
                             "&TH02_DATA_ID=" & indexArray(k) & _
                             "&CRLT_CFCN=" & valueofcorrelation & _
                             "&OCR_DT=" & baseDt & _
                             "&PGM_ID=MANUALLY_INPUT" & _
                             "&WRKR_ID=HS" & _
                             "&WORK_TRIP=0.0.0.0"
                
            End If
        Next k
    Next j
    Debug.Print DataString
    
            ' Encode the DataString for URL (x-www-form-urlencoded)
    DataString = URLEncode(DataString)

    ' Create a new XML HTTP request
    Set xmlhttp = CreateObject("WinHttp.WinHttpRequest.5.1")

    ' The URL to send the request to
    Dim url As String
    url = "http://localhost:8080/val/postindexcorrelation"

    ' Open the HTTP request as a POST method
    xmlhttp.Open "POST", url, False

    ' Set the request content-type header to application/x-www-form-urlencoded
    xmlhttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    
    ' Send the request with the DataString
    xmlhttp.Send "a=" & DataString

    ' Check the status of the request
    If xmlhttp.Status = 200 Then
        ' If the request was successful, output the response
        MsgBox xmlhttp.responseText
    Else
        ' If the request failed, output the status
        MsgBox "Error: " & xmlhttp.Status & " - " & xmlhttp.statusText
    End If

    ' Clean up
    Set xmlhttp = Nothing
End Sub

Function URLEncode(StringVal As String, Optional SpaceAsPlus As Boolean = False) As String
    Dim StringLen As Long: StringLen = Len(StringVal)

    If StringLen > 0 Then
        ReDim result(StringLen) As String
        Dim i As Long, CharCode As Integer
        Dim Char As String, Space As String

        If SpaceAsPlus Then Space = "+" Else Space = "%20"

        For i = 1 To StringLen
            Char = Mid$(StringVal, i, 1)
            CharCode = Asc(Char)

            Select Case CharCode
                Case 97 To 122, 65 To 90, 48 To 57, 45, 46, 95, 126
                    result(i) = Char
                Case 32
                    result(i) = Space
                Case 0 To 15
                    result(i) = "%0" & Hex(CharCode)
                Case Else
                    result(i) = "%" & Hex(CharCode)
            End Select
        Next i

        URLEncode = Join(result, "")
    End If
End Function

