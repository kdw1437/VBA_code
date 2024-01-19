Attribute VB_Name = "Module6"
Sub PostCorrelation2()
    
    Dim xmlhttp As Object
    Dim i As Integer
    Dim baseDt As String
    Dim dataSetId As String
    Dim StartingPoint As String
    Dim dataId As String
    
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
    
    Dim lastRow As Long
    
    lastRow = ws.Cells(ws.Rows.Count, Table1Point.Column).End(xlUp).Row
    
    ' Find the cell that contains "FX" after "Equity" table
    Dim fxRow As Range
    Set fxRow = ws.Range(Table1Point.Offset(1, 0), ws.Cells(lastRow, Table1Point.Column)).Find(What:="FX", LookIn:=xlValues, LookAt:=xlWhole)
    Debug.Print fxRow.value
    Debug.Print Table1Point.value
    Debug.Print fxRow.Row
    Debug.Print Table1Point.Row
    Debug.Print lastRow
    
    Dim Table2Point As Range
    Set Table2Point = fxRow.Offset(3, 0)
    
    Dim YieldCurveRow As Range
    Set YieldCurveRow = ws.Range(Table1Point.Offset(1, 0), ws.Cells(lastRow, Table1Point.Column)).Find(What:="Yield Curve", LookIn:=xlValues, LookAt:=xlWhole)
    'Debug.Print Table2Point.value
    Debug.Print Table2Point.Row
    Debug.Print YieldCurveRow.value
    Debug.Print YieldCurveRow.Row
    
    Dim indexArray1() As Variant
    
    Dim ArraySize1 As Integer
    ArraySize1 = YieldCurveRow.Row - Table2Point.Row - 2
    
    Dim currentCol As Long
    Dim lastDataCol As Long
    
    currentCol = Table2Point.Column
    lastDataCol = currentCol
    
    'Loop thorugh the cells in the row starting from Table2Point
    Do While ws.Cells(Table2Point.Row, currentCol).value <> ""
        lastDataCol = currentCol
        currentCol = currentCol + 1
    Loop
    Dim indexArray2() As Variant
    
    Dim ArraySize2 As Integer
    ArraySize2 = lastDataCol - Table2Point.Column + 1 - 3
    
    ReDim indexArray1(1 To ArraySize1)
    ReDim indexArray2(1 To ArraySize2)
    'indexArray1
    For i = 1 To ArraySize1
        indexArray1(i) = ws.Cells(Table2Point.Row + i, Table2Point.Column).value
    Next i

    For i = 1 To ArraySize2
        indexArray2(i) = ws.Cells(Table2Point.Row, Table2Point.Column + 2 + i).value
    Next i
    
    Dim combined_name As String
    Dim valueofcorrelation As Double
    Dim DataString As String
    ' Initialize the DataString
    DataString = ""
    
    Dim j As Integer
    
    For i = 1 To ArraySize2
        For j = 1 To ArraySize1
            If ws.Cells(Table2Point.Row + j, Table2Point.Column + 2 + i).value <> "" Then
                combined_name = indexArray2(i) & ":" & indexArray1(j)
                valueofcorrelation = ws.Cells(Table2Point.Row + j, Table2Point.Column + 2 + i).value
                If Len(DataString) > 0 Then
                    DataString = DataString & "&"
                End If
                DataString = DataString & "BASE_DT=" & baseDt & _
                             "&DATA_SET_ID=" & dataSetId & _
                             "&DATA_ID=" & combined_name & _
                             "&CRLT_CFCN_MATX_ID=CORR" & _
                             "&TH01_DATA_ID=" & indexArray2(i) & _
                             "&TH02_DATA_ID=" & indexArray1(j) & _
                             "&CRLT_CFCN=" & valueofcorrelation & _
                             "&OCR_DT=" & baseDt & _
                             "&PGM_ID=MANUALLY_INPUT" & _
                             "&WRKR_ID=HS" & _
                             "&WORK_TRIP=0.0.0.0"
            End If
        Next j
    Next i
    Debug.Print DataString
    
    ' Encode the DataString for URL (x-www-form-urlencoded)
    DataString = URLEncode(DataString)

    ' Create a new XML HTTP request
    Set xmlhttp = CreateObject("WinHttp.WinHttpRequest.5.1")

    ' The URL to send the request to
    Dim url As String
    url = "http://localhost:8080/val/extend_8"

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


