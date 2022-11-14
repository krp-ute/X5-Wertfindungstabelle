Attribute VB_Name = "modfunctions"
Option Explicit


Function GetAccountType(nRow As Long, sAccountCol As String) As String
    'sAccountCol wird übergeben als Spalte, welche die Kontonummern enthält
    Dim sCell As String, sNumber As String, sType As String
    sCell = Range(sAccountCol & nRow).Value
    sNumber = Left(sCell, 1)
    Select Case sNumber
        Case 1: sType = "Bilanz"
        Case 2: sType = "Bilanz"
        Case 3: sType = "Ertrag"
        Case 4: sType = "Aufwand"
        Case 5: sType = "Bilanz"
        Case 6: sType = "Bilanz"
        Case 7: sType = "Bilanz"
        Case 8: sType = "Bilanz"
        Case 9: sType = "Bilanz"
    End Select
    GetAccountType = sType
End Function


Function GetDataRangeFromWorksheet(ws As Worksheet, nHeaderRow) As String
    Dim sRg As String, arr As Variant
    sRg = ws.Range("A" & nHeaderRow).CurrentRegion.Address
    arr = Split(sRg, "$")
    sRg = "A" & nHeaderRow + 1 & ":" & arr(UBound(arr) - 1) & arr(UBound(arr))
    GetDataRangeFromWorksheet = sRg
End Function
