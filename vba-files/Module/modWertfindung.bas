Attribute VB_Name = "modWertfindung"
Option Explicit

Public Const nHeaderRow = 5

Sub FillSummarySheet()
'(0) Define variables
    Dim wSource As Worksheet, wTarget As Worksheet
    Set wSource = Worksheets("Wertfindung")
    Set wTarget = Worksheets("Zusammenfassung")
    Dim i As Long, j As Long
    Dim rg As Range, sType As String
    Dim dictKontenplan As Scripting.Dictionary
'(1) Clear data from 'Zusammenfassung'
    wTarget.Cells(nHeaderRow, 1).CurrentRegion.ClearContents
'(2) (Re-) Create Header
    wTarget.Cells(nHeaderRow, 1).Value = "Art"
    wTarget.Cells(nHeaderRow, 2).Value = "Konto"
    wTarget.Cells(nHeaderRow, 3).Value = "Hauptkategorie"
    wTarget.Cells(nHeaderRow, 4).Value = "Subkategorie"
    wTarget.Cells(nHeaderRow, 5).Value = "Detailkategorie"
    wTarget.Cells(nHeaderRow, 6).Value = "Bezeichnung Kontenplan"
    Set rg = wTarget.Range(Cells(nHeaderRow, 1), Cells(nHeaderRow, 6))
    rg.Font.Bold = True
'(3) Fill data from 'Wertfindung'
    Set rg = wSource.Range(GetDataRangeFromWorksheet(wSource, nHeaderRow))
    j = nHeaderRow + 1
    For i = 1 To rg.Rows.Count
        sType = rg.Cells(i, 1).Value
        If sType = "Aufwand" Or sType = "Ertrag" Then
            'Struktur der Tabellen in 'Zusammenfassung' und 'Wertfindung' berücksichtigen!
            wTarget.Cells(j, 1).Value = rg.Cells(i, 1).Value    'Art
            wTarget.Cells(j, 2).Value = rg.Cells(i, 6).Value    'Konto
            wTarget.Cells(j, 3).Value = rg.Cells(i, 3).Value    'Hauptkategorie
            wTarget.Cells(j, 4).Value = rg.Cells(i, 4).Value    'Subkategorie
            wTarget.Cells(j, 5).Value = rg.Cells(i, 5).Value    'Detailkategorie
            j = j + 1
        End If
    Next i
'(4) Kontenbezeichnungen hinzufügen
    Set wSource = Worksheets("Kontenplan")
    Set rg = wSource.Range(GetDataRangeFromWorksheet(wSource, nHeaderRow))
    Set dictKontenplan = New Scripting.Dictionary
    For i = 1 To rg.Rows.Count
        'Struktur der Tabell 'Kontenplan' berücksichtigen!
        'Spalte 2 = Kontonummer
        'Spalte 3 = Kontobezeichnung
        dictKontenplan.Add rg.Cells(i, 2).Value, rg.Cells(i, 3).Value
    Next i
    Set rg = wTarget.Range(GetDataRangeFromWorksheet(wTarget, nHeaderRow))
    For i = 1 To rg.Rows.Count
        rg.Cells(i, 6).Value = dictKontenplan(rg.Cells(i, 2).Value)
    Next i
'(5) Autofit data table
    Set rg = wTarget.Cells(nHeaderRow, 1).CurrentRegion
    rg.Columns.AutoFit
'() Release variables
    Set wSource = Nothing
    Set wTarget = Nothing
    Set rg = Nothing
End Sub

