Attribute VB_Name = "Module1"
Option Explicit
Sub SummarizeYear(ws As Worksheet)
    ThisWorkbook.Application.ReferenceStyle = xlR1C1
    ws.Range(ws.Cells(1, 9), ws.Cells(1, 17)).EntireColumn.Delete
    
    ws.Sort.SortFields.Clear
    ws.Sort.SortFields.Add2 Key:=Range("A:A"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ws.Sort.SortFields.Add2 Key:=Range("B:B"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ws.Sort
        .SetRange Range("A:G")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Dim varInput() As Variant
    varInput = ws.Range("A:G").Value2
    
    Dim zColHdr As Object
    Set zColHdr = CreateObject("scripting.dictionary")
    
    Dim i As Long, iVol As LongLong, iCurR As Long, iMaxV As LongLong
    Dim dblJunk As Double, dblOpen As Double, dblClose As Double, dblMaxI As Double, dblMaxD As Double
    Dim strJunk As String, strMaxI As String, strMaxD As String, strMaxV As String
    
    For i = 1 To UBound(varInput, 2)
        strJunk = varInput(1, i)
        zColHdr(Mid(strJunk, 2, Len(strJunk) - 2)) = i
    Next
    
    Call zzz_SetScreenUpdating(True)
    
    iCurR = 1
    dblOpen = 1
    dblClose = 1

    For i = 2 To UBound(varInput, 1)
        'Still evaluating the same stock
        If varInput(i, zColHdr("ticker")) = strJunk Then
            iVol = iVol + varInput(i, zColHdr("vol"))
            dblClose = varInput(i, zColHdr("close"))
        
        'End of the previous stock
        Else
            dblJunk = ((dblClose - dblOpen) / dblOpen)
            If dblMaxI < dblJunk Then
                dblMaxI = dblJunk
                strMaxI = strJunk
            ElseIf dblMaxD > dblJunk Then
                dblMaxD = dblJunk
                strMaxD = strJunk
            End If
            If iMaxV < iVol Then
                iMaxV = iVol
                strMaxV = strJunk
            End If
            ws.Range(ws.Cells(iCurR, 9), ws.Cells(iCurR, 12)) = Array(strJunk, dblClose - dblOpen, dblJunk, iVol)
            If varInput(i, zColHdr("ticker")) = "" Then
                Exit For
            End If
            iCurR = iCurR + 1
            strJunk = varInput(i, zColHdr("ticker"))
            dblOpen = varInput(i, zColHdr("open"))
            dblClose = varInput(i, zColHdr("close"))
            iVol = varInput(i, zColHdr("vol"))
        End If
    Next
    With ws.Range(ws.Cells(2, 10), ws.Cells(iCurR, 10))
        .FormatConditions.Add xlExpression, xlEqual, "=RC<0"
        .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(255, 0, 0)
        .FormatConditions.Add xlExpression, xlEqual, "=RC>=0"
        .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(0, 255, 0)
    End With
    ws.Range(ws.Cells(1, 9), ws.Cells(1, 12)) = Array("Ticker", "Yearly Change", "Percent Change", "Total Stock Volume")
    ws.Columns(11).NumberFormat = "0.00%"
    ws.Columns(12).EntireColumn.AutoFit
    
    ws.Range(ws.Cells(1, 16), ws.Cells(1, 17)) = Array("Ticker", "Value")
    ws.Range(ws.Cells(2, 15), ws.Cells(2, 17)) = Array("Greatest % Increase", strMaxI, dblMaxI)
    ws.Range(ws.Cells(3, 15), ws.Cells(3, 17)) = Array("Greatest % Decrease", strMaxD, dblMaxD)
    ws.Range(ws.Cells(4, 15), ws.Cells(4, 17)) = Array("Greatest Total Volume", strMaxV, iMaxV)
    ws.Range(ws.Cells(2, 17), ws.Cells(3, 17)).NumberFormat = "0.00%"
    ws.Range(ws.Cells(1, 15), ws.Cells(1, 17)).EntireColumn.AutoFit

    Call zzz_SetScreenUpdating
End Sub
Sub SummarizeAll()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        Call SummarizeYear(ws)
    Next
    MsgBox ("Complete")
End Sub
Sub zzz_SetScreenUpdating(Optional bOff As Boolean)
    Dim xlApp As Application
    If xlApp Is Nothing Then Set xlApp = ThisWorkbook.Application
    With xlApp
        If bOff Then
            .ScreenUpdating = False
            .Calculation = xlManual
        Else
            .ScreenUpdating = True
            .Calculation = xlAutomatic
            .StatusBar = False
        End If
    End With
End Sub
