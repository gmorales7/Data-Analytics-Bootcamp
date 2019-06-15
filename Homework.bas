Attribute VB_Name = "Module1"

    Sub doinsheets()

    
    Dim ws As Worksheet
        Application.ScreenUpdating = False
        
        For Each ws In Worksheets
        
            ws.Select
            Call stocks
        Next
            Application.ScreenUpdating = True
    End Sub
            
    Sub stocks()
    
    Dim lastRow As Long
    Dim Actual_ticker As String
    Dim Volume As Double
    Volume = 0
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
    For Each ws In Worksheets
    
    Range(1, 9).Value = "<Ticker>"
    Range(1, 10).Value = "<Total Stock Volume>"

    
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastRow
        If Cells(i + 1, 1).Value <> Cells(i, 1) Then

        Actual_ticker = Cells(i, 1).Value
        Volume = Volume + Cells(i, 7).Value
        Range("I" & Summary_Table_Row).Value = Actual_ticker
        Range("J" & Summary_Table_Row).Value = Volume

        Summary_Table_Row = Summary_Table_Row + 1

        Volume = 0

        Else
        Volume = Volume + Cells(i, 7).Value

        End If
    Next i

Next ws

End Sub
