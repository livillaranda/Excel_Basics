Sub Ticker()

'Summary Table Creation
Range("I1").value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"

'Format PercentChange Column
Columns("K:K").Select
Selection.NumberFormat = "0.00%"

'Identifying Values
Dim Ticker as String
Dim i as Long
Dim OpenPrice as Long
Dim TableRow as Integer
Dim TotalVolume as Long
Dim YearlyChange as Long
Dim PercentChange as String

    'Find Last Row
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    TableRow = 2
    TotalVolume = 0
    OpenPrice = 2


        'For Loop
        For i = 2 to LastRow
'EX: If Cells(7, 1) <> Cells(8, 1) Then
        If Cells(i, 1) <> Cells(i + 1, 1) Then


        'Table Generation
        Ticker = Cells(i, 1).Value
'EX: Ticker = Cells(7, 1).Value
        OpenPrice = Cells(i, 3).Value
'EX: OpenPrice = Cells (7,3)
        TotalVolume = TotalVolume + Cells(i, 7).Value
'EX: TotalVolume = 0 + Cells(7, 7).Value       
        YearlyChange = Cells(i, 6).Value - OpenPrice
'EX: YearlyChange = Cells(7, 6).Value - Cells(7, 3)        
        PercentChange = YearlyChange / OpenPrice
'Ex: PercentChange = Yearly Change / OpenPrice


    Range("I" & TableRow) = Ticker
    Range("L" & TableRow) = TotalVolume
    Range("J" & TableRow) = YearlyChange
    Range("K" & TableRow) = PercentChange
    
            Else 
        
            If YearlyChange > 0 Then
            Range("K" & TableRow).Interior.ColorIndex = 4

            Else
            Range("K" & TableRow).Interior.ColorIndex = 3        

            End If
         

    End If
    
    'Reset Volume >> Next Ticker
       TotalVolume = 0
     
       TableRow = Range("K" & TableRow) + 1
    
    Next i

    End Sub