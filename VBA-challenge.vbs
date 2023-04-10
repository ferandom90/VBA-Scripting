Sub Hoja2018()

    Dim fin As Long
    Dim index As Long
    Dim uniqueThickers As Variant
    
    With Sheets("2018")
        ' Obtener el último row que tiene datos y guardar ese número en "fin".
        fin = WorksheetFunction.CountA(.Range("A:A"))
        
        ' Obtener todos los "Thickers" sin repetirse.
        uniqueThickers = WorksheetFunction.Unique(Range("A2:A" & fin))
        
        ' Recorremos un bucle del listado de Thickers sin repetirse y asignar su Valor en una celda.
        For index = LBound(uniqueThickers) To UBound(uniqueThickers)
            Range("I" & index + 1) = uniqueThickers(index, 1)
        Next
        
        Dim uniqueThickerCounter As Long
        Dim thickerCounter As Long
        uniqueThickerCounter = 1
        thickerCounter = 1
            
        ' Posición inicial de la celda Open del primer Thicker.
        Dim punteroInicialOpen As Long
        punteroInicialOpen = 2
        
        For Each uniqueThicker In uniqueThickers
            
            For Each thicker In Worksheets("2018").Range("A2:A" & fin)
                
                If uniqueThicker = thicker Then
                    
                    ' Obtener la posicion final de la celda Close del Thicker actual.
                    Dim punteroFinalClose As Long
                    punteroFinalClose = Worksheets("2018").Columns("A").Find(thicker, searchorder:=xlByRows, searchDirection:=xlPrevious).Row
                
                End If
    
                thickerCounter = thickerCounter + 1
                
            Next ' Fin de bucle de Thickers.
                
            ' Obtener el Valor de la primera celda de Open del Thicker actual y guardarlo en "firstValueOpen".
            Dim firstValueOpen As Double
            firstValueOpen = Worksheets("2018").Range("C" & punteroInicialOpen).Value
            
            ' Obtener el Valor de la última celda de Close del Thicker actual y guardarlo en "lastValueClose".
            Dim lastValueClose As Double
            lastValueClose = Worksheets("2018").Range("F" & punteroFinalClose).Value
            
            ' Calcular el Valor de Yearly Change del Thicker actual y asignarlo a la celda correspondiente.
            Dim yearlyChange As Double
            yearlyChange = lastValueClose - firstValueOpen
            Worksheets("2018").Range("J" & (uniqueThickerCounter + 1)) = yearlyChange
            
            ' Asignarle un color de fondo a la celda Yearly Change.
            If yearlyChange < 0 Then
                Worksheets("2018").Range("J" & (uniqueThickerCounter + 1)).Interior.ColorIndex = 3 ' ColorIdex=3 es el rojo
            Else
                Worksheets("2018").Range("J" & (uniqueThickerCounter + 1)).Interior.ColorIndex = 4 ' ColorIdex=4es el verde
            End If
            
            ' Calcular el Valor de Percentage Change del Thicker actual y asignarlo a la celda correspondiente.
            Dim pctChange As Double
            'pctChange = yearlyChange / firstValueOpen * 100
            pctChange = (yearlyChange * 100) / firstValueOpen
            pctChange = Round(pctChange, 2)
            Worksheets("2018").Range("K" & (uniqueThickerCounter + 1)) = pctChange & "%"
            
            ' Recorrer con un bucle las celdas de la columna Vol para calcular el "Total Stock Volume".
            Dim totalStockVolume As LongLong
            totalStockVolume = 0
            For Each volume In Worksheets("2018").Range("G" & punteroInicialOpen & ":G" & punteroFinalClose)
                totalStockVolume = totalStockVolume + volume.Value
            Next ' Fin de bucle volume
            
            ' Asignar el valor de "totalStockVolume" a la celda correspondiente.
            Worksheets("2018").Range("L" & (uniqueThickerCounter + 1)) = totalStockVolume
            
            punteroInicialOpen = punteroFinalClose + 1
            uniqueThickerCounter = uniqueThickerCounter + 1
            
        Next ' Fin de bucle de uniqueThickers.
        
    End With

End Sub


' ------------------------------------------------------------------------------------------
' ------------------------------------------------------------------------------------------

Sub Hoja2019()

    Dim fin As Long
    Dim index As Long
    Dim uniqueThickers As Variant
    
    With Sheets("2019")
        ' Obtener el último row que tiene datos y guardar ese número en "fin".
        fin = WorksheetFunction.CountA(.Range("A:A"))
        
        ' Obtener todos los "Thickers" sin repetirse.
        uniqueThickers = WorksheetFunction.Unique(Range("A2:A" & fin))
        
        ' Recorremos un bucle del listado de Thickers sin repetirse y asignar su Valor en una celda.
        For index = LBound(uniqueThickers) To UBound(uniqueThickers)
            Range("I" & index + 1) = uniqueThickers(index, 1)
        Next
        
        Dim uniqueThickerCounter As Long
        Dim thickerCounter As Long
        uniqueThickerCounter = 1
        thickerCounter = 1
            
        ' Posición inicial de la celda Open del primer Thicker.
        Dim punteroInicialOpen As Long
        punteroInicialOpen = 2
        
        For Each uniqueThicker In uniqueThickers
            
            For Each thicker In Worksheets("2019").Range("A2:A" & fin)
                
                If uniqueThicker = thicker Then
                    
                    ' Obtener la posicion final de la celda Close del Thicker actual.
                    Dim punteroFinalClose As Long
                    punteroFinalClose = Worksheets("2019").Columns("A").Find(thicker, searchorder:=xlByRows, searchDirection:=xlPrevious).Row
                
                End If
    
                thickerCounter = thickerCounter + 1
                
            Next ' Fin de bucle de Thickers.
                
            ' Obtener el Valor de la primera celda de Open del Thicker actual y guardarlo en "firstValueOpen".
            Dim firstValueOpen As Double
            firstValueOpen = Worksheets("2019").Range("C" & punteroInicialOpen).Value
            
            ' Obtener el Valor de la última celda de Close del Thicker actual y guardarlo en "lastValueClose".
            Dim lastValueClose As Double
            lastValueClose = Worksheets("2019").Range("F" & punteroFinalClose).Value
            
            ' Calcular el Valor de Yearly Change del Thicker actual y asignarlo a la celda correspondiente.
            Dim yearlyChange As Double
            yearlyChange = lastValueClose - firstValueOpen
            Worksheets("2019").Range("J" & (uniqueThickerCounter + 1)) = yearlyChange
            
            ' Asignarle un color de fondo a la celda Yearly Change.
            If yearlyChange < 0 Then
                Worksheets("2019").Range("J" & (uniqueThickerCounter + 1)).Interior.ColorIndex = 3 ' ColorIdex=3 es el rojo
            Else
                Worksheets("2019").Range("J" & (uniqueThickerCounter + 1)).Interior.ColorIndex = 4 ' ColorIdex=4es el verde
            End If
            
            ' Calcular el Valor de Percentage Change del Thicker actual y asignarlo a la celda correspondiente.
            Dim pctChange As Double
            'pctChange = yearlyChange / firstValueOpen * 100
            pctChange = (yearlyChange * 100) / firstValueOpen
            pctChange = Round(pctChange, 2)
            Worksheets("2019").Range("K" & (uniqueThickerCounter + 1)) = pctChange & "%"
            
            ' Recorrer con un bucle las celdas de la columna Vol para calcular el "Total Stock Volume".
            Dim totalStockVolume As LongLong
            totalStockVolume = 0
            For Each volume In Worksheets("2019").Range("G" & punteroInicialOpen & ":G" & punteroFinalClose)
                totalStockVolume = totalStockVolume + volume.Value
            Next ' Fin de bucle volume
            
            ' Asignar el valor de "totalStockVolume" a la celda correspondiente.
            Worksheets("2019").Range("L" & (uniqueThickerCounter + 1)) = totalStockVolume
            
            punteroInicialOpen = punteroFinalClose + 1
            uniqueThickerCounter = uniqueThickerCounter + 1
            
        Next ' Fin de bucle de uniqueThickers.
        
    End With

End Sub


' ------------------------------------------------------------------------------------------
' ------------------------------------------------------------------------------------------

Sub Hoja2020()

    Dim fin As Long
    Dim index As Long
    Dim uniqueThickers As Variant
    
    With Sheets("2020")
        ' Obtener el último row que tiene datos y guardar ese número en "fin".
        fin = WorksheetFunction.CountA(.Range("A:A"))
        
        ' Obtener todos los "Thickers" sin repetirse.
        uniqueThickers = WorksheetFunction.Unique(Range("A2:A" & fin))
        
        ' Recorremos un bucle del listado de Thickers sin repetirse y asignar su Valor en una celda.
        For index = LBound(uniqueThickers) To UBound(uniqueThickers)
            Range("I" & index + 1) = uniqueThickers(index, 1)
        Next
        
        Dim uniqueThickerCounter As Long
        Dim thickerCounter As Long
        uniqueThickerCounter = 1
        thickerCounter = 1
            
        ' Posición inicial de la celda Open del primer Thicker.
        Dim punteroInicialOpen As Long
        punteroInicialOpen = 2
        
        For Each uniqueThicker In uniqueThickers
            
            For Each thicker In Worksheets("2020").Range("A2:A" & fin)
                
                If uniqueThicker = thicker Then
                    
                    ' Obtener la posicion final de la celda Close del Thicker actual.
                    Dim punteroFinalClose As Long
                    punteroFinalClose = Worksheets("2020").Columns("A").Find(thicker, searchorder:=xlByRows, searchDirection:=xlPrevious).Row
                
                End If
    
                thickerCounter = thickerCounter + 1
                
            Next ' Fin de bucle de Thickers.
                
            ' Obtener el Valor de la primera celda de Open del Thicker actual y guardarlo en "firstValueOpen".
            Dim firstValueOpen As Double
            firstValueOpen = Worksheets("2020").Range("C" & punteroInicialOpen).Value
            
            ' Obtener el Valor de la última celda de Close del Thicker actual y guardarlo en "lastValueClose".
            Dim lastValueClose As Double
            lastValueClose = Worksheets("2020").Range("F" & punteroFinalClose).Value
            
            ' Calcular el Valor de Yearly Change del Thicker actual y asignarlo a la celda correspondiente.
            Dim yearlyChange As Double
            yearlyChange = lastValueClose - firstValueOpen
            Worksheets("2020").Range("J" & (uniqueThickerCounter + 1)) = yearlyChange
            
            ' Asignarle un color de fondo a la celda Yearly Change.
            If yearlyChange < 0 Then
                Worksheets("2020").Range("J" & (uniqueThickerCounter + 1)).Interior.ColorIndex = 3 ' ColorIdex=3 es el rojo
            Else
                Worksheets("2020").Range("J" & (uniqueThickerCounter + 1)).Interior.ColorIndex = 4 ' ColorIdex=4es el verde
            End If
            
            ' Calcular el Valor de Percentage Change del Thicker actual y asignarlo a la celda correspondiente.
            Dim pctChange As Double
            'pctChange = yearlyChange / firstValueOpen * 100
            pctChange = (yearlyChange * 100) / firstValueOpen
            pctChange = Round(pctChange, 2)
            Worksheets("2020").Range("K" & (uniqueThickerCounter + 1)) = pctChange & "%"
            
            ' Recorrer con un bucle las celdas de la columna Vol para calcular el "Total Stock Volume".
            Dim totalStockVolume As LongLong
            totalStockVolume = 0
            For Each volume In Worksheets("2020").Range("G" & punteroInicialOpen & ":G" & punteroFinalClose)
                totalStockVolume = totalStockVolume + volume.Value
            Next ' Fin de bucle volume
            
            ' Asignar el valor de "totalStockVolume" a la celda correspondiente.
            Worksheets("2020").Range("L" & (uniqueThickerCounter + 1)) = totalStockVolume
            
            punteroInicialOpen = punteroFinalClose + 1
            uniqueThickerCounter = uniqueThickerCounter + 1
            
        Next ' Fin de bucle de uniqueThickers.
        
    End With

End Sub
