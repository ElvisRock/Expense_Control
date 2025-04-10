Sub LimparDadosSemanal()

    Dim wsSemanal As Worksheet
    Dim rngUltimoDado As Range
    Dim ultimaLinha As Long
    
    Set wsSemanal = Sheets("Semanal")
    
    ' Procura a última célula usada no intervalo B:R
    Set rngUltimoDado = wsSemanal.Range("B:R").Find(What:="*", _
                        LookIn:=xlValues, _
                        SearchOrder:=xlByRows, _
                        SearchDirection:=xlPrevious)
    
    ' Se encontrar algo...
    If Not rngUltimoDado Is Nothing Then
        ultimaLinha = rngUltimoDado.Row
        ' Se a última linha é >= 4 (onde começam seus dados), limpa
        If ultimaLinha >= 4 Then
            wsSemanal.Range("B4:R" & ultimaLinha).ClearContents
        End If
    End If

End Sub
