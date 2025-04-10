Sub Macro1()
    Dim wsSemanal As Worksheet, wsAnual As Worksheet
    Dim ultimaLinhaSemanal As Long, ultimaLinhaAnual As Long

    Set wsSemanal = Sheets("Semanal")
    Set wsAnual = Sheets("Anual")

    ' 1) Acha a última linha preenchida na Semanal (coluna B)
    ultimaLinhaSemanal = wsSemanal.Cells(wsSemanal.Rows.Count, "B").End(xlUp).Row
    
    ' 2) Acha a última linha preenchida na Anual (coluna B)
    ultimaLinhaAnual = wsAnual.Cells(wsAnual.Rows.Count, "B").End(xlUp).Row
    
    ' 3) Se a aba Anual está vazia (ou com menos de 3 linhas preenchidas em B),
    '    copiamos o cabeçalho (linha 3) da Semanal para a linha 3 da Anual
    If ultimaLinhaAnual < 3 Then
        wsSemanal.Range("B3:R3").Copy Destination:=wsAnual.Range("B3")
        ultimaLinhaAnual = 3
    End If
    
    ' 4) Agora copiamos apenas as linhas de dados, que começam em B4:R4 até a última linha
    If ultimaLinhaSemanal >= 4 Then
        wsSemanal.Range("B4:R" & ultimaLinhaSemanal).Copy
        ' Cola na Anual, a partir da próxima linha livre (ultimaLinhaAnual + 1)
        wsAnual.Cells(ultimaLinhaAnual + 1, "B").PasteSpecial xlPasteAll
    End If

    ' Ajustar a largura das colunas
    wsAnual.Columns("B:R").AutoFit
    
    ' Remove a seleção de "Copy"
    Application.CutCopyMode = False
    
    ' (Opcional) Volta para a aba "Semanal"
    wsSemanal.Activate
    wsSemanal.Range("A2").Select
End Sub
