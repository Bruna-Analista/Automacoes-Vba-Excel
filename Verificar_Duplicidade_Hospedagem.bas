Attribute VB_Name = "Módulo1"
Sub VerificarHospedagemDuplicada()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("maio.2025")
    
    Dim ultimaLinha As Long
    ultimaLinha = ws.Cells(ws.Rows.Count, "H").End(xlUp).Row
    
    ' Definindo colunas fixas por nome
    Dim colNome As Integer: colNome = ws.Range("H1").Column
    Dim colTipoHosp As Integer: colTipoHosp = ws.Range("J1").Column
    Dim colInicio As Integer: colInicio = ws.Range("L1").Column
    Dim colFim As Integer: colFim = ws.Range("AP1").Column

    Dim col As Integer, linha As Long
    Dim dict As Object
    Dim nomeFuncionario As String
    Dim tipoHospedagem As String
    Dim chave As String

    ' Limpar formatação anterior
    ws.Range(ws.Cells(5, colInicio), ws.Cells(ultimaLinha, colFim)).Interior.ColorIndex = xlNone

    ' Verifica cada coluna de data (dia)
    For col = colInicio To colFim
        Set dict = CreateObject("Scripting.Dictionary")
        
        For linha = 5 To ultimaLinha
            nomeFuncionario = Trim(ws.Cells(linha, colNome).Value)
            tipoHospedagem = UCase(Trim(ws.Cells(linha, colTipoHosp).Value))
            
            ' Verifica se é hospedagem e a célula do dia não está vazia
            If nomeFuncionario <> "" And tipoHospedagem = "HOSP" And ws.Cells(linha, col).Value <> "" Then
                chave = nomeFuncionario
                
                If dict.exists(chave) Then
                    ' Marcar duplicações em vermelho
                    ws.Cells(dict(chave), col).Interior.Color = RGB(255, 150, 150)
                    ws.Cells(linha, col).Interior.Color = RGB(255, 150, 150)
                    
                    ' Opcional: adicionar comentário (descomente se quiser usar)
                    ' On Error Resume Next
                    ' ws.Cells(linha, col).AddComment "Duplicado: " & nomeFuncionario
                    ' On Error GoTo 0
                    
                Else
                    dict.Add chave, linha
                End If
            End If
        Next linha
    Next col

    MsgBox "Verificação concluída. Células duplicadas foram destacadas.", vbInformation

End Sub

