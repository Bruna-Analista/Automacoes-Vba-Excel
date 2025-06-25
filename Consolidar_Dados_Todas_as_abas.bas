Attribute VB_Name = "Módulo1"
Sub Consolidar_ExpandindoMescladas_ComHotel()

    Dim PastaOrigem As String
    Dim Arquivo As String
    Dim wbOrigem As Workbook
    Dim wsOrigem As Worksheet
    Dim wsDestino As Worksheet
    Dim UltimaLinhaDestino As Long
    Dim linha As Long
    Dim col As Long
    Dim valorCelula As Variant
    Dim cel As Range
    Dim mergeAltura As Long
    Dim i As Long
    Dim nomeHotel As String

    ' Caminho da sua pasta
    PastaOrigem = "C:\Users\bruna.zordenoni\Desktop\Bruna\Projects\POWER BI\Desafio Junior\vba copiar e colar excel\Mapas de hospedagem 21.04 à 20.05\"

    Set wsDestino = ThisWorkbook.Sheets("Consolidado")
    

    Arquivo = Dir(PastaOrigem & "*.xlsx")

    Do While Arquivo <> ""
        Set wbOrigem = Workbooks.Open(PastaOrigem & Arquivo)

        For Each wsOrigem In wbOrigem.Sheets

            ' Captura o nome do hotel da célula mesclada C4:D4
            nomeHotel = wsOrigem.Range("C4").Value

            linha = 8
            Do While linha <= 52

                ' Determina quantas linhas vamos copiar com base em mesclagens
                mergeAltura = 1
                For col = 2 To 36
                    Set cel = wsOrigem.Cells(linha, col)
                    If cel.MergeCells Then
                        If cel.MergeArea.Rows.Count > mergeAltura Then
                            mergeAltura = cel.MergeArea.Rows.Count
                        End If
                    End If
                Next col

                ' Copia cada linha individual
                For i = 0 To mergeAltura - 1
                    UltimaLinhaDestino = wsDestino.Cells(wsDestino.Rows.Count, 1).End(xlUp).Row + 1
                    For col = 2 To 36
                        Set cel = wsOrigem.Cells(linha + i, col)
                        If cel.MergeCells Then
                            valorCelula = cel.MergeArea.Cells(1, 1).Value
                        Else
                            valorCelula = cel.Value
                        End If
                        wsDestino.Cells(UltimaLinhaDestino, col - 1).Value = valorCelula
                    Next col

                    ' Coloca o nome do hotel na próxima coluna (AK = 37)
                    wsDestino.Cells(UltimaLinhaDestino, 36).Value = nomeHotel
                Next i

                linha = linha + mergeAltura

            Loop

        Next wsOrigem

        wbOrigem.Close SaveChanges:=False
        Arquivo = Dir
    Loop

    MsgBox "Consolidação finalizada com sucesso!"

End Sub
