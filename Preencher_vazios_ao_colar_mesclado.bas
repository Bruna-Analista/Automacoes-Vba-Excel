Attribute VB_Name = "Módulo2"
Sub PreencherVaziosMultiplasColunas_Seguro()
    Dim rng As Range
    Dim celulasVazias As Range
    Dim cel As Range

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    On Error Resume Next
    Set rng = Selection
    Set celulasVazias = rng.SpecialCells(xlCellTypeBlanks)
    On Error GoTo Limpar

    If Not celulasVazias Is Nothing Then
        For Each cel In celulasVazias
            If cel.Row > 1 Then
                ' Verifica se a célula acima não está vazia
                If Not IsEmpty(cel.Offset(-1, 0)) Then
                    cel.Value = cel.Offset(-1, 0).Value
                End If
            End If
        Next cel
    Else
        MsgBox "Não há células vazias no intervalo selecionado.", vbInformation
    End If

Limpar:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub


