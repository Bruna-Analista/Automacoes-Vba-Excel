# 🧰 Automação VBA para Excel 

Este repositório reúne **macros VBA** desenvolvidas para automatizar tarefas comuns no Excel. São soluções aplicadas na rotina de controle de custos e dados operacionais em empresas de construção civil.

---

## 📌 Objetivo

Facilitar tarefas repetitivas, melhorar a confiabilidade dos dados e acelerar a consolidação de informações, como controle de hospedagem e verificação de inconsistências.

---

## 📂 Scripts incluídos

### 🔄 `consolidar-todas-as-abas.bas`
Consolida dados de todas as abas de arquivos Excel, copia apenas os valores (inclusive células mescladas) e insere o nome do hotel na primeira coluna.

## 💻 Script PowerShell

```powershell
Sub Consolidar_Mescladas()

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
    PastaOrigem = "S:\S_TECNICA_1\CUSTOS\Backup\Bruna\2025\hotel102025\"

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
                For col = 2 To 37
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
                    For col = 2 To 37
                        Set cel = wsOrigem.Cells(linha + i, col)
                        If cel.MergeCells Then
                            valorCelula = cel.MergeArea.Cells(1, 1).Value
                        Else
                            valorCelula = cel.Value
                        End If
                        wsDestino.Cells(UltimaLinhaDestino, col - 1).Value = valorCelula
                    Next col

                    ' Coloca o nome do hotel na próxima coluna (AL = 38)
                    wsDestino.Cells(UltimaLinhaDestino, 38).Value = nomeHotel
                Next i

                linha = linha + mergeAltura

            Loop

        Next wsOrigem

        wbOrigem.Close SaveChanges:=False
        Arquivo = Dir
    Loop

    MsgBox "Consolidação finalizada com sucesso!"

End Sub
```

### 📋 `replicar-celulas-mescladas.bas`
Replica valores de células mescladas em todas as linhas correspondentes, útil para manter dados completos linha a linha para uso no Power BI.

### 🔍 `marcar-duplicidades-cor.bas`
Verifica duplicidades em uma coluna e aplica cor de fundo nas células duplicadas para facilitar análise visual.

---

## 🧪 Exemplos de uso

📁 Pasta `exemplos/`  
Contém planilhas de exemplo usadas para teste das macros.

🖼️ Pasta `imagens/`  
Contém capturas que demonstram os resultados visuais dos scripts em execução.

---

## 🚀 Como usar

1. Baixe ou clone este repositório.
2. Abra o Excel e pressione `Alt + F11` para acessar o Editor VBA.
3. Vá em `Arquivo > Importar arquivo` e selecione o `.bas` desejado da pasta `scripts`.
4. Execute a macro conforme instruções no código.

---

## 👩‍💻 Sobre mim

Sou Bruna Zordenoni, em transição de carreira para a área de dados. Apaixonada por automatizar processos e extrair valor de planilhas com Power BI, Excel e VBA.

[🔗 LinkedIn](https://www.linkedin.com/in/bruna-zordenoni-096a011b2)
