Sub Painel()
    
    Dim itemID As String, DataInicio As String, DataFim As String
    Dim usuario As String, senha As String
    Dim wsPainel As Worksheet
    With ThisWorkbook.Sheets("Painel")
    
        Application.Calculation = xlCalculationManual
        
        usuario = .Range("B6").Value
        senha = .Range("B7").Value
        Set wsPainel = ThisWorkbook.Sheets("Painel")
    
    End With
    
    Dim linhaInicial As Long: linhaInicial = 15
    Dim linhaFinal As Long: linhaFinal = 18
    
    Dim caminhoArquivo As String, nomeArquivo As String
    Dim anexo As String
    
    Dim j As Long
    For j = linhaInicial To linhaFinal
        If wsPainel.Range("B" & j).Value = "Sim" Then
            
            itemID = wsPainel.Range("A" & j).Value
            DataInicio = wsPainel.Range("C" & j).Value
            DataFim = wsPainel.Range("D" & j).Value
            caminhoArquivo = wsPainel.Range("E" & j).Value
            nomeArquivo = wsPainel.Range("F" & j).Value
            anexo = wsPainel.Range("H" & j).Value
            
            If caminhoArquivo = "" Or nomeArquivo = "" Then
                MsgBox "Caminho ou nome do arquivo ausente na linha " & j & ".", vbExclamation
                GoTo Proximo
            End If
            
            ' Chama rotina de extração de dados, passando parâmetros genéricos
            Call ExtrairDadosAPI(itemID, DataInicio, DataFim, usuario, senha)
            
            ' Chama criação de tabela dinâmica genérica
            Call CriarTabelaDinamicaPersonalizada
            
            ' Chama criação do gráfico genérico
            Call CriarGraficoComTabelaDinamica
            
            If anexo = "Sim" Then
                ' Copiar abas e gerar novo arquivo
                Dim wbAtual As Workbook
                Set wbAtual = ThisWorkbook
                
                Sheets(Array("BASE", "Tabela Dinamica", "Grafico")).Copy
                
                Dim wbNovo As Workbook
                Set wbNovo = Workbooks(Workbooks.Count)
                
                ' Montar caminho completo e salvar arquivo
                If Right(caminhoArquivo, 1) <> "\" Then
                    caminhoArquivo = caminhoArquivo & "\"
                End If
                
                Dim caminhoCompleto As String
                caminhoCompleto = caminhoArquivo & nomeArquivo & ".xlsx"
                
                Application.DisplayAlerts = False
                wbNovo.SaveAs Filename:=caminhoCompleto, FileFormat:=xlOpenXMLWorkbook
                
                ' Atualiza tabelas dinâmicas no novo arquivo
                On Error Resume Next
                Dim ws As Worksheet, pt As PivotTable
                For Each ws In wbNovo.Worksheets
                    For Each pt In ws.PivotTables
                        pt.PivotCache.MissingItemsLimit = xlMissingItemsNone
                        pt.RefreshTable
                    Next pt
                Next ws
                On Error GoTo 0
                
                Application.Calculation = xlCalculationAutomatic
                
                wbNovo.Close SaveChanges:=False
                Application.DisplayAlerts = True
                
Proximo:
            End If
        End If
    Next j

    Application.Calculation = xlCalculationAutomatic

End Sub
