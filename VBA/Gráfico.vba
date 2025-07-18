Sub CriarGraficoComTabelaDinamica()

    Dim wsBase As Worksheet
    Dim wsGrafico As Worksheet
    Dim pc As PivotCache
    Dim pt As PivotTable
    Dim chartObj As ChartObject
    Dim grafico As Chart

    ' === Define aba BASE ===
    Set wsBase = ThisWorkbook.Sheets("BASE")

    ' === Cria nova aba para gráfico ===
    On Error Resume Next
    Set wsGrafico = ThisWorkbook.Sheets("Grafico")
    If wsGrafico Is Nothing Then
        Set wsGrafico = ThisWorkbook.Sheets.Add(After:=wsBase)
        wsGrafico.Name = "Grafico"
    Else
        wsGrafico.Cells.Clear
    End If
    On Error GoTo 0

    ' === Cria PivotCache com base na tabela genérica ===
    Set pc = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:="Tabela_Dados", _               ' Nome genérico da tabela de dados
        Version:=8)

    ' === Cria Tabela Dinâmica na aba Gráfico ===
    Set pt = pc.CreatePivotTable( _
        TableDestination:=wsGrafico.Range("A1"), _
        TableName:="TabelaDinamicaGrafico")

    ' === Define campos da Tabela Dinâmica ===
    With pt
        .RowAxisLayout xlTabularRow
        .RepeatAllLabels xlRepeatLabels
        .ColumnGrand = True
        .RowGrand = True

        ' Limpa campos anteriores (nomes genéricos)
        On Error Resume Next
        .PivotFields("Categoria").Orientation = xlHidden
        .PivotFields("Data").Orientation = xlHidden
        .PivotFields("Dias").Orientation = xlHidden
        .PivotFields("Status").Orientation = xlHidden
        On Error GoTo 0

        ' Linha: Categoria e Data agrupada
        With .PivotFields("Categoria")
            .Orientation = xlRowField
            .Position = 1
        End With

        With pt.PivotFields("Data")
            .Orientation = xlRowField
            .Position = 2
            .AutoGroup
            .Orientation = xlHidden
                       
        On Error GoTo 0
        
        ' Oculta campos auxiliares
        On Error Resume Next
            pt.PivotFields("Dias").Orientation = xlHidden
            pt.PivotFields("Trimestre").Orientation = xlHidden
            
            For Each pf In pt.RowFields
                pf.ShowDetail = True
            Next pf
            
        On Error GoTo 0
                
        End With
        
        With pt
            ' Adiciona campo de valor genérico
            .AddDataField .PivotFields("Status"), "Contagem de Status", xlCount

            ' Garante campo como coluna
            With .PivotFields("Status")
                .Orientation = xlColumnField
                .Position = 1
            End With
        End With
    End With

    ' === Cria gráfico em colunas agrupadas ===
    Set chartObj = wsGrafico.ChartObjects.Add(Left:=50, Top:=50, Width:=700, Height:=400)
    Set grafico = chartObj.Chart

    With grafico
        .SetSourceData Source:=pt.TableRange1
        .ChartType = xlColumnClustered
        .ChartStyle = 209 ' Estilo visual genérico
        .HasTitle = True
        .ChartTitle.Text = "Distribuição por Categoria"
        .Legend.Position = xlLegendPositionRight
    End With

    ' === Oculta itens em branco, se existirem ===
    On Error Resume Next
    pt.PivotFields("Status").PivotItems("(blank)").Visible = False
    On Error GoTo 0

End Sub
