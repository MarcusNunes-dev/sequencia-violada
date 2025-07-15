Sub CriarTabelaDinamicaPersonalizada()

    Dim wsBase As Worksheet
    Dim wsTD As Worksheet
    Dim pt As PivotTable
    Dim pc As PivotCache
    Dim rngTabela As Range
    Dim shtNome As String
    Dim tabelaNome As String
    
    ' Nome da tabela e aba base genéricos
    tabelaNome = "NomeDaTabelaBase" ' <-- Ajuste para o nome da sua tabela no Excel
    shtNome = "AbaTabelaDinamica"   ' <-- Ajuste para o nome desejado da aba de tabela dinâmica
    
    ' === Referência à aba base de dados ===
    Set wsBase = ThisWorkbook.Sheets("BASE") ' Ajuste caso sua aba base tenha outro nome
    
    ' === Garante que a aba de Tabela Dinâmica existe ===
    On Error Resume Next
    Set wsTD = ThisWorkbook.Sheets(shtNome)
    If wsTD Is Nothing Then
        Set wsTD = ThisWorkbook.Sheets.Add(After:=wsBase)
        wsTD.Name = shtNome
    Else
        wsTD.Cells.Clear ' Limpa o conteúdo caso já exista
    End If
    On Error GoTo 0
    
    ' === Define o range da tabela base para origem da Tabela Dinâmica ===
    Set rngTabela = wsBase.ListObjects(tabelaNome).Range
    
    ' === Cria o PivotCache ===
    Set pc = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=rngTabela)
    
    ' === Cria a Tabela Dinâmica ===
    Set pt = pc.CreatePivotTable( _
        TableDestination:=wsTD.Range("A3"), _
        TableName:="TabelaDinamicaPersonalizada")
    
    With pt
        Dim campos As Variant
        Dim i As Integer
        
        ' Campos genéricos para linhas - ajuste conforme seus campos reais
        campos = Array("Campo1", "Campo2", "Campo3", "Campo4", "Campo5")
        
        For i = LBound(campos) To UBound(campos)
            With .PivotFields(campos(i))
                .Orientation = xlRowField
                .Position = i + 1
                ' Remove todos os subtotais
                .Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
            End With
        Next i
        
        ' Campo genérico de valores (exemplo: soma do campo "CampoValor")
        On Error Resume Next
        .AddDataField .PivotFields("CampoValor"), "Soma de CampoValor", xlSum
        On Error GoTo 0
        
        ' Estilo tabular e sem rótulos repetidos
        .RowAxisLayout xlTabularRow
        
        ' Remove totais gerais de linhas e colunas
        .RowGrand = False
        .ColumnGrand = False
    End With
    
    ' === Exemplo de cálculo com dados da tabela dinâmica - ajuste conforme necessário ===
    Dim rngValores As Range
    Dim cel As Range
    Dim maxValor As Long
    
    maxValor = 0
    On Error Resume Next
    Set rngValores = pt.DataBodyRange
    On Error GoTo 0
    
    If Not rngValores Is Nothing Then
        For Each cel In rngValores
            If IsNumeric(cel.Value) And cel.Font.Bold = False Then
                If cel.Value > maxValor Then
                    maxValor = cel.Value
                End If
            End If
        Next cel
        
        wsTD.Range("G1").Value = "Máximo Valor"
        wsTD.Range("G2").Value = maxValor
    End If
    
    ' === Criação de segmentações de dados (Slicers) genéricos ===
    Dim slicerCache1 As SlicerCache
    Dim slicerCache2 As SlicerCache
    
    ' Ajuste os nomes dos campos para os que você quer segmentar
    On Error Resume Next
    Set slicerCache1 = ThisWorkbook.SlicerCaches("Slicer_Campo1")
    If Not slicerCache1 Is Nothing Then slicerCache1.Delete
    Set slicerCache1 = ThisWorkbook.SlicerCaches.Add2(pt, "Campo1", "Slicer_Campo1")
    slicerCache1.Slicers.Add wsTD, , "Slicer_Campo1", "Campo1", 10, 10, 150, 200
    
    On Error Resume Next
    Set slicerCache2 = ThisWorkbook.SlicerCaches("Slicer_CampoValor")
    If Not slicerCache2 Is Nothing Then slicerCache2.Delete
    Set slicerCache2 = ThisWorkbook.SlicerCaches.Add2(pt, "CampoValor", "Slicer_CampoValor")
    slicerCache2.Slicers.Add wsTD, , "Slicer_CampoValor", "CampoValor", 180, 10, 150, 200
    
End Sub
