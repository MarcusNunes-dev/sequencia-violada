Sub Extrair_API_Nova(coligada As String, DataInicio As String, DataFim As String, login As String, Senha As String)

    Dim http As Object
    Dim url As String
    Dim response As String
    Dim linhas() As String
    Dim i As Long, j As Long
    Dim ws As Worksheet
    Dim campos() As String
    Dim campoNome As String, campoValor As String
    Dim dict As Object
    Dim colunasEsperadas As Variant
    Dim colIndex As Long
    Dim pt As PivotTable
    Dim tbl As ListObject
    Dim abaExiste As Boolean
    Dim sht As Worksheet
    Dim partes() As String
    Dim k As Long, valorCampo As String

    colunasEsperadas = Array( _
        "COLIGADA", "CHAPA", "COLABORADOR", "DT.APURACAO", "PERÍODO", _
        "DIA SEMANA", "SECAO", "PROJETO", "MAO DE OBRA", "SITUACAO", _
        "DATA ADMISSAO", "DATA RESCISAO", "DESCR. CARGO", "ENTRADA", "SEQUENCIA", _
        "SEQUENCIATOTAL", "SAIDA", "ENTRADA1", "SAIDA1", "CLASSIFICACAO")

    ' === URL da API ===
    url = "https://<DOMINIO>/api/framework/v1/consultaSQLServer/<CAMINHO>/A/?parameters=CODCOLIGADA=" & coligada & ";" & "Data_Inicio=" & DataInicio & ";" & "Data_Fim=" & DataFim

    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", url, False
    http.setRequestHeader "Authorization", "Basic " & Base64Encode(login & ":" & Senha)
    http.setRequestHeader "Content-Type", "application/json"
    http.send

    response = http.responseText

    abaExiste = False
    For Each sht In ThisWorkbook.Sheets
        If UCase(sht.Name) = "BASE" Then
            abaExiste = True
            Set ws = sht
            Exit For
        End If
    Next sht

    If Not abaExiste Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "BASE"
    End If

    On Error Resume Next
    Set tbl = ws.ListObjects("Base_dados")
    On Error GoTo 0

    If tbl Is Nothing Then
        Dim ultimaColuna As Long
        Dim ultimaLinha As Long

        For i = 0 To UBound(colunasEsperadas)
            ws.Cells(1, i + 1).Value = colunasEsperadas(i)
        Next i

        ultimaColuna = UBound(colunasEsperadas) + 1

        If ws.Cells(2, 1).Value = "" Then
            ws.Cells(2, 1).Value = "TEMP"
        End If

        ultimaLinha = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

        Set tbl = ws.ListObjects.Add(xlSrcRange, ws.Range(ws.Cells(1, 1), ws.Cells(ultimaLinha, ultimaColuna)), , xlYes)
        tbl.Name = "Base_dados"

        If ws.Cells(2, 1).Value = "TEMP" Then
            tbl.DataBodyRange.Rows(1).Delete
        End If
    End If

    If tbl Is Nothing Then
        ws.Cells.Clear
    Else
        If Not tbl.DataBodyRange Is Nothing Then
            tbl.DataBodyRange.ClearContents
        End If
    End If

    response = Replace(response, "},{", "}§{")
    response = Replace(response, "[", "")
    response = Replace(response, "]", "")
    linhas = Split(response, "§")

    For j = 0 To UBound(colunasEsperadas)
        ws.Cells(1, j + 1).Value = colunasEsperadas(j)
    Next j

    Dim matrizDados() As Variant
    ReDim matrizDados(1 To UBound(linhas) + 1, 1 To UBound(colunasEsperadas) + 1)

    For i = 0 To UBound(linhas)
        linhas(i) = Replace(linhas(i), "{", "")
        linhas(i) = Replace(linhas(i), "}", "")
        campos = Split(linhas(i), ",")

        Set dict = CreateObject("Scripting.Dictionary")

        For j = 0 To UBound(campos)
            If InStr(campos(j), ":") > 0 Then
                partes = Split(campos(j), ":")
                If UBound(partes) >= 1 Then
                    campoNome = Replace(partes(0), """", "")
                    campoValor = ""
                    For k = 1 To UBound(partes)
                        campoValor = campoValor & partes(k)
                        If k < UBound(partes) Then campoValor = campoValor & ":"
                    Next k
                    campoValor = Replace(campoValor, """", "")
                    campoValor = Replace(campoValor, "'", "")

                    If campoNome Like "ENTRADA*" Or campoNome Like "SAIDA*" Then
                        On Error Resume Next
                        campoValor = Format(CDate(campoValor), "hh:mm")
                        On Error GoTo 0
                    End If

                    dict(campoNome) = campoValor
                End If
            End If
        Next j

        For j = 0 To UBound(colunasEsperadas)
            campoNome = colunasEsperadas(j)
            valorCampo = dict(campoNome)

            If campoNome Like "*DATA*" Or campoNome Like "*DT.*" Or campoNome = "PERÍODO" Then
                If InStr(valorCampo, "T") > 0 Then
                    valorCampo = Split(valorCampo, "T")(0)
                End If
                If IsDate(valorCampo) Then
                    matrizDados(i + 1, j + 1) = CDate(valorCampo)
                Else
                    matrizDados(i + 1, j + 1) = valorCampo
                End If
            ElseIf campoNome Like "*ENTRADA*" Or campoNome Like "*SAIDA*" Then
                If IsDate(valorCampo) Then
                    matrizDados(i + 1, j + 1) = Format(CDate(valorCampo), "hh:mm")
                Else
                    matrizDados(i + 1, j + 1) = valorCampo
                End If
            Else
                matrizDados(i + 1, j + 1) = valorCampo
            End If
        Next j
    Next i

    ws.Range("A2").Resize(UBound(matrizDados), UBound(matrizDados, 2)).Value = matrizDados

End Sub
