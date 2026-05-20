Attribute VB_Name = "ComexStatMonitor"
' ======================================================================
' ComexStat Monitor — Pine Chemicals
' RJ Comércio & Alliance Resin Partners
'
' API: POST https://api-comexstat.mdic.gov.br/general
' HTTP: WinHttp.WinHttpRequest.5.1  (nativo Windows, sem dependência)
' SSL:  Option(4) = 13056  (ignora ICP-Brasil, igual ao verify=False)
' ======================================================================

Sub AtualizarComexStat()
    Dim wsConf As Worksheet
    Set wsConf = ThisWorkbook.Sheets("Configuracoes")

    Dim anoIni As String, anoFim As String
    Dim mesIni As String, mesFim As String
    Dim fluxo  As String

    anoIni = CStr(CInt(wsConf.Range("C5").Value))
    anoFim = CStr(CInt(wsConf.Range("C6").Value))
    mesIni = Right("0" & CStr(CInt(wsConf.Range("C7").Value)), 2)
    mesFim = Right("0" & CStr(CInt(wsConf.Range("C8").Value)), 2)
    fluxo  = Trim(LCase(wsConf.Range("C9").Value))

    If fluxo <> "export" And fluxo <> "import" Then
        MsgBox "Fluxo invalido em C9. Use 'export' ou 'import'.", vbCritical
        Exit Sub
    End If

    Dim periodoIni As String: periodoIni = anoIni & "-" & mesIni
    Dim periodoFim As String: periodoFim = anoFim & "-" & mesFim

    ' Le NCMs da tabela (linhas 13 a 30, coluna B=NCM, D=Ativo)
    Dim ncms(10) As String
    Dim descs(10) As String
    Dim nNcms As Integer: nNcms = 0

    Dim rowNcm As Integer
    For rowNcm = 13 To 30
        Dim ncmCell As String
        ncmCell = Trim(CStr(wsConf.Cells(rowNcm, 2).Value))
        If ncmCell = "" Then Exit For
        If UCase(Trim(CStr(wsConf.Cells(rowNcm, 4).Value))) = "SIM" Then
            ncms(nNcms)  = Replace(Replace(ncmCell, ".", ""), " ", "")
            descs(nNcms) = Trim(CStr(wsConf.Cells(rowNcm, 3).Value))
            nNcms = nNcms + 1
        End If
    Next rowNcm

    If nNcms = 0 Then
        MsgBox "Nenhum NCM ativo encontrado.", vbExclamation
        Exit Sub
    End If

    Dim wsRes As Worksheet
    Set wsRes = ThisWorkbook.Sheets("Resumo")

    Application.ScreenUpdating = False

    Dim totalReg As Long: totalReg = 0
    Dim erros As String:  erros = ""

    Dim i As Integer
    For i = 0 To nNcms - 1
        If i > 0 Then Application.Wait Now + TimeSerial(0, 0, 12)
        Application.StatusBar = "ComexStat [" & (i + 1) & "/" & nNcms & "] " & ncms(i) & "..."
        Dim nReg As Long
        nReg = ConsultarNCM(ncms(i), periodoIni, periodoFim, fluxo)
        If nReg >= 0 Then
            totalReg = totalReg + nReg
            AtualizarResumo wsRes, i + 5, ncms(i), descs(i), nReg
        Else
            erros = erros & ncms(i) & " "
        End If
    Next i

    wsRes.Cells(2, 2).Value = "Atualizado: " & Format(Now, "dd/mm/yyyy hh:mm") & _
                              "  |  " & periodoIni & " a " & periodoFim & _
                              "  |  Fluxo: " & UCase(fluxo)

    Application.ScreenUpdating = True
    Application.StatusBar = False

    Dim msg As String
    msg = "Atualizacao concluida!" & Chr(10) & Chr(10) & _
          "Total de registros: " & totalReg & Chr(10) & _
          "NCMs consultados: " & nNcms
    If erros <> "" Then msg = msg & Chr(10) & Chr(10) & "Erros em: " & erros
    MsgBox msg, vbInformation, "ComexStat Monitor"
End Sub


Function ConsultarNCM(ncm As String, periodoIni As String, _
                       periodoFim As String, fluxo As String) As Long
    Const API_URL = "https://api-comexstat.mdic.gov.br/general"
    Dim q As String: q = Chr(34)

    ' JSON identico ao Colab que funcionou
    Dim js As String
    js = "{" & _
         q & "flow" & q & ":" & q & fluxo & q & "," & _
         q & "monthDetail" & q & ":true," & _
         q & "period" & q & ":{" & _
             q & "from" & q & ":" & q & periodoIni & q & "," & _
             q & "to" & q   & ":" & q & periodoFim & q & _
         "}," & _
         q & "filters" & q & ":[{" & _
             q & "filter" & q & ":" & q & "ncm" & q & "," & _
             q & "values" & q & ":[" & q & ncm & q & "]" & _
         "}]," & _
         q & "details" & q & ":[" & q & "country" & q & "," & q & "ncm" & q & "]," & _
         q & "metrics" & q & ":[" & _
             q & "metricFOB" & q & "," & _
             q & "metricKG" & q & "," & _
             q & "metricStatistic" & q & _
         "]}"

    ' WinHttp.WinHttpRequest.5.1 — nativo Windows, sem registro extra
    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

    On Error GoTo ErroHTTP

    http.Open "POST", API_URL, False
    http.Option(4) = 13056          ' WinHttpRequestOption_SslErrorIgnoreFlags
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Accept", "application/json"
    http.Send js

    If http.Status <> 200 Then
        MsgBox "HTTP " & http.Status & " — NCM " & ncm & Chr(10) & _
               Left(http.responseText, 400), vbCritical, "Erro API"
        ConsultarNCM = -1
        Set http = Nothing
        Exit Function
    End If

    Dim resposta As String: resposta = http.responseText
    Set http = Nothing

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("D_" & ncm)
    On Error GoTo 0

    If ws Is Nothing Then
        MsgBox "Aba 'D_" & ncm & "' nao encontrada.", vbExclamation
        ConsultarNCM = -1
        Exit Function
    End If

    ConsultarNCM = GravarDados(ws, resposta)
    Exit Function

ErroHTTP:
    MsgBox "Falha na conexao: " & Err.Description, vbCritical, "Erro de Rede"
    On Error GoTo 0
    Set http = Nothing
    ConsultarNCM = -1
End Function


Function GravarDados(ws As Worksheet, jsonStr As String) As Long
    ' Limpa dados anteriores
    Dim ul As Long
    ul = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    If ul >= 5 Then ws.Rows("5:" & ul).ClearContents

    If InStr(jsonStr, """list"":[]") > 0 Then
        ws.Cells(5, 2).Value = "Sem dados para o periodo/NCM consultado"
        GravarDados = 0
        Exit Function
    End If

    If InStr(jsonStr, """list"":[{") = 0 Then
        ws.Cells(5, 2).Value = "Resposta inesperada: " & Left(jsonStr, 150)
        GravarDados = 0
        Exit Function
    End If

    ' Limite de segurança do array "list"
    Dim posListaFim As Long
    Dim posLista As Long: posLista = InStr(jsonStr, """list"":[{")
    posListaFim = InStr(posLista, jsonStr, "]}")

    Dim linhaAtual As Long: linhaAtual = 5
    Dim posAtual As Long: posAtual = InStr(posLista, jsonStr, "{")

    Do While posAtual > 0 And posAtual < posListaFim And linhaAtual <= 5004
        ' Fecha objeto respeitando aninhamento
        Dim prof As Integer: prof = 0
        Dim posFim As Long: posFim = posAtual
        Dim ch As String
        Do
            ch = Mid(jsonStr, posFim, 1)
            If ch = "{" Then prof = prof + 1
            If ch = "}" Then prof = prof - 1
            posFim = posFim + 1
        Loop While prof > 0 And posFim <= Len(jsonStr)

        Dim obj As String: obj = Mid(jsonStr, posAtual, posFim - posAtual)

        Dim ano  As String: ano  = Campo(obj, "year")
        Dim mes  As String: mes  = Campo(obj, "monthNumber")
        Dim pais As String: pais = Campo(obj, "country")
        Dim ncmN As String: ncmN = Campo(obj, "ncm")
        Dim fob  As String: fob  = Campo(obj, "metricFOB"):  If fob  = "" Then fob  = Campo(obj, "vlFob")
        Dim kg   As String: kg   = Campo(obj, "metricKG"):   If kg   = "" Then kg   = Campo(obj, "kgLiquido")
        Dim qtd  As String: qtd  = Campo(obj, "metricStatistic"): If qtd  = "" Then qtd  = Campo(obj, "qtEstat")

        If ano <> "" Then
            Dim fobV As Double: fobV = ToDouble(fob)
            Dim kgV  As Double: kgV  = ToDouble(kg)

            ws.Cells(linhaAtual, 2).Value = ToLong(ano)
            ws.Cells(linhaAtual, 3).Value = ToLong(mes)
            ws.Cells(linhaAtual, 4).Value = pais
            ws.Cells(linhaAtual, 5).Value = ncmN
            ws.Cells(linhaAtual, 6).Value = fobV
            ws.Cells(linhaAtual, 6).NumberFormat = "#,##0.00"
            ws.Cells(linhaAtual, 7).Value = kgV
            ws.Cells(linhaAtual, 7).NumberFormat = "#,##0"
            If kgV > 0 Then
                ws.Cells(linhaAtual, 8).Value = fobV / kgV
                ws.Cells(linhaAtual, 8).NumberFormat = "0.0000"
            End If
            ws.Cells(linhaAtual, 9).Value = ToDouble(qtd)
            ws.Cells(linhaAtual, 9).NumberFormat = "#,##0"

            Dim bgC As Long
            bgC = IIf(linhaAtual Mod 2 = 0, RGB(245, 245, 245), RGB(255, 255, 255))
            Dim c As Integer
            For c = 2 To 9
                ws.Cells(linhaAtual, c).Interior.Color = bgC
                ws.Cells(linhaAtual, c).Font.Name = "Arial"
                ws.Cells(linhaAtual, c).Font.Size = 10
                ws.Cells(linhaAtual, c).HorizontalAlignment = IIf(c <= 5, xlLeft, xlRight)
            Next c
            ws.Cells(linhaAtual, 2).HorizontalAlignment = xlCenter
            ws.Cells(linhaAtual, 3).HorizontalAlignment = xlCenter

            linhaAtual = linhaAtual + 1
        End If

        posAtual = InStr(posFim, jsonStr, "{")
    Loop

    GravarDados = linhaAtual - 5
End Function


Function Campo(obj As String, nome As String) As String
    Dim q As String: q = Chr(34)
    Dim pos As Long: pos = InStr(obj, q & nome & q)
    If pos = 0 Then Campo = "": Exit Function
    pos = InStr(pos, obj, ":") + 1
    Do While Mid(obj, pos, 1) = " ": pos = pos + 1: Loop
    If Mid(obj, pos, 4) = "null" Then Campo = "": Exit Function
    Dim valor As String
    If Mid(obj, pos, 1) = q Then
        pos = pos + 1
        Dim pf As Long: pf = InStr(pos, obj, q)
        If pf = 0 Then Campo = "": Exit Function
        valor = Mid(obj, pos, pf - pos)
    Else
        Dim pn As Long: pn = pos
        Do While pn <= Len(obj) And InStr(",}]", Mid(obj, pn, 1)) = 0: pn = pn + 1: Loop
        valor = Trim(Mid(obj, pos, pn - pos))
    End If
    Campo = valor
End Function


Function ToDouble(s As String) As Double
    If s = "" Or s = "null" Then ToDouble = 0: Exit Function
    On Error Resume Next
    ToDouble = CDbl(Replace(s, ",", "."))
    If Err.Number <> 0 Then ToDouble = 0
    On Error GoTo 0
End Function


Sub AtualizarResumo(wsRes As Worksheet, row As Integer, ncm As String, _
                    descricao As String, nReg As Long)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("D_" & ncm)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    Dim ul As Long: ul = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    Dim fobTotal As Double: fobTotal = 0
    Dim kgTotal  As Double: kgTotal  = 0
    Dim paises   As New Collection
    Dim i As Long

    On Error Resume Next
    For i = 5 To ul
        fobTotal = fobTotal + CDbl(ws.Cells(i, 6).Value)
        kgTotal  = kgTotal  + CDbl(ws.Cells(i, 7).Value)
        Dim p As String: p = CStr(ws.Cells(i, 4).Value)
        If p <> "" Then paises.Add p, p
    Next i
    On Error GoTo 0

    wsRes.Cells(row, 4).Value = nReg:         wsRes.Cells(row, 4).NumberFormat = "#,##0"
    wsRes.Cells(row, 5).Value = fobTotal:     wsRes.Cells(row, 5).NumberFormat = "#,##0.00"
    wsRes.Cells(row, 6).Value = kgTotal:      wsRes.Cells(row, 6).NumberFormat = "#,##0"
    If kgTotal > 0 Then
        wsRes.Cells(row, 7).Value = fobTotal / kgTotal
        wsRes.Cells(row, 7).NumberFormat = "0.0000"
    End If
    wsRes.Cells(row, 8).Value = paises.Count
End Sub


Function ToLong(s As String) As Long
    If s = "" Or s = "null" Then ToLong = 0: Exit Function
    On Error Resume Next
    ToLong = CLng(Trim(s))
    If Err.Number <> 0 Then ToLong = 0
    On Error GoTo 0
End Function


' ──────────────────────────────────────────────────────────────────────
' DEBUG — Mostra os primeiros 800 chars do JSON bruto no Imediato
' Execute: Alt+F8 → DebugJSON → Executar
' Depois verifique a janela Imediato (Ctrl+G no VBA Editor)
' ──────────────────────────────────────────────────────────────────────
Sub DebugJSON()
    Const API_URL = "https://api-comexstat.mdic.gov.br/general"
    Dim q As String: q = Chr(34)

    Dim js As String
    js = "{" & _
         q & "flow" & q & ":" & q & "export" & q & "," & _
         q & "monthDetail" & q & ":true," & _
         q & "period" & q & ":{" & q & "from" & q & ":""2025-01""," & q & "to" & q & ":""2025-01""}," & _
         q & "filters" & q & ":[{" & q & "filter" & q & ":" & q & "ncm" & q & "," & q & "values" & q & ":[" & q & "13019090" & q & "]}]," & _
         q & "details" & q & ":[" & q & "country" & q & "," & q & "ncm" & q & "]," & _
         q & "metrics" & q & ":[" & q & "metricFOB" & q & "," & q & "metricKG" & q & "]}"

    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "POST", API_URL, False
    http.Option(4) = 13056
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Accept", "application/json"
    http.Send js

    Dim resp As String: resp = http.responseText
    Set http = Nothing

    ' Mostra o primeiro objeto do array "list" para ver os nomes dos campos
    Dim posObj As Long: posObj = InStr(resp, """list"":[{")
    If posObj = 0 Then
        Debug.Print "Sem 'list' na resposta:"
        Debug.Print Left(resp, 500)
        MsgBox Left(resp, 800), vbInformation, "JSON bruto"
        Exit Sub
    End If

    posObj = InStr(posObj, resp, "{")
    Dim posFim As Long: posFim = InStr(posObj + 1, resp, "}")
    Dim primeiroObj As String
    primeiroObj = Mid(resp, posObj, posFim - posObj + 1)

    Debug.Print "=== PRIMEIRO OBJETO DO LIST ==="
    Debug.Print primeiroObj
    Debug.Print "=== 500 chars do inicio ==="
    Debug.Print Left(resp, 500)

    MsgBox "Primeiro objeto:" & Chr(10) & Chr(10) & primeiroObj & Chr(10) & Chr(10) & _
           "Cole isso numa mensagem para corrigir os campos.", _
           vbInformation, "Debug JSON"
End Sub
