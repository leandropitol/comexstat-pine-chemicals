Attribute VB_Name = "ComexStatMunicipios"
' ======================================================================
' ComexStat Monitor — Pine Chemicals por Município
' RJ Comércio & Alliance Resin Partners
'
' API: POST https://api-comexstat.mdic.gov.br/cities
' Filtro: heading (SH4 como inteiro: 1301, 3805, 3806)
' SSL: WinHttp.WinHttpRequest.5.1 + Option(4)=13056
' Fabricantes: lidos da aba "Fabricantes" — sem precisar editar VBA
' ======================================================================

Sub AtualizarMunicipios()

    Dim wsConf As Worksheet
    Set wsConf = ThisWorkbook.Sheets("Configuracoes")

    ' Lê parâmetros
    Dim anoIni As String: anoIni = CStr(CInt(wsConf.Range("C5").Value))
    Dim anoFim As String: anoFim = CStr(CInt(wsConf.Range("C6").Value))
    Dim mesIni As String: mesIni = Right("0" & CStr(CInt(wsConf.Range("C7").Value)), 2)
    Dim mesFim As String: mesFim = Right("0" & CStr(CInt(wsConf.Range("C8").Value)), 2)
    Dim fluxo  As String: fluxo  = Trim(LCase(wsConf.Range("C9").Value))

    If fluxo <> "export" And fluxo <> "import" Then
        MsgBox "Fluxo invalido em C9. Use 'export' ou 'import'.", vbCritical
        Exit Sub
    End If

    Dim periodoIni As String: periodoIni = anoIni & "-" & mesIni
    Dim periodoFim As String: periodoFim = anoFim & "-" & mesFim

    ' Lê SH4s ativos
    Dim sh4s(10)  As String
    Dim descs(10) As String
    Dim nSh4 As Integer: nSh4 = 0

    Dim r As Integer
    For r = 13 To 30
        Dim sh4Cell As String: sh4Cell = Trim(CStr(wsConf.Cells(r, 2).Value))
        If sh4Cell = "" Then Exit For
        If UCase(Trim(CStr(wsConf.Cells(r, 4).Value))) = "SIM" Then
            sh4s(nSh4)  = sh4Cell
            descs(nSh4) = Trim(CStr(wsConf.Cells(r, 3).Value))
            nSh4 = nSh4 + 1
        End If
    Next r

    If nSh4 = 0 Then MsgBox "Nenhum SH4 ativo.", vbExclamation: Exit Sub

    Application.ScreenUpdating = False

    Dim wsRes As Worksheet
    Set wsRes = ThisWorkbook.Sheets("Resumo")
    wsRes.Cells(2, 2).Value = "Atualizando..."

    Dim totalReg As Long: totalReg = 0
    Dim erros As String:  erros = ""

    Dim i As Integer
    For i = 0 To nSh4 - 1
        ' Pausa antes de TODA requisição (não só entre elas)
        ' Primeira: 3s de cortesia; demais: 15s para respeitar rate limit
        Dim pausaSeg As Integer: pausaSeg = IIf(i = 0, 3, 15)
        Application.StatusBar = "Aguardando " & pausaSeg & "s (rate limit)..."
        Application.Wait Now + TimeSerial(0, 0, pausaSeg)

        Application.StatusBar = "ComexStat Municipios [" & (i + 1) & "/" & nSh4 & "] SH4=" & sh4s(i) & "..."

        Dim nReg As Long
        nReg = ConsultarSH4(sh4s(i), periodoIni, periodoFim, fluxo)

        ' Retry automático se 429
        If nReg = -429 Then
            Application.StatusBar = "Rate limit 429 — aguardando 20s e tentando SH4=" & sh4s(i) & "..."
            Application.Wait Now + TimeSerial(0, 0, 20)
            nReg = ConsultarSH4(sh4s(i), periodoIni, periodoFim, fluxo)
        End If

        If nReg >= 0 Then
            totalReg = totalReg + nReg
            AtualizarResumo wsRes, i + 5, sh4s(i), descs(i), nReg
        Else
            erros = erros & sh4s(i) & " "
        End If
    Next i

    ' Preenche Top_Municipios
    PreencherTop

    wsRes.Cells(2, 2).Value = "Atualizado: " & Format(Now, "dd/mm/yyyy hh:mm") & _
                              "  |  " & periodoIni & " a " & periodoFim & _
                              "  |  Fluxo: " & UCase(fluxo)

    Application.ScreenUpdating = True
    Application.StatusBar = False

    Dim msg As String
    msg = "Atualizacao concluida!" & Chr(10) & Chr(10) & _
          "Total de registros: " & totalReg & Chr(10) & _
          "SH4s consultados: " & nSh4
    If erros <> "" Then msg = msg & Chr(10) & Chr(10) & "Erros: " & erros
    MsgBox msg, vbInformation, "ComexStat Municipios"
End Sub


' ──────────────────────────────────────────────────────────────────────
' Consulta API /cities para um SH4 e grava na aba D_XXXX
' ──────────────────────────────────────────────────────────────────────
Function ConsultarSH4(sh4 As String, periodoIni As String, _
                       periodoFim As String, fluxo As String) As Long

    Const API_URL = "https://api-comexstat.mdic.gov.br/cities"
    Dim q As String: q = Chr(34)

    ' JSON — heading como inteiro (requisito da API /cities)
    Dim js As String
    js = "{" & _
         q & "flow" & q & ":" & q & fluxo & q & "," & _
         q & "monthDetail" & q & ":true," & _
         q & "period" & q & ":{" & _
             q & "from" & q & ":" & q & periodoIni & q & "," & _
             q & "to" & q & ":" & q & periodoFim & q & _
         "}," & _
         q & "filters" & q & ":[{" & _
             q & "filter" & q & ":" & q & "heading" & q & "," & _
             q & "values" & q & ":[" & CStr(CLng(sh4)) & "]" & _
         "}]," & _
         q & "details" & q & ":[" & q & "city" & q & "," & q & "state" & q & "]," & _
         q & "metrics" & q & ":[" & q & "metricFOB" & q & "," & q & "metricKG" & q & "]" & _
         "}"

    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

    On Error GoTo ErroHTTP
    http.Open "POST", API_URL, False
    http.Option(4) = 13056
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Accept", "application/json"
    http.Send js

    ' 429 = rate limit — retorna código especial para retry automático no chamador
    If http.Status = 429 Then
        Set http = Nothing
        ConsultarSH4 = -429
        Exit Function
    End If

    If http.Status <> 200 Then
        MsgBox "HTTP " & http.Status & " — SH4 " & sh4 & Chr(10) & _
               Left(http.responseText, 400), vbCritical, "Erro API"
        ConsultarSH4 = -1: Set http = Nothing: Exit Function
    End If

    Dim resposta As String: resposta = http.responseText
    Set http = Nothing

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("D_" & sh4)
    On Error GoTo 0

    If ws Is Nothing Then
        MsgBox "Aba 'D_" & sh4 & "' nao encontrada.", vbExclamation
        ConsultarSH4 = -1: Exit Function
    End If

    ConsultarSH4 = GravarDados(ws, resposta, sh4)
    Exit Function

ErroHTTP:
    MsgBox "Falha na conexao: " & Err.Description, vbCritical, "Erro de Rede"
    On Error GoTo 0
    Set http = Nothing
    ConsultarSH4 = -1
End Function


' ──────────────────────────────────────────────────────────────────────
' Grava dados na aba de detalhes
' ──────────────────────────────────────────────────────────────────────
Function GravarDados(ws As Worksheet, jsonStr As String, sh4 As String) As Long

    ' Limpa dados anteriores
    Dim ul As Long: ul = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    If ul >= 5 Then ws.Rows("5:" & ul).ClearContents

    If InStr(jsonStr, """list"":[]") > 0 Then
        ws.Cells(5, 2).Value = "Sem dados para o periodo/SH4 consultado"
        GravarDados = 0: Exit Function
    End If

    If InStr(jsonStr, """list"":[{") = 0 Then
        ws.Cells(5, 2).Value = "Resposta inesperada: " & Left(jsonStr, 150)
        GravarDados = 0: Exit Function
    End If

    Dim posListaFim As Long
    Dim posLista As Long: posLista = InStr(jsonStr, """list"":[{")
    posListaFim = InStr(posLista, jsonStr, "]}")

    Dim linhaAtual As Long: linhaAtual = 5
    Dim posAtual As Long: posAtual = InStr(posLista, jsonStr, "{")

    Do While posAtual > 0 And posAtual < posListaFim And linhaAtual <= 5004
        ' Fecha objeto
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

        Dim ano As String: ano = Campo(obj, "year")
        Dim mes As String: mes = Campo(obj, "monthNumber")
        If mes = "" Then mes = Campo(obj, "month")

        ' Municipio: tenta campo combinado primeiro, depois city
        Dim mun As String: mun = Campo(obj, "noMunMinsgUf")
        If mun = "" Then mun = Campo(obj, "city")
        If mun = "" Then mun = "N/D"

        Dim uf  As String: uf  = Campo(obj, "state")
        If uf = "" Then uf = Campo(obj, "sgUfMun")

        Dim fob As Double: fob = ToDouble(Campo(obj, "metricFOB"))
        If fob = 0 Then fob = ToDouble(Campo(obj, "vlFob"))
        Dim kg  As Double: kg  = ToDouble(Campo(obj, "metricKG"))
        If kg  = 0 Then kg  = ToDouble(Campo(obj, "kgLiquido"))

        ' Lookup fabricante na aba Fabricantes
        Dim fab As String: fab = BuscarFabricante(mun, sh4)

        If ano <> "" Then
            ws.Cells(linhaAtual, 2).Value = ToLong(ano)
            ws.Cells(linhaAtual, 3).Value = ToLong(mes)
            ws.Cells(linhaAtual, 4).Value = mun
            ws.Cells(linhaAtual, 5).Value = uf
            ws.Cells(linhaAtual, 6).Value = fab
            ws.Cells(linhaAtual, 7).Value = sh4
            ws.Cells(linhaAtual, 8).Value = fob:  ws.Cells(linhaAtual, 8).NumberFormat = "#,##0.00"
            ws.Cells(linhaAtual, 9).Value = kg:   ws.Cells(linhaAtual, 9).NumberFormat = "#,##0"
            If kg > 0 Then
                ws.Cells(linhaAtual, 10).Value = fob / kg
                ws.Cells(linhaAtual, 10).NumberFormat = "0.0000"
            End If

            ' Zebrado
            Dim bgC As Long: bgC = IIf(linhaAtual Mod 2 = 0, RGB(245, 245, 245), RGB(255, 255, 255))
            Dim col As Integer
            For col = 2 To 10
                ws.Cells(linhaAtual, col).Interior.Color = bgC
                ws.Cells(linhaAtual, col).Font.Name = "Arial"
                ws.Cells(linhaAtual, col).Font.Size = 10
                ws.Cells(linhaAtual, col).HorizontalAlignment = IIf(col <= 7, xlLeft, xlRight)
            Next col
            ws.Cells(linhaAtual, 2).HorizontalAlignment = xlCenter  ' ano
            ws.Cells(linhaAtual, 3).HorizontalAlignment = xlCenter  ' mes
            ws.Cells(linhaAtual, 5).HorizontalAlignment = xlCenter  ' uf
            ws.Cells(linhaAtual, 7).HorizontalAlignment = xlCenter  ' sh4

            linhaAtual = linhaAtual + 1
        End If

        posAtual = InStr(posFim, jsonStr, "{")
    Loop

    GravarDados = linhaAtual - 5
End Function


' ──────────────────────────────────────────────────────────────────────
' Lê a aba "Fabricantes" e retorna o fabricante para município + SH4
' Normalização idêntica ao Python: sem acento, maiúsculas, hífen c/ espaço
' ──────────────────────────────────────────────────────────────────────
Function BuscarFabricante(municipio As String, sh4 As String) As String
    Dim wsFab As Worksheet
    On Error Resume Next
    Set wsFab = ThisWorkbook.Sheets("Fabricantes")
    On Error GoTo 0
    If wsFab Is Nothing Then BuscarFabricante = "OUTROS": Exit Function

    ' Normaliza o município buscado
    Dim munNorm As String: munNorm = NormMun(municipio)
    Dim sh4Str  As String: sh4Str  = Trim(sh4)

    ' Determina coluna do SH4 (col 3=3806, col 4=3805, col 5=1301)
    Dim colFab As Integer
    Select Case sh4Str
        Case "3806": colFab = 3
        Case "3805": colFab = 4
        Case "1301": colFab = 5
        Case Else:   BuscarFabricante = "OUTROS": Exit Function
    End Select

    Dim i As Long
    For i = 4 To 500
        Dim munAba As String: munAba = Trim(CStr(wsFab.Cells(i, 2).Value))
        If munAba = "" Then Exit For
        If NormMun(munAba) = munNorm Then
            BuscarFabricante = Trim(CStr(wsFab.Cells(i, colFab).Value))
            Exit Function
        End If
    Next i

    BuscarFabricante = "OUTROS"
End Function


' ──────────────────────────────────────────────────────────────────────
' Normaliza município: sem acento, maiúsculas, hífen com espaços
' Equivalente ao norm_mun() do Python
' ──────────────────────────────────────────────────────────────────────
Function NormMun(s As String) As String
    If s = "" Then NormMun = "": Exit Function

    Dim result As String: result = s

    ' Substitui espaço não-quebrável e variantes de hífen
    result = Replace(result, Chr(160), " ")
    result = Replace(result, ChrW(8211), "-")  ' en-dash (Unicode — requer ChrW no VBA)
    result = Replace(result, ChrW(8212), "-")  ' em-dash (Unicode — requer ChrW no VBA)

    ' Remove acentos (tabela completa PT-BR)
    Dim acentos As String:   acentos   = "àáâãäåèéêëìíîïòóôõöùúûüýÀÁÂÃÄÅÈÉÊËÌÍÎÏÒÓÔÕÖÙÚÛÜÝçÇñÑ"
    Dim sem As String:       sem       = "aaaaaaeeeeiiiioooooouuuuyAAAAAAAAAEEEEIIIIOOOOOUUUUYcCnN"

    ' Versão compacta — pares explícitos para PT-BR
    Dim pares(29, 1) As String
    pares(0,0)="à": pares(0,1)="a": pares(1,0)="á": pares(1,1)="a"
    pares(2,0)="â": pares(2,1)="a": pares(3,0)="ã": pares(3,1)="a"
    pares(4,0)="ä": pares(4,1)="a": pares(5,0)="è": pares(5,1)="e"
    pares(6,0)="é": pares(6,1)="e": pares(7,0)="ê": pares(7,1)="e"
    pares(8,0)="ë": pares(8,1)="e": pares(9,0)="ì": pares(9,1)="i"
    pares(10,0)="í": pares(10,1)="i": pares(11,0)="î": pares(11,1)="i"
    pares(12,0)="ï": pares(12,1)="i": pares(13,0)="ò": pares(13,1)="o"
    pares(14,0)="ó": pares(14,1)="o": pares(15,0)="ô": pares(15,1)="o"
    pares(16,0)="õ": pares(16,1)="o": pares(17,0)="ö": pares(17,1)="o"
    pares(18,0)="ù": pares(18,1)="u": pares(19,0)="ú": pares(19,1)="u"
    pares(20,0)="û": pares(20,1)="u": pares(21,0)="ü": pares(21,1)="u"
    pares(22,0)="ý": pares(22,1)="y": pares(23,0)="ç": pares(23,1)="c"
    pares(24,0)="ñ": pares(24,1)="n": pares(25,0)="À": pares(25,1)="A"
    pares(26,0)="Á": pares(26,1)="A": pares(27,0)="Â": pares(27,1)="A"
    pares(28,0)="Ã": pares(28,1)="A": pares(29,0)="Ä": pares(29,1)="A"

    Dim k As Integer
    For k = 0 To 29
        result = Replace(result, pares(k, 0), pares(k, 1))
    Next k

    ' Mais maiúsculas BR
    Dim pares2(19, 1) As String
    pares2(0,0)="È": pares2(0,1)="E": pares2(1,0)="É": pares2(1,1)="E"
    pares2(2,0)="Ê": pares2(2,1)="E": pares2(3,0)="Ë": pares2(3,1)="E"
    pares2(4,0)="Ì": pares2(4,1)="I": pares2(5,0)="Í": pares2(5,1)="I"
    pares2(6,0)="Î": pares2(6,1)="I": pares2(7,0)="Ó": pares2(7,1)="O"
    pares2(8,0)="Ô": pares2(8,1)="O": pares2(9,0)="Õ": pares2(9,1)="O"
    pares2(10,0)="Ö": pares2(10,1)="O": pares2(11,0)="Ú": pares2(11,1)="U"
    pares2(12,0)="Û": pares2(12,1)="U": pares2(13,0)="Ü": pares2(13,1)="U"
    pares2(14,0)="Ý": pares2(14,1)="Y": pares2(15,0)="Ç": pares2(15,1)="C"
    pares2(16,0)="Ñ": pares2(16,1)="N": pares2(17,0)="Ò": pares2(17,1)="O"
    pares2(18,0)="Ù": pares2(18,1)="U": pares2(19,0)="Ã": pares2(19,1)="A"

    For k = 0 To 19
        result = Replace(result, pares2(k, 0), pares2(k, 1))
    Next k

    ' Maiúsculas
    result = UCase(Trim(result))

    ' Padroniza hífen: normaliza qualquer variação para " - "
    ' Passo 1: remove espaços duplos ao redor do hífen
    result = Replace(result, "  -  ", " - ")
    result = Replace(result, "  - ", " - ")
    result = Replace(result, " -  ", " - ")
    ' Passo 2: garante espaço antes do hífen se não tiver
    Dim posH As Integer: posH = 1
    Do
        posH = InStr(posH, result, "-")
        If posH = 0 Then Exit Do
        If posH > 1 And Mid(result, posH - 1, 1) <> " " Then
            result = Left(result, posH - 1) & " " & Mid(result, posH)
            posH = posH + 2
        Else
            posH = posH + 1
        End If
    Loop
    ' Passo 3: garante espaço depois do hífen se não tiver
    posH = 1
    Do
        posH = InStr(posH, result, "-")
        If posH = 0 Then Exit Do
        If posH < Len(result) And Mid(result, posH + 1, 1) <> " " Then
            result = Left(result, posH) & " " & Mid(result, posH + 1)
            posH = posH + 3
        Else
            posH = posH + 1
        End If
    Loop

    ' Colapsa espaços múltiplos
    Do While InStr(result, "  ") > 0
        result = Replace(result, "  ", " ")
    Loop

    NormMun = Trim(result)
End Function


' ──────────────────────────────────────────────────────────────────────
' Preenche aba Top_Municipios com ranking consolidado
' ──────────────────────────────────────────────────────────────────────
Sub PreencherTop()
    Dim wsTop As Worksheet
    On Error Resume Next
    Set wsTop = ThisWorkbook.Sheets("Top_Municipios")
    On Error GoTo 0
    If wsTop Is Nothing Then Exit Sub

    ' Limpa dados anteriores
    Dim ulTop As Long: ulTop = wsTop.Cells(wsTop.Rows.Count, 2).End(xlUp).Row
    If ulTop >= 5 Then wsTop.Rows("5:" & ulTop).ClearContents

    ' Coleta dados de todas as abas D_
    Dim sh4s(2) As String: sh4s(0) = "1301": sh4s(1) = "3805": sh4s(2) = "3806"

    ' Estrutura simples: arrays de dados para ordenar depois
    Dim munArr(5000)  As String
    Dim ufArr(5000)   As String
    Dim sh4Arr(5000)  As String
    Dim fabArr(5000)  As String
    Dim fobArr(5000)  As Double
    Dim kgArr(5000)   As Double
    Dim nItens As Long: nItens = 0

    Dim s As Integer
    For s = 0 To 2
        Dim ws As Worksheet
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets("D_" & sh4s(s))
        On Error GoTo 0
        If ws Is Nothing Then GoTo ProxSH4

        Dim ul As Long: ul = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
        Dim r As Long
        For r = 5 To ul
            If CStr(ws.Cells(r, 2).Value) <> "" And CStr(ws.Cells(r, 2).Value) <> "←" Then
                munArr(nItens) = CStr(ws.Cells(r, 4).Value)
                ufArr(nItens)  = CStr(ws.Cells(r, 5).Value)
                fabArr(nItens) = CStr(ws.Cells(r, 6).Value)
                sh4Arr(nItens) = CStr(ws.Cells(r, 7).Value)
                On Error Resume Next
                fobArr(nItens) = CDbl(ws.Cells(r, 8).Value)
                kgArr(nItens)  = CDbl(ws.Cells(r, 9).Value)
                On Error GoTo 0
                nItens = nItens + 1
                If nItens >= 5000 Then GoTo FimColeta
            End If
        Next r
ProxSH4:
    Next s
FimColeta:

    If nItens = 0 Then Exit Sub

    ' Agrupa por municipio+sh4 (bubble sort simples nos top 50)
    ' Para simplificar: ordena por FOB desc usando bubble sort nos primeiros 50 maiores
    Dim maxFob(49) As Double
    Dim maxIdx(49) As Long
    Dim nTop As Integer: nTop = 0

    Dim j As Long
    For j = 0 To nItens - 1
        If nTop < 50 Then
            maxFob(nTop) = fobArr(j)
            maxIdx(nTop) = j
            nTop = nTop + 1
        Else
            ' Encontra o menor
            Dim minVal As Double: minVal = maxFob(0)
            Dim minPos As Integer: minPos = 0
            Dim t As Integer
            For t = 1 To 49
                If maxFob(t) < minVal Then minVal = maxFob(t): minPos = t
            Next t
            If fobArr(j) > minVal Then
                maxFob(minPos) = fobArr(j)
                maxIdx(minPos) = j
            End If
        End If
    Next j

    ' Ordena os top 50 por FOB desc (bubble sort)
    Dim m As Integer, n As Integer
    For m = 0 To nTop - 2
        For n = 0 To nTop - m - 2
            If maxFob(n) < maxFob(n + 1) Then
                Dim tmpF As Double: tmpF = maxFob(n): maxFob(n) = maxFob(n + 1): maxFob(n + 1) = tmpF
                Dim tmpI As Long: tmpI = maxIdx(n): maxIdx(n) = maxIdx(n + 1): maxIdx(n + 1) = tmpI
            End If
        Next n
    Next m

    ' Grava no Top_Municipios
    Dim lRow As Long: lRow = 5
    For m = 0 To nTop - 1
        Dim idx As Long: idx = maxIdx(m)
        Dim bgC As Long: bgC = IIf(lRow Mod 2 = 0, RGB(245, 245, 245), RGB(255, 255, 255))

        wsTop.Cells(lRow, 2).Value = munArr(idx)
        wsTop.Cells(lRow, 3).Value = ufArr(idx)
        wsTop.Cells(lRow, 4).Value = sh4Arr(idx)
        wsTop.Cells(lRow, 5).Value = fabArr(idx)
        wsTop.Cells(lRow, 6).Value = fobArr(idx): wsTop.Cells(lRow, 6).NumberFormat = "#,##0.00"
        wsTop.Cells(lRow, 7).Value = kgArr(idx):  wsTop.Cells(lRow, 7).NumberFormat = "#,##0"
        wsTop.Cells(lRow, 8).Value = m + 1

        Dim col As Integer
        For col = 2 To 8
            wsTop.Cells(lRow, col).Interior.Color = bgC
            wsTop.Cells(lRow, col).Font.Name = "Arial"
            wsTop.Cells(lRow, col).Font.Size = 10
        Next col
        wsTop.Cells(lRow, 2).HorizontalAlignment = xlLeft
        wsTop.Cells(lRow, 3).HorizontalAlignment = xlCenter
        wsTop.Cells(lRow, 4).HorizontalAlignment = xlCenter
        wsTop.Cells(lRow, 5).HorizontalAlignment = xlLeft
        wsTop.Cells(lRow, 6).HorizontalAlignment = xlRight
        wsTop.Cells(lRow, 7).HorizontalAlignment = xlRight
        wsTop.Cells(lRow, 8).HorizontalAlignment = xlCenter

        wsTop.Rows(lRow).RowHeight = 16
        lRow = lRow + 1
    Next m
End Sub


' ──────────────────────────────────────────────────────────────────────
' Atualiza linha do Resumo
' ──────────────────────────────────────────────────────────────────────
Sub AtualizarResumo(wsRes As Worksheet, row As Integer, sh4 As String, _
                    descricao As String, nReg As Long)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("D_" & sh4)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    Dim ul As Long: ul = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    Dim fobTotal As Double: fobTotal = 0
    Dim kgTotal  As Double: kgTotal  = 0
    Dim paises   As New Collection
    Dim i As Long

    On Error Resume Next
    For i = 5 To ul
        fobTotal = fobTotal + CDbl(ws.Cells(i, 8).Value)
        kgTotal  = kgTotal  + CDbl(ws.Cells(i, 9).Value)
        Dim p As String: p = CStr(ws.Cells(i, 4).Value)
        If p <> "" And p <> "N/D" Then paises.Add p, p
    Next i
    On Error GoTo 0

    wsRes.Cells(row, 4).Value = nReg:          wsRes.Cells(row, 4).NumberFormat = "#,##0"
    wsRes.Cells(row, 5).Value = fobTotal:      wsRes.Cells(row, 5).NumberFormat = "#,##0.00"
    wsRes.Cells(row, 6).Value = kgTotal:       wsRes.Cells(row, 6).NumberFormat = "#,##0"
    If kgTotal > 0 Then
        wsRes.Cells(row, 7).Value = fobTotal / kgTotal
        wsRes.Cells(row, 7).NumberFormat = "0.0000"
    End If
    wsRes.Cells(row, 8).Value = paises.Count
End Sub


' ──────────────────────────────────────────────────────────────────────
' Utilitários de parsing JSON
' ──────────────────────────────────────────────────────────────────────
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

Function ToLong(s As String) As Long
    If s = "" Or s = "null" Then ToLong = 0: Exit Function
    On Error Resume Next
    ToLong = CLng(Trim(s))
    If Err.Number <> 0 Then ToLong = 0
    On Error GoTo 0
End Function


' ──────────────────────────────────────────────────────────────────────
' DEBUG — mostra JSON bruto do primeiro objeto
' Alt+F8 → DebugMunicipios → Executar
' ──────────────────────────────────────────────────────────────────────
Sub DebugMunicipios()
    Const API_URL = "https://api-comexstat.mdic.gov.br/cities"
    Dim q As String: q = Chr(34)

    Dim js As String
    js = "{" & q & "flow" & q & ":""export""," & _
         q & "monthDetail" & q & ":true," & _
         q & "period" & q & ":{" & q & "from" & q & ":""2025-01""," & q & "to" & q & ":""2025-01""}," & _
         q & "filters" & q & ":[{" & q & "filter" & q & ":""heading""," & q & "values" & q & ":[1301]}]," & _
         q & "details" & q & ":[""city"",""state""]," & _
         q & "metrics" & q & ":[""metricFOB"",""metricKG""]}"

    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "POST", API_URL, False
    http.Option(4) = 13056
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Accept", "application/json"
    http.Send js

    Dim resp As String: resp = http.responseText
    Set http = Nothing

    Dim pos As Long: pos = InStr(resp, """list"":[{")
    If pos = 0 Then MsgBox "Sem 'list': " & Left(resp, 500), , "Debug": Exit Sub

    pos = InStr(pos, resp, "{")
    Dim pf As Long: pf = InStr(pos + 1, resp, "}")
    MsgBox "Primeiro objeto:" & Chr(10) & Chr(10) & Mid(resp, pos, pf - pos + 1) & _
           Chr(10) & Chr(10) & "Cole numa mensagem para corrigir campos.", _
           vbInformation, "Debug Municipios"
End Sub
