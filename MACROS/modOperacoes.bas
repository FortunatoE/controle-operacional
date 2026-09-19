Attribute VB_Name = "modOperacoes"
Option Explicit

Private Const SH_DASH As String = "DASHBOARD"
Private Const SH_BASE As String = "BASE_CONSOLIDADA"
Private Const SH_CONFIG As String = "CONFIG"
Private Const PREFIXO_EXP As String = "exp_"
Private Const PREFIXO_BTN As String = "btn_"

' ================================================================
' APLICAÇÃO CENTRAL DE ACOMPANHAMENTO DE OPERAÇÕES
' Fluxo: 9 arquivos da pasta de rede -> BASE_CONSOLIDADA -> Dashboard
' ================================================================

Public Sub InstalarAplicacao()
    On Error GoTo TrataErro
    Application.ScreenUpdating = False

    CriarBotoes
    AtualizarListaValidacao
    MsgBox "Aplicação instalada. Agora clique em 'ATUALIZAR DADOS'.", vbInformation, "Operações Agrícolas"

Saida:
    Application.ScreenUpdating = True
    Exit Sub
TrataErro:
    MsgBox "Erro ao instalar: " & Err.Description, vbExclamation
    Resume Saida
End Sub

Public Sub AtualizarDados()
    Dim caminho As String, padrao As String, abaOrigem As String
    Dim linhaCab As Long, linhaIni As Long
    Dim arq As String, caminhoCompleto As String
    Dim wbSrc As Workbook, wsSrc As Worksheet
    Dim wsBase As Worksheet, wsCfg As Worksheet
    Dim ultima As Long, r As Long, outRow As Long
    Dim unidade As String, abriuAqui As Boolean
    Dim op As String, planejado As Double, realizado As Double, pendente As Double
    Dim pExec As Double, pPend As Double, concluidos As Long, pendTalhoes As Long
    Dim calcAnterior As XlCalculation

    On Error GoTo TrataErro
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    calcAnterior = Application.Calculation
    Application.Calculation = xlCalculationManual

    Set wsBase = ThisWorkbook.Worksheets(SH_BASE)
    Set wsCfg = ThisWorkbook.Worksheets(SH_CONFIG)

    caminho = Trim$(CStr(wsCfg.Range("B2").value))
    padrao = Trim$(CStr(wsCfg.Range("B3").value))
    abaOrigem = Trim$(CStr(wsCfg.Range("B4").value))
    linhaCab = CLng(Val(wsCfg.Range("B5").value))
    linhaIni = CLng(Val(wsCfg.Range("B6").value))

    If Len(caminho) = 0 Then Err.Raise vbObjectError + 10, , "Caminho da pasta não configurado."
    If Right$(caminho, 1) <> "\" Then caminho = caminho & "\"
    If Dir(caminho, vbDirectory) = vbNullString Then Err.Raise vbObjectError + 11, , "Pasta não encontrada: " & caminho

    ' Limpa base mantendo cabeçalho
    wsBase.Range("A2:K" & wsBase.Rows.Count).ClearContents
    outRow = 2

    arq = Dir(caminho & padrao)
    Do While Len(arq) > 0
        caminhoCompleto = caminho & arq
        unidade = ExtrairUnidade(arq)

        Set wbSrc = WorkbookAberto(arq)
        abriuAqui = False
        If wbSrc Is Nothing Then
            Set wbSrc = Workbooks.Open(Filename:=caminhoCompleto, UpdateLinks:=False, ReadOnly:=True, AddToMru:=False)
            abriuAqui = True
        End If

        Set wsSrc = Nothing
        On Error Resume Next
        Set wsSrc = wbSrc.Worksheets(abaOrigem)
        On Error GoTo TrataErro

        If Not wsSrc Is Nothing Then
            ultima = wsSrc.Cells(wsSrc.Rows.Count, "A").End(xlUp).Row
            For r = linhaIni To ultima
                op = Trim$(CStr(wsSrc.Cells(r, 1).value))
                If Len(op) > 0 Then
                    planejado = Num(wsSrc.Cells(r, 2).value)
                    realizado = Num(wsSrc.Cells(r, 3).value)
                    pendente = Num(wsSrc.Cells(r, 4).value)

                    If planejado > 0 Then
                        pExec = realizado / planejado
                    Else
                        pExec = Num(wsSrc.Cells(r, 5).value)
                    End If
                    pExec = Limitar01(pExec)
                    pPend = 1 - pExec

                    concluidos = CLng(Num(wsSrc.Cells(r, 6).value))
                    pendTalhoes = CLng(Num(wsSrc.Cells(r, 7).value))

                    wsBase.Cells(outRow, 1).value = unidade
                    wsBase.Cells(outRow, 2).value = op
                    wsBase.Cells(outRow, 3).value = planejado
                    wsBase.Cells(outRow, 4).value = realizado
                    wsBase.Cells(outRow, 5).value = pendente
                    wsBase.Cells(outRow, 6).value = pExec
                    wsBase.Cells(outRow, 7).value = pPend
                    wsBase.Cells(outRow, 8).value = concluidos
                    wsBase.Cells(outRow, 9).value = pendTalhoes
                    wsBase.Cells(outRow, 10).value = arq
                    wsBase.Cells(outRow, 11).value = FileDateTime(caminhoCompleto)
                    outRow = outRow + 1
                End If
            Next r
        End If

        If abriuAqui Then wbSrc.Close SaveChanges:=False
        Set wbSrc = Nothing
        Set wsSrc = Nothing
        arq = Dir()
    Loop

    If outRow = 2 Then
        MsgBox "Nenhum registro foi encontrado nos arquivos da pasta.", vbExclamation
    Else
        wsBase.Range("C2:E" & outRow - 1).NumberFormat = "#,##0.0"
        wsBase.Range("F2:G" & outRow - 1).NumberFormat = "0%"
        wsBase.Range("K2:K" & outRow - 1).NumberFormat = "dd/mm/yyyy hh:mm"
    End If

    AjustarTabelaBase outRow - 1
    AtualizarUnidadesConfig
    AtualizarListaValidacao
    AtualizarDashboard

    MsgBox CStr(outRow - 2) & " registros consolidados com sucesso.", vbInformation, "Operações Agrícolas"

Saida:
    Application.Calculation = calcAnterior
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

TrataErro:
    On Error Resume Next
    If abriuAqui And Not wbSrc Is Nothing Then wbSrc.Close SaveChanges:=False
    On Error GoTo 0
    MsgBox "Erro ao atualizar dados: " & Err.Description, vbCritical, "Operações Agrícolas"
    Resume Saida
End Sub

Public Sub AtualizarDashboard()
    Dim ws As Worksheet, wsBase As Worksheet
    Dim unidade As String
    Dim dict As Object
    Dim lastRow As Long, r As Long
    Dim op As String, key As Variant, arr As Variant
    Dim keys As Variant, i As Long
    Dim planned As Double, realized As Double, pending As Double
    Dim concluidos As Long, pendTalhoes As Long
    Dim x As Double, y As Double, cardW As Double, cardH As Double, gapX As Double, gapY As Double
    Dim cols As Long, idxCol As Long, idxRow As Long
    Dim totalP As Double, totalR As Double, totalPend As Double
    Dim pGeral As Double

    On Error GoTo TrataErro
    Application.ScreenUpdating = False

    Set ws = ThisWorkbook.Worksheets(SH_DASH)
    Set wsBase = ThisWorkbook.Worksheets(SH_BASE)
    unidade = Trim$(CStr(ws.Range("B5").value))
    If Len(unidade) = 0 Then unidade = "TODAS"

    LimparObjetosGerados

    Set dict = CreateObject("Scripting.Dictionary")
    lastRow = wsBase.Cells(wsBase.Rows.Count, "A").End(xlUp).Row

    For r = 2 To lastRow
        If Len(Trim$(CStr(wsBase.Cells(r, 2).value))) > 0 Then
            If UCase$(unidade) = "TODAS" Or UCase$(Trim$(CStr(wsBase.Cells(r, 1).value))) = UCase$(unidade) Then
                op = Trim$(CStr(wsBase.Cells(r, 2).value))
                planned = Num(wsBase.Cells(r, 3).value)
                realized = Num(wsBase.Cells(r, 4).value)
                pending = Num(wsBase.Cells(r, 5).value)
                concluidos = CLng(Num(wsBase.Cells(r, 8).value))
                pendTalhoes = CLng(Num(wsBase.Cells(r, 9).value))

                If dict.Exists(op) Then
                    arr = dict(op)
                    arr(0) = arr(0) + planned
                    arr(1) = arr(1) + realized
                    arr(2) = arr(2) + pending
                    arr(3) = arr(3) + concluidos
                    arr(4) = arr(4) + pendTalhoes
                    dict(op) = arr
                Else
                    dict.Add op, Array(planned, realized, pending, concluidos, pendTalhoes)
                End If
            End If
        End If
    Next r

    CriarCabecalhoExportacao ws, unidade, dict.Count

    If dict.Count = 0 Then
        AddTexto ws, PREFIXO_EXP & "sem_dados", "SEM DADOS PARA A UNIDADE SELECIONADA", ws.Range("A12").Left, ws.Range("A12").Top, 700, 40, 16, True, RGB(120, 120, 120), xlCenter
        Application.ScreenUpdating = True
        Exit Sub
    End If

    keys = dict.keys
    OrdenarPorExecucao keys, dict

    cardW = 360
    cardH = 245
    gapX = 18
    gapY = 18
    cols = 2
    x = ws.Range("A12").Left
    y = ws.Range("A12").Top

    For i = LBound(keys) To UBound(keys)
        key = keys(i)
        arr = dict(key)
        idxCol = i Mod cols
        idxRow = i \ cols

        CriarCard ws, i + 1, CStr(key), CDbl(arr(0)), CDbl(arr(1)), CDbl(arr(2)), CLng(arr(3)), CLng(arr(4)), _
                  x + idxCol * (cardW + gapX), y + idxRow * (cardH + gapY), cardW, cardH

        totalP = totalP + CDbl(arr(0))
        totalR = totalR + CDbl(arr(1))
        totalPend = totalPend + CDbl(arr(2))
    Next i

    If totalP > 0 Then pGeral = totalR / totalP Else pGeral = 0
    ws.Range("A7").value = "Execução geral: " & Format(pGeral, "0%") & "   |   Planejado: " & FormatHa(totalP) & "   |   Realizado: " & FormatHa(totalR) & "   |   Pendente: " & FormatHa(totalPend)
    ws.Range("A7:Q7").Merge
    ws.Range("A7:Q7").HorizontalAlignment = xlCenter
    ws.Range("A7:Q7").Font.Bold = True
    ws.Range("A7:Q7").Font.Color = RGB(77, 77, 77)

    ws.Range("A8").value = "Última atualização: " & Format(Now, "dd/mm/yyyy hh:mm")

Saida:
    Application.ScreenUpdating = True
    Exit Sub
TrataErro:
    MsgBox "Erro ao gerar dashboard: " & Err.Description, vbExclamation
    Resume Saida
End Sub

Public Sub ExportarPNG()
    Dim ws As Worksheet
    Dim unidade As String, pasta As String, nome As String
    Dim nomes() As Variant, n As Long
    Dim shp As Shape, sr As ShapeRange
    Dim co As ChartObject
    Dim maxR As Double, maxB As Double, minL As Double, minT As Double

    On Error GoTo TrataErro
    Application.ScreenUpdating = False
    Set ws = ThisWorkbook.Worksheets(SH_DASH)
    unidade = Trim$(CStr(ws.Range("B5").value))
    If Len(unidade) = 0 Then unidade = "TODAS"

    For Each shp In ws.Shapes
        If Left$(shp.Name, Len(PREFIXO_EXP)) = PREFIXO_EXP Then
            n = n + 1
            ReDim Preserve nomes(1 To n)
            nomes(n) = shp.Name
            If n = 1 Then
                minL = shp.Left: minT = shp.Top
                maxR = shp.Left + shp.Width: maxB = shp.Top + shp.Height
            Else
                If shp.Left < minL Then minL = shp.Left
                If shp.Top < minT Then minT = shp.Top
                If shp.Left + shp.Width > maxR Then maxR = shp.Left + shp.Width
                If shp.Top + shp.Height > maxB Then maxB = shp.Top + shp.Height
            End If
        End If
    Next shp

    If n = 0 Then Err.Raise vbObjectError + 20, , "Gere o dashboard antes de exportar."

    Set sr = ws.Shapes.Range(nomes)
    sr.CopyPicture Appearance:=xlScreen, Format:=xlPicture

    pasta = PastaExportacao()
    GarantirPasta pasta
    nome = pasta & "\Resumo_Operacoes_" & LimparNomeArquivo(unidade) & "_" & Format(Now, "yyyymmdd_hhnn") & ".png"

    Set co = ws.ChartObjects.Add(Left:=maxR + 30, Top:=minT, Width:=maxR - minL, Height:=maxB - minT)
    co.Name = "tmp_export_chart"
    co.Chart.Paste
    co.Chart.Export Filename:=nome, FilterName:="PNG"
    co.Delete

    MsgBox "Imagem gerada:" & vbCrLf & nome, vbInformation, "Exportar PNG"

Saida:
    Application.ScreenUpdating = True
    Exit Sub
TrataErro:
    On Error Resume Next
    If Not co Is Nothing Then co.Delete
    On Error GoTo 0
    MsgBox "Erro ao exportar PNG: " & Err.Description, vbExclamation
    Resume Saida
End Sub

Public Sub GerarTodasFazendas()
    Dim wsCfg As Worksheet, wsDash As Worksheet
    Dim r As Long, unidade As String, unidadeAnterior As String

    On Error GoTo TrataErro
    Application.ScreenUpdating = False
    Set wsCfg = ThisWorkbook.Worksheets(SH_CONFIG)
    Set wsDash = ThisWorkbook.Worksheets(SH_DASH)
    unidadeAnterior = CStr(wsDash.Range("B5").value)

    r = 10
    Do While Len(Trim$(CStr(wsCfg.Cells(r, 1).value))) > 0
        unidade = Trim$(CStr(wsCfg.Cells(r, 1).value))
        wsDash.Range("B5").value = unidade
        AtualizarDashboard
        ExportarPNG
        r = r + 1
    Loop

    wsDash.Range("B5").value = unidadeAnterior
    AtualizarDashboard
    MsgBox "Imagens das unidades concluídas.", vbInformation

Saida:
    Application.ScreenUpdating = True
    Exit Sub
TrataErro:
    MsgBox "Erro ao gerar todas as fazendas: " & Err.Description, vbExclamation
    Resume Saida
End Sub

' ------------------------- VISUAL -------------------------

Private Sub CriarBotoes()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SH_DASH)

    ApagarShapeSeExiste ws, PREFIXO_BTN & "atualizar"
    ApagarShapeSeExiste ws, PREFIXO_BTN & "painel"
    ApagarShapeSeExiste ws, PREFIXO_BTN & "exportar"
    ApagarShapeSeExiste ws, PREFIXO_BTN & "todas"

    CriarBotao ws, PREFIXO_BTN & "atualizar", "ATUALIZAR DADOS", ws.Range("F5").Left, ws.Range("F5").Top, ws.Range("F5:H6").Width, ws.Range("F5:H6").Height, RGB(11, 137, 201), "AtualizarDados"
    CriarBotao ws, PREFIXO_BTN & "painel", "APLICAR UNIDADE", ws.Range("B6").Left, ws.Range("B6").Top + 4, ws.Range("B6:D6").Width, 24, RGB(93, 85, 79), "AtualizarDashboard"
    CriarBotao ws, PREFIXO_BTN & "exportar", "EXPORTAR PNG", ws.Range("J5").Left, ws.Range("J5").Top, ws.Range("J5:L6").Width, ws.Range("J5:L6").Height, RGB(21, 115, 71), "ExportarPNG"
    CriarBotao ws, PREFIXO_BTN & "todas", "GERAR TODAS AS FAZENDAS", ws.Range("N5").Left, ws.Range("N5").Top, ws.Range("N5:Q6").Width, ws.Range("N5:Q6").Height, RGB(107, 114, 128), "GerarTodasFazendas"
End Sub

Private Sub CriarBotao(ByVal ws As Worksheet, ByVal nome As String, ByVal texto As String, ByVal l As Double, ByVal t As Double, ByVal w As Double, ByVal h As Double, ByVal cor As Long, ByVal macro As String)
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, l, t, w, h)
    With shp
        .Name = nome
        .Fill.ForeColor.RGB = cor
        .Line.Visible = msoFalse
        .TextFrame2.TextRange.Text = texto
        .TextFrame2.TextRange.Font.Name = "Aptos"
        .TextFrame2.TextRange.Font.Size = 10
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .OnAction = macro
    End With
End Sub

Private Sub CriarCabecalhoExportacao(ByVal ws As Worksheet, ByVal unidade As String, ByVal qtd As Long)
    Dim bg As Shape, tx As Shape, subtx As Shape
    Set bg = ws.Shapes.AddShape(msoShapeRectangle, ws.Range("A9").Left, ws.Range("A9").Top, 738, 58)
    With bg
        .Name = PREFIXO_EXP & "cab_bg"
        .Fill.ForeColor.RGB = RGB(93, 85, 79)
        .Line.Visible = msoFalse
    End With

    Set tx = AddTexto(ws, PREFIXO_EXP & "cab_titulo", "ACOMPANHAMENTO DE OPERAÇÕES", bg.Left + 16, bg.Top + 8, 470, 25, 18, True, RGB(255, 255, 255), xlLeft)
    Set subtx = AddTexto(ws, PREFIXO_EXP & "cab_sub", "UNIDADE: " & UCase$(unidade) & "  •  " & qtd & " operações  •  " & Format(Now, "dd/mm/yyyy hh:mm"), bg.Left + 16, bg.Top + 32, 620, 18, 9, False, RGB(235, 235, 235), xlLeft)
End Sub

Private Sub CriarCard(ByVal ws As Worksheet, ByVal idx As Long, ByVal operacao As String, ByVal planejado As Double, ByVal realizado As Double, ByVal pendente As Double, ByVal concluidos As Long, ByVal pendTalhoes As Long, ByVal l As Double, ByVal t As Double, ByVal w As Double, ByVal h As Double)
    Dim pExec As Double, pPend As Double
    Dim card As Shape, pill As Shape, progBg As Shape, prog As Shape
    Dim ch As ChartObject
    Dim s As Series
    Dim status As String, corStatus As Long
    Dim donutL As Double, donutT As Double

    If planejado > 0 Then pExec = realizado / planejado Else pExec = 0
    pExec = Limitar01(pExec)
    pPend = 1 - pExec
    status = StatusTexto(pExec)
    corStatus = StatusCor(pExec)

    Set card = ws.Shapes.AddShape(msoShapeRoundedRectangle, l, t, w, h)
    With card
        .Name = PREFIXO_EXP & "card_" & idx
        .Fill.ForeColor.RGB = RGB(255, 255, 255)
        .Line.ForeColor.RGB = RGB(220, 225, 220)
        .Line.Weight = 1
        .Adjustments.item(1) = 0.08
        .Shadow.Visible = msoTrue
        .Shadow.Transparency = 0.86
        .Shadow.Blur = 3
        .Shadow.OffsetX = 0
        .Shadow.OffsetY = 2
    End With

    AddTexto ws, PREFIXO_EXP & "kicker_" & idx, "OPERAÇÃO", l + 18, t + 14, 110, 14, 8, True, RGB(120, 103, 82), xlLeft
    AddTexto ws, PREFIXO_EXP & "nome_" & idx, operacao, l + 18, t + 30, w - 135, 30, 15, True, RGB(28, 32, 34), xlLeft

    Set pill = ws.Shapes.AddShape(msoShapeRoundedRectangle, l + w - 92, t + 14, 75, 27)
    With pill
        .Name = PREFIXO_EXP & "pill_" & idx
        .Fill.ForeColor.RGB = RGB(255, 244, 222)
        If pExec >= 1 Then .Fill.ForeColor.RGB = RGB(226, 244, 235)
        .Line.Visible = msoFalse
        .TextFrame2.TextRange.Text = status
        .TextFrame2.TextRange.Font.Name = "Aptos"
        .TextFrame2.TextRange.Font.Size = 8
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = corStatus
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    End With

    donutL = l + 18
    donutT = t + 68
    Set ch = ws.ChartObjects.Add(Left:=donutL, Top:=donutT, Width:=145, Height:=125)
    ch.Name = PREFIXO_EXP & "chart_" & idx
    With ch.Chart
        .ChartType = xlDoughnut
        .HasLegend = False
        .HasTitle = False
        .ChartArea.Format.Fill.Visible = msoFalse
        .ChartArea.Format.Line.Visible = msoFalse
        .PlotArea.Format.Fill.Visible = msoFalse
        .PlotArea.Format.Line.Visible = msoFalse
        .ChartGroups(1).DoughnutHoleSize = 62
        Set s = .SeriesCollection.NewSeries
        If planejado <= 0 Then
            s.Values = Array(0.0001, 1)
        Else
            s.Values = Array(pExec, pPend)
        End If
        s.Points(1).Format.Fill.ForeColor.RGB = RGB(11, 137, 201)
        s.Points(1).Format.Line.Visible = msoFalse
        s.Points(2).Format.Fill.ForeColor.RGB = RGB(225, 0, 0)
        s.Points(2).Format.Line.Visible = msoFalse
    End With

    AddTexto ws, PREFIXO_EXP & "pct_" & idx, Format(pExec, "0%"), donutL + 39, donutT + 41, 66, 31, 20, True, RGB(33, 38, 40), xlCenter
    AddTexto ws, PREFIXO_EXP & "pct_lbl_" & idx, "EXECUTADO", donutL + 38, donutT + 72, 68, 15, 7, False, RGB(130, 130, 130), xlCenter

    AddMetrica ws, PREFIXO_EXP & "real_" & idx, l + 178, t + 73, RGB(11, 137, 201), "Realizado", Format(pExec, "0%"), FormatHa(realizado)
    AddMetrica ws, PREFIXO_EXP & "pend_" & idx, l + 178, t + 119, RGB(225, 0, 0), "Pendente", Format(pPend, "0%"), FormatHa(pendente)
    AddMetrica ws, PREFIXO_EXP & "plan_" & idx, l + 178, t + 165, RGB(164, 174, 174), "Planejado", "", FormatHa(planejado)

    Set progBg = ws.Shapes.AddShape(msoShapeRoundedRectangle, l + 18, t + h - 47, w - 36, 7)
    With progBg
        .Name = PREFIXO_EXP & "prog_bg_" & idx
        .Fill.ForeColor.RGB = RGB(233, 237, 234)
        .Line.Visible = msoFalse
    End With
    If pExec > 0 Then
        Set prog = ws.Shapes.AddShape(msoShapeRoundedRectangle, l + 18, t + h - 47, (w - 36) * pExec, 7)
        With prog
            .Name = PREFIXO_EXP & "prog_" & idx
            .Fill.ForeColor.RGB = RGB(11, 137, 201)
            .Line.Visible = msoFalse
        End With
    End If

    AddTexto ws, PREFIXO_EXP & "foot1_" & idx, concluidos & " talhões concluídos", l + 18, t + h - 30, 150, 18, 8, False, RGB(105, 105, 105), xlLeft
    AddTexto ws, PREFIXO_EXP & "foot2_" & idx, pendTalhoes & " pendentes", l + w - 130, t + h - 30, 112, 18, 8, False, RGB(105, 105, 105), xlRight
End Sub

Private Sub AddMetrica(ByVal ws As Worksheet, ByVal nome As String, ByVal l As Double, ByVal t As Double, ByVal cor As Long, ByVal rotulo As String, ByVal pct As String, ByVal valor As String)
    Dim dot As Shape
    Set dot = ws.Shapes.AddShape(msoShapeRoundedRectangle, l, t + 8, 9, 9)
    dot.Name = nome & "_dot"
    dot.Fill.ForeColor.RGB = cor
    dot.Line.Visible = msoFalse

    AddTexto ws, nome & "_lbl", IIf(Len(pct) > 0, rotulo & "    " & pct, rotulo), l + 17, t, 145, 17, 8, False, RGB(105, 105, 105), xlLeft
    AddTexto ws, nome & "_val", valor, l + 17, t + 17, 145, 22, 12, True, RGB(20, 24, 25), xlLeft
End Sub

Private Function AddTexto(ByVal ws As Worksheet, ByVal nome As String, ByVal texto As String, ByVal l As Double, ByVal t As Double, ByVal w As Double, ByVal h As Double, ByVal tam As Double, ByVal negrito As Boolean, ByVal cor As Long, ByVal alinhamento As XlHAlign) As Shape
    Dim shp As Shape
    Set shp = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, l, t, w, h)
    With shp
        .Name = nome
        .Fill.Visible = msoFalse
        .Line.Visible = msoFalse
        .TextFrame2.TextRange.Text = texto
        .TextFrame2.TextRange.Font.Name = "Aptos"
        .TextFrame2.TextRange.Font.Size = tam
        .TextFrame2.TextRange.Font.Bold = IIf(negrito, msoTrue, msoFalse)
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = cor
        .TextFrame2.MarginLeft = 0
        .TextFrame2.MarginRight = 0
        .TextFrame2.MarginTop = 0
        .TextFrame2.MarginBottom = 0
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        Select Case alinhamento
            Case xlCenter: .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
            Case xlRight: .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignRight
            Case Else: .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignLeft
        End Select
    End With
    Set AddTexto = shp
End Function

Private Sub LimparObjetosGerados()
    Dim ws As Worksheet, i As Long
    Set ws = ThisWorkbook.Worksheets(SH_DASH)
    For i = ws.Shapes.Count To 1 Step -1
        If Left$(ws.Shapes(i).Name, Len(PREFIXO_EXP)) = PREFIXO_EXP Then ws.Shapes(i).Delete
    Next i
End Sub

' ------------------------- BASE / CONFIG -------------------------

Private Sub AtualizarUnidadesConfig()
    Dim wsB As Worksheet, wsC As Worksheet, dict As Object
    Dim lastRow As Long, r As Long, u As String, arr As Variant, i As Long
    Set wsB = ThisWorkbook.Worksheets(SH_BASE)
    Set wsC = ThisWorkbook.Worksheets(SH_CONFIG)
    Set dict = CreateObject("Scripting.Dictionary")

    wsC.Range("A10:A200").ClearContents
    lastRow = wsB.Cells(wsB.Rows.Count, "A").End(xlUp).Row
    For r = 2 To lastRow
        u = Trim$(CStr(wsB.Cells(r, 1).value))
        If Len(u) > 0 Then If Not dict.Exists(u) Then dict.Add u, u
    Next r

    arr = dict.keys
    For i = LBoundSafe(arr) To UBoundSafe(arr)
        wsC.Cells(10 + i, 1).value = arr(i)
    Next i
End Sub

Private Sub AtualizarListaValidacao()
    Dim wsD As Worksheet, wsC As Worksheet
    Dim r As Long, lista As String, u As String
    Set wsD = ThisWorkbook.Worksheets(SH_DASH)
    Set wsC = ThisWorkbook.Worksheets(SH_CONFIG)

    lista = "TODAS"
    r = 10
    Do While Len(Trim$(CStr(wsC.Cells(r, 1).value))) > 0
        u = Trim$(CStr(wsC.Cells(r, 1).value))
        lista = lista & "," & u
        r = r + 1
    Loop

    With wsD.Range("B5")
        .Validation.Delete
        .Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=lista
        .Validation.IgnoreBlank = True
        .Validation.InCellDropdown = True
    End With
End Sub

Private Sub AjustarTabelaBase(ByVal lastDataRow As Long)
    Dim ws As Worksheet, lo As ListObject, lr As Long
    Set ws = ThisWorkbook.Worksheets(SH_BASE)
    On Error Resume Next
    Set lo = ws.ListObjects("tbBase")
    On Error GoTo 0
    If lo Is Nothing Then Exit Sub
    lr = IIf(lastDataRow < 2, 2, lastDataRow)
    lo.Resize ws.Range("A1:K" & lr)
End Sub

Private Function WorkbookAberto(ByVal nomeArquivo As String) As Workbook
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If StrComp(wb.Name, nomeArquivo, vbTextCompare) = 0 Then
            Set WorkbookAberto = wb
            Exit Function
        End If
    Next wb
End Function

Private Function ExtrairUnidade(ByVal nomeArquivo As String) As String
    Dim s As String
    s = nomeArquivo
    s = Replace(s, ".xlsx", "", 1, -1, vbTextCompare)
    s = Replace(s, ".xlsm", "", 1, -1, vbTextCompare)
    s = Replace(s, "OPERAÇÕES - ", "", 1, -1, vbTextCompare)
    ExtrairUnidade = Trim$(s)
End Function

' ------------------------- EXPORTAÇÃO -------------------------

Private Function PastaExportacao() As String
    Dim subpasta As String
    subpasta = Trim$(CStr(ThisWorkbook.Worksheets(SH_CONFIG).Range("B7").value))
    If Len(subpasta) = 0 Then subpasta = "EXPORTADOS"
    If Len(ThisWorkbook.Path) = 0 Then
        PastaExportacao = Environ$("USERPROFILE") & "\Desktop\" & subpasta
    Else
        PastaExportacao = ThisWorkbook.Path & "\" & subpasta
    End If
End Function

Private Sub GarantirPasta(ByVal caminho As String)
    If Dir(caminho, vbDirectory) = vbNullString Then MkDir caminho
End Sub

Private Function LimparNomeArquivo(ByVal s As String) As String
    Dim chars As Variant, c As Variant
    chars = Array("\", "/", ":", "*", "?", Chr$(34), "<", ">", "|")
    For Each c In chars
        s = Replace(s, CStr(c), "-")
    Next c
    LimparNomeArquivo = s
End Function

' ------------------------- AUXILIARES -------------------------

Private Function Num(ByVal v As Variant) As Double
    If IsError(v) Or IsEmpty(v) Or Len(Trim$(CStr(v))) = 0 Then
        Num = 0
    ElseIf IsNumeric(v) Then
        Num = CDbl(v)
    Else
        Num = 0
    End If
End Function

Private Function Limitar01(ByVal v As Double) As Double
    If v < 0 Then v = 0
    If v > 1 Then v = 1
    Limitar01 = v
End Function

Private Function FormatHa(ByVal v As Double) As String
    If Abs(v - Round(v, 0)) < 0.001 Then
        FormatHa = Format(v, "#,##0") & " ha"
    Else
        FormatHa = Format(v, "#,##0.0") & " ha"
    End If
End Function

Private Function StatusTexto(ByVal p As Double) As String
    If p >= 0.9999 Then
        StatusTexto = "CONCLUÍDO"
    ElseIf p >= 0.8 Then
        StatusTexto = "FINALIZANDO"
    ElseIf p >= 0.5 Then
        StatusTexto = "EM ANDAMENTO"
    Else
        StatusTexto = "ATENÇÃO"
    End If
End Function

Private Function StatusCor(ByVal p As Double) As Long
    If p >= 0.9999 Then
        StatusCor = RGB(21, 115, 71)
    ElseIf p >= 0.8 Then
        StatusCor = RGB(86, 103, 37)
    Else
        StatusCor = RGB(163, 101, 0)
    End If
End Function

Private Sub OrdenarPorExecucao(ByRef keys As Variant, ByVal dict As Object)
    Dim i As Long, j As Long, tmp As Variant
    Dim a As Variant, b As Variant, pa As Double, pb As Double
    If Not IsArray(keys) Then Exit Sub
    For i = LBound(keys) To UBound(keys) - 1
        For j = i + 1 To UBound(keys)
            a = dict(keys(i)): b = dict(keys(j))
            If a(0) > 0 Then pa = a(1) / a(0) Else pa = 0
            If b(0) > 0 Then pb = b(1) / b(0) Else pb = 0
            If pa > pb Then
                tmp = keys(i): keys(i) = keys(j): keys(j) = tmp
            End If
        Next j
    Next i
End Sub

Private Sub ApagarShapeSeExiste(ByVal ws As Worksheet, ByVal nome As String)
    On Error Resume Next
    ws.Shapes(nome).Delete
    On Error GoTo 0
End Sub

Private Function LBoundSafe(ByVal arr As Variant) As Long
    On Error GoTo vazio
    LBoundSafe = LBound(arr)
    Exit Function
vazio:
    LBoundSafe = 0
End Function

Private Function UBoundSafe(ByVal arr As Variant) As Long
    On Error GoTo vazio
    UBoundSafe = UBound(arr)
    Exit Function
vazio:
    UBoundSafe = -1
End Function
