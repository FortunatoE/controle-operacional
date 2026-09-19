Attribute VB_Name = "modPainelWeb"

Option Explicit

Private Const SH_DASH As String = "DASHBOARD"
Private Const SH_BASE As String = "BASE_CONSOLIDADA"
Private Const SH_PREMISSA As String = "PREMISSA"
Private Const WEB_FOLDER As String = "web"
Private Const EXPORT_FOLDER As String = "EXPORTADOS"
Private Const COL_CODIGO_FAZENDA As String = "R"
Private Const COL_NOME_FAZENDA As String = "S"

' ============================================================================
' PAINEL WEB DE OPERACOES
' Importar este arquivo como um modulo completo. Nao cole procedimentos soltos.
' ============================================================================

Public Sub InstalarPainelWeb()
    Dim ws As Worksheet
    Dim leftPos As Double, topPos As Double

    On Error GoTo TrataErro
    Set ws = ThisWorkbook.Worksheets(SH_DASH)

    ExcluirShape ws, "web_btn_abrir"
    ExcluirShape ws, "web_btn_png"
    ExcluirShape ws, "web_btn_todas"

    leftPos = ws.Range("M2").Left
    topPos = ws.Range("M2").Top

    CriarBotao ws, "web_btn_abrir", "ABRIR PAINEL WEB", leftPos, topPos, 150, 30, RGB(17, 143, 192), "AtualizarEAbrirPainelWeb"
    CriarBotao ws, "web_btn_png", "EXPORTAR PNG HQ", leftPos, topPos + 36, 150, 30, RGB(76, 93, 84), "ExportarPNGWeb"
    CriarBotao ws, "web_btn_todas", "PNG TODAS UNIDADES", leftPos, topPos + 72, 150, 30, RGB(100, 116, 139), "ExportarTodasPNGWeb"

    MsgBox "Botões do painel web instalados.", vbInformation, "Painel Web"
    Exit Sub

TrataErro:
    MsgBox "Erro ao instalar o painel web: " & Err.Description, vbExclamation, "Painel Web"
End Sub

Public Sub AtualizarEAbrirPainelWeb()
    On Error GoTo TrataErro
    Application.ScreenUpdating = False

    ' Atualiza a BASE_CONSOLIDADA quando o modulo de consolidacao tambem existir.
    On Error Resume Next
    Application.Run "'" & ThisWorkbook.Name & "'!AtualizarDados"
    Err.Clear
    On Error GoTo TrataErro

    GerarDadosWeb
    DefinirSelecaoWeb UnidadeSelecionada(), False
    AbrirPainelWeb

Saida:
    Application.ScreenUpdating = True
    Exit Sub

TrataErro:
    MsgBox "Erro ao abrir o painel: " & Err.Description, vbExclamation, "Painel Web"
    Resume Saida
End Sub

Public Sub AbrirPainelWeb()
    Dim arquivo As String
    Dim edge As String
    Dim comando As String

    arquivo = CaminhoWeb() & "\index.html"
    If Dir(arquivo) = vbNullString Then
        MsgBox "index.html não encontrado em: " & arquivo, vbExclamation, "Painel Web"
        Exit Sub
    End If

    CreateObject("Shell.Application").Open arquivo
    edge = LocalizarEdge()
    If Len(edge) = 0 Then
        MsgBox "Microsoft Edge não foi localizado. Abra manualmente: " & arquivo, vbExclamation, "Painel Web"
        Exit Sub
    End If

    comando = Aspas(edge) & " " & Aspas(CaminhoParaFileUrl(arquivo))
    Shell comando, vbNormalFocus
End Sub

Public Sub GerarDadosWeb()
    Dim ws As Worksheet
    Dim ultima As Long, r As Long
    Dim arquivo As String
    Dim f As Integer
    Dim primeiro As Boolean
    Dim registro As String
    Dim codigoFazenda As String
    Dim nomeFazenda As String

    On Error GoTo TrataErro

    Set ws = ThisWorkbook.Worksheets(SH_BASE)
    codigoFazenda = CodigoFazendaDoArquivo()
    nomeFazenda = NomeFazendaDoArquivo()
    GarantirPasta CaminhoWeb()

    arquivo = CaminhoWeb() & "\dados.js"
    f = FreeFile
    Open arquivo For Output As #f

    Print #f, "window.DASHBOARD_DATA = {"
    Print #f, "  geradoEm: """ & JsEscape(Format(Now, "dd/mm/yyyy hh:nn")) & ""","
    Print #f, "  origem: """ & JsEscape(ThisWorkbook.Name) & ""","
    Print #f, "  codigoFazenda: """ & JsEscape(codigoFazenda) & ""","
    Print #f, "  fazenda: """ & JsEscape(nomeFazenda) & ""","
    Print #f, "  registros: ["

    ultima = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    primeiro = True

    For r = 2 To ultima
        If Len(Trim$(CStr(ws.Cells(r, 2).value))) > 0 Then
            registro = "    {" & _
                "unidade:""" & JsEscape(CStr(ws.Cells(r, 1).value)) & """," & _
                "operacao:""" & JsEscape(CStr(ws.Cells(r, 2).value)) & """," & _
                "planejado:" & JsNumero(ws.Cells(r, 3).value) & "," & _
                "realizado:" & JsNumero(ws.Cells(r, 4).value) & "," & _
                "pendente:" & JsNumero(ws.Cells(r, 5).value) & "," & _
                "talhoesConcluidos:" & JsNumero(ws.Cells(r, 8).value) & "," & _
                "talhoesPendentes:" & JsNumero(ws.Cells(r, 9).value) & "}"

            If Not primeiro Then Print #f, ","
            Print #f, registro;
            primeiro = False
        End If
    Next r

    If Not primeiro Then Print #f, ""
    Print #f, "  ]"
    Print #f, "};"
    Close #f

    Exit Sub

TrataErro:
    On Error Resume Next
    If f <> 0 Then Close #f
    On Error GoTo 0
    MsgBox "Erro ao gerar dados web: " & Err.Description, vbExclamation, "Painel Web"
End Sub

Public Sub ExportarPNGWeb()
    On Error GoTo TrataErro
    GerarDadosWeb
    ExportarUnidadeWeb UnidadeSelecionada(), True
    DefinirSelecaoWeb UnidadeSelecionada(), False
    Exit Sub

TrataErro:
    MsgBox "Erro ao exportar PNG: " & Err.Description, vbExclamation, "Painel Web"
End Sub

Public Sub ExportarTodasPNGWeb()
    Dim ws As Worksheet
    Dim unidades As Object
    Dim ultima As Long, r As Long
    Dim unidade As Variant

    On Error GoTo TrataErro
    GerarDadosWeb
    Set ws = ThisWorkbook.Worksheets(SH_BASE)
    Set unidades = CreateObject("Scripting.Dictionary")

    ultima = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    For r = 2 To ultima
        If Len(Trim$(CStr(ws.Cells(r, 1).value))) > 0 Then
            unidades(Trim$(CStr(ws.Cells(r, 1).value))) = True
        End If
    Next r

    For Each unidade In unidades.keys
        ExportarUnidadeWeb CStr(unidade), False
    Next unidade

    DefinirSelecaoWeb "", False
    MsgBox unidades.Count & " PNG(s) gerado(s) em: " & CaminhoExportacao(), vbInformation, "Painel Web"
    Exit Sub

TrataErro:
    DefinirSelecaoWeb "", False
    MsgBox "Erro ao exportar unidades: " & Err.Description, vbExclamation, "Painel Web"
End Sub

Private Sub ExportarUnidadeWeb(ByVal unidade As String, ByVal mostrarMensagem As Boolean)
    Dim edge As String, indexPath As String, outputPath As String
    Dim altura As Long, qtd As Long, comando As String
    Dim retorno As Long

    edge = LocalizarEdge()
    If Len(edge) = 0 Then Err.Raise vbObjectError + 101, , "Microsoft Edge não encontrado."

    indexPath = CaminhoWeb() & "\index.html"
    If Dir(indexPath) = vbNullString Then Err.Raise vbObjectError + 102, , "index.html não encontrado."

    qtd = ContarOperacoes(unidade)
    If qtd = 0 Then Err.Raise vbObjectError + 103, , "Não há operações para a unidade selecionada."

    GarantirPasta CaminhoExportacao()
    DefinirSelecaoWeb unidade, True
    altura = 112 + 26 + ((qtd + 1) \ 2) * 390 + (((qtd + 1) \ 2) - 1) * 26 + 34
    outputPath = CaminhoExportacao() & "\Resumo_Operacoes_" & NomeArquivoSeguro(unidade) & "_" & Format(Now, "yyyymmdd_hhnnss") & ".png"

    comando = Aspas(edge) & " --headless=new --disable-gpu --hide-scrollbars --force-device-scale-factor=2" & _
              " --window-size=1400," & CStr(altura) & " --screenshot=" & Aspas(outputPath) & " " & Aspas(CaminhoParaFileUrl(indexPath))
    retorno = CreateObject("WScript.Shell").Run(comando, 0, True)

    If Dir(outputPath) = vbNullString Then Err.Raise vbObjectError + 104, , "O Edge não gerou o PNG. Código: " & retorno
    If mostrarMensagem Then MsgBox "PNG gerado: " & outputPath, vbInformation, "Painel Web"
End Sub

Private Sub DefinirSelecaoWeb(ByVal unidade As String, ByVal captura As Boolean)
    Dim f As Integer

    GarantirPasta CaminhoWeb()
    f = FreeFile
    Open CaminhoWeb() & "\selecao.js" For Output As #f
    Print #f, "window.DASHBOARD_SELECTED_UNIT = """ & JsEscape(unidade) & """;"
    If captura Then
        Print #f, "window.DASHBOARD_CAPTURE = true;"
    Else
        Print #f, "window.DASHBOARD_CAPTURE = false;"
    End If
    Close #f
End Sub

Private Function UnidadeSelecionada() As String
    UnidadeSelecionada = Trim$(CStr(ThisWorkbook.Worksheets(SH_DASH).Range("B5").value))
    If Len(UnidadeSelecionada) = 0 Then UnidadeSelecionada = "TODAS"
End Function

' Cadastro configurável na aba PREMISSA:
' R2 = CODIGO FAZENDA | S2 = NOME FAZENDA
' R3 = AGV            | S3 = Fazenda Agrovera
Private Function CodigoFazendaDoArquivo() As String
    Dim nomeSemExtensao As String
    Dim ponto As Long

    nomeSemExtensao = Trim$(ThisWorkbook.Name)
    ponto = InStrRev(nomeSemExtensao, ".")
    If ponto > 1 Then nomeSemExtensao = Left$(nomeSemExtensao, ponto - 1)

    If Len(nomeSemExtensao) >= 3 Then
        CodigoFazendaDoArquivo = UCase$(Right$(nomeSemExtensao, 3))
    Else
        CodigoFazendaDoArquivo = UCase$(nomeSemExtensao)
    End If
End Function

Private Function NomeFazendaDoArquivo() As String
    Dim premissa As Worksheet
    Dim codigo As String
    Dim ultima As Long, r As Long

    codigo = CodigoFazendaDoArquivo()
    NomeFazendaDoArquivo = codigo
    If codigo = "AGV" Then NomeFazendaDoArquivo = "Fazenda Agrovera"

    On Error GoTo Saida
    Set premissa = ThisWorkbook.Worksheets(SH_PREMISSA)
    ultima = premissa.Cells(premissa.Rows.Count, COL_CODIGO_FAZENDA).End(xlUp).Row

    For r = 3 To ultima
        If UCase$(Trim$(CStr(premissa.Cells(r, COL_CODIGO_FAZENDA).Value))) = codigo Then
            If Len(Trim$(CStr(premissa.Cells(r, COL_NOME_FAZENDA).Value))) > 0 Then
                NomeFazendaDoArquivo = Trim$(CStr(premissa.Cells(r, COL_NOME_FAZENDA).Value))
            End If
            Exit For
        End If
    Next r

Saida:
End Function

Private Function ContarOperacoes(ByVal unidade As String) As Long
    Dim ws As Worksheet, itens As Object
    Dim ultima As Long, r As Long, operacao As String, unidadeLinha As String

    Set ws = ThisWorkbook.Worksheets(SH_BASE)
    Set itens = CreateObject("Scripting.Dictionary")
    ultima = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    For r = 2 To ultima
        unidadeLinha = Trim$(CStr(ws.Cells(r, 1).value))
        operacao = Trim$(CStr(ws.Cells(r, 2).value))
        If Len(operacao) > 0 Then
            If UCase$(unidade) = "TODAS" Or UCase$(unidadeLinha) = UCase$(unidade) Then itens(UCase$(operacao)) = True
        End If
    Next r

    ContarOperacoes = itens.Count
End Function

Private Sub CriarBotao(ByVal ws As Worksheet, ByVal nome As String, ByVal texto As String, ByVal leftPos As Double, ByVal topPos As Double, ByVal largura As Double, ByVal altura As Double, ByVal cor As Long, ByVal macro As String)
    Dim forma As Shape
    Set forma = ws.Shapes.AddShape(msoShapeRoundedRectangle, leftPos, topPos, largura, altura)
    With forma
        .Name = nome
        .Fill.ForeColor.RGB = cor
        .Line.Visible = msoFalse
        .TextFrame2.TextRange.Text = texto
        .TextFrame2.TextRange.Font.Name = "Aptos"
        .TextFrame2.TextRange.Font.Size = 9
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .OnAction = macro
    End With
End Sub

Private Sub ExcluirShape(ByVal ws As Worksheet, ByVal nome As String)
    On Error Resume Next
    ws.Shapes(nome).Delete
    On Error GoTo 0
End Sub

Private Function CaminhoWeb() As String
    CaminhoWeb = ThisWorkbook.Path & "\" & WEB_FOLDER
End Function

Private Function CaminhoExportacao() As String
    CaminhoExportacao = ThisWorkbook.Path & "\" & EXPORT_FOLDER
End Function

Private Sub GarantirPasta(ByVal pasta As String)
    If Dir(pasta, vbDirectory) = vbNullString Then MkDir pasta
End Sub

Private Function LocalizarEdge() As String
    Dim candidato As String

    candidato = Environ$("ProgramFiles(x86)") & "\Microsoft\Edge\Application\msedge.exe"
    If Dir(candidato) <> vbNullString Then LocalizarEdge = candidato: Exit Function
    candidato = Environ$("ProgramFiles") & "\Microsoft\Edge\Application\msedge.exe"
    If Dir(candidato) <> vbNullString Then LocalizarEdge = candidato: Exit Function
    candidato = Environ$("LOCALAPPDATA") & "\Microsoft\Edge\Application\msedge.exe"
    If Dir(candidato) <> vbNullString Then LocalizarEdge = candidato
End Function

Private Function CaminhoParaFileUrl(ByVal caminho As String) As String
    Dim resultado As String
    resultado = Replace(caminho, "\", "/")
    resultado = Replace(resultado, "%", "%25")
    resultado = Replace(resultado, "#", "%23")
    resultado = Replace(resultado, "?", "%3F")
    resultado = Replace(resultado, " ", "%20")
    CaminhoParaFileUrl = "file:///" & resultado
End Function

Private Function Aspas(ByVal valor As String) As String
    Aspas = Chr$(34) & valor & Chr$(34)
End Function

Private Function JsNumero(ByVal valor As Variant) As String
    Dim numero As String
    If IsError(valor) Or IsEmpty(valor) Or Not IsNumeric(valor) Then
        JsNumero = "0"
        Exit Function
    End If
    numero = Format$(CDbl(valor), "0.############")
    JsNumero = Replace(numero, Application.International(xlDecimalSeparator), ".")
End Function

Private Function JsEscape(ByVal valor As String) As String
    Dim i As Long, caractere As String, codigo As Long, resultado As String

    For i = 1 To Len(valor)
        caractere = Mid$(valor, i, 1)
        codigo = AscW(caractere)
        If codigo < 0 Then codigo = codigo + 65536

        Select Case codigo
            Case 34: resultado = resultado & Chr$(92) & Chr$(34)
            Case 92: resultado = resultado & Chr$(92) & Chr$(92)
            Case 32 To 126: resultado = resultado & caractere
            Case Else: resultado = resultado & Chr$(92) & "u" & Right$("0000" & Hex$(codigo), 4)
        End Select
    Next i

    JsEscape = resultado
End Function

Private Function NomeArquivoSeguro(ByVal valor As String) As String
    Dim item As Variant
    For Each item In Array("\", "/", ":", "*", "?", Chr$(34), "<", ">", "|")
        valor = Replace(valor, CStr(item), "_")
    Next item
    NomeArquivoSeguro = Trim$(valor)
End Function

