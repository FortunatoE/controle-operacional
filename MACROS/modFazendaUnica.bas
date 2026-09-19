Attribute VB_Name = "modFazendaUnica"
Option Explicit

Private Const ABA_RESUMO As String = "RESUMO"
Private Const ABA_BASE As String = "BASE_CONSOLIDADA"
Private Const ABA_DASHBOARD As String = "DASHBOARD"
Private Const ABA_PREMISSA As String = "PREMISSA"
Private Const COL_CODIGO_FAZENDA As String = "R"
Private Const COL_NOME_FAZENDA As String = "S"

' A fazenda edita somente este arquivo: OPERACOES - AGV.xlsm.
' Os dados da aba RESUMO alimentam a BASE_CONSOLIDADA e o painel HTML.

Public Sub InstalarPainelFazenda()
    Dim ws As Worksheet, shp As Shape
    Set ws = ThisWorkbook.Worksheets(ABA_DASHBOARD)

    On Error Resume Next
    ws.Shapes("btn_painel_fazenda").Delete
    On Error GoTo 0

    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, ws.Range("M2").Left, ws.Range("M2").Top, 175, 32)
    With shp
        .Name = "btn_painel_fazenda"
        .Fill.ForeColor.RGB = RGB(17, 143, 192)
        .Line.Visible = msoFalse
        .TextFrame2.TextRange.Text = "ATUALIZAR PAINEL WEB"
        .TextFrame2.TextRange.Font.Name = "Aptos"
        .TextFrame2.TextRange.Font.Size = 9
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .OnAction = "AtualizarPainelFazenda"
    End With
End Sub

Public Sub AtualizarPainelFazenda()
    On Error GoTo TrataErro
    Application.ScreenUpdating = False
    Application.CalculateFull
    AtualizarBaseDaFazenda
    GerarDadosWeb
    AbrirPainelWeb

Saida:
    Application.ScreenUpdating = True
    Exit Sub
TrataErro:
    MsgBox "Erro ao atualizar o painel: " & Err.Description, vbExclamation, "Painel da Fazenda"
    Resume Saida
End Sub

Public Sub AtualizarBaseDaFazenda()
    Dim origem As Worksheet, base As Worksheet
    Dim ultima As Long, r As Long, destino As Long
    Dim unidade As String, operacao As String
    Dim planejado As Double, realizado As Double, pendente As Double

    Set origem = ThisWorkbook.Worksheets(ABA_RESUMO)
    Set base = ThisWorkbook.Worksheets(ABA_BASE)
    unidade = UnidadeDaFazenda()

    base.Range("A2:K" & base.Rows.Count).ClearContents
    destino = 2
    ultima = origem.Cells(origem.Rows.Count, "A").End(xlUp).Row

    For r = 4 To ultima
        operacao = Trim$(CStr(origem.Cells(r, 1).Value))
        If Len(operacao) > 0 Then
            planejado = Numero(origem.Cells(r, 2).Value)
            realizado = Numero(origem.Cells(r, 3).Value)
            pendente = Numero(origem.Cells(r, 4).Value)

            base.Cells(destino, 1).Value = unidade
            base.Cells(destino, 2).Value = operacao
            base.Cells(destino, 3).Value = planejado
            base.Cells(destino, 4).Value = realizado
            base.Cells(destino, 5).Value = pendente
            base.Cells(destino, 6).Value = IIf(planejado > 0, realizado / planejado, 0)
            base.Cells(destino, 7).Value = IIf(planejado > 0, pendente / planejado, 0)
            base.Cells(destino, 8).Value = Numero(origem.Cells(r, 6).Value)
            base.Cells(destino, 9).Value = Numero(origem.Cells(r, 7).Value)
            base.Cells(destino, 10).Value = ThisWorkbook.Name
            base.Cells(destino, 11).Value = Now
            destino = destino + 1
        End If
    Next r

    If destino > 2 Then
        base.Range("C2:E" & destino - 1).NumberFormat = "#,##0.0"
        base.Range("F2:G" & destino - 1).NumberFormat = "0%"
        base.Range("K2:K" & destino - 1).NumberFormat = "dd/mm/yyyy hh:mm"
    End If
End Sub

Public Sub GerarDadosWeb()
    Dim ws As Worksheet, ultima As Long, r As Long, f As Integer
    Dim primeiro As Boolean, registro As String
    Dim codigoFazenda As String, nomeFazenda As String

    Set ws = ThisWorkbook.Worksheets(ABA_BASE)
    codigoFazenda = CodigoFazendaDoArquivo()
    nomeFazenda = NomeFazendaDoArquivo()
    GarantirPasta CaminhoWeb()
    f = FreeFile
    Open CaminhoWeb() & "\dados.js" For Output As #f
    Print #f, "window.DASHBOARD_DATA = {"
    Print #f, "  geradoEm: """ & Format(Now, "dd/mm/yyyy hh:nn") & ""","
    Print #f, "  origem: """ & ThisWorkbook.Name & ""","
    Print #f, "  codigoFazenda: """ & Json(codigoFazenda) & ""","
    Print #f, "  fazenda: """ & Json(nomeFazenda) & ""","
    Print #f, "  registros: ["

    ultima = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    primeiro = True
    For r = 2 To ultima
        If Len(Trim$(CStr(ws.Cells(r, 2).Value))) > 0 Then
            registro = "    {unidade:""" & Json(ws.Cells(r, 1).Value) & """,operacao:""" & Json(ws.Cells(r, 2).Value) & """,planejado:" & JsNumero(ws.Cells(r, 3).Value) & ",realizado:" & JsNumero(ws.Cells(r, 4).Value) & ",pendente:" & JsNumero(ws.Cells(r, 5).Value) & ",talhoesConcluidos:" & JsNumero(ws.Cells(r, 8).Value) & ",talhoesPendentes:" & JsNumero(ws.Cells(r, 9).Value) & "}"
            If Not primeiro Then Print #f, ","
            Print #f, registro;
            primeiro = False
        End If
    Next r
    If Not primeiro Then Print #f, ""
    Print #f, "  ]"
    Print #f, "};"
    Close #f
End Sub

Public Sub AbrirPainelWeb()
    Dim arquivo As String, edge As String
    arquivo = CaminhoWeb() & "\index.html"
    If Dir(arquivo) = vbNullString Then Err.Raise vbObjectError + 100, , "index.html não encontrado: " & arquivo
    edge = LocalizarEdge()
    If Len(edge) = 0 Then Err.Raise vbObjectError + 101, , "Microsoft Edge não encontrado."
    Shell Chr$(34) & edge & Chr$(34) & " " & Chr$(34) & FileUrl(arquivo) & Chr$(34), vbNormalFocus
End Sub

Private Function CaminhoWeb() As String
    CaminhoWeb = ThisWorkbook.Path & "\..\web"
End Function

Private Function UnidadeDaFazenda() As String
    UnidadeDaFazenda = Trim$(CStr(ThisWorkbook.Worksheets(ABA_DASHBOARD).Range("B5").Value))
    If Len(UnidadeDaFazenda) = 0 Then UnidadeDaFazenda = "AGV"
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
    Set premissa = ThisWorkbook.Worksheets(ABA_PREMISSA)
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

Private Function Numero(ByVal valor As Variant) As Double
    If IsError(valor) Or Not IsNumeric(valor) Then Exit Function
    Numero = CDbl(valor)
End Function

Private Function JsNumero(ByVal valor As Variant) As String
    JsNumero = Replace(Format$(Numero(valor), "0.############"), Application.International(xlDecimalSeparator), ".")
End Function

Private Function Json(ByVal valor As Variant) As String
    Dim i As Long
    Dim caractere As String
    Dim codigo As Long
    Dim resultado As String

    For i = 1 To Len(CStr(valor))
        caractere = Mid$(CStr(valor), i, 1)
        codigo = AscW(caractere)
        If codigo < 0 Then codigo = codigo + 65536

        Select Case codigo
            Case 34
                resultado = resultado & Chr$(92) & Chr$(34)
            Case 92
                resultado = resultado & Chr$(92) & Chr$(92)
            Case 8
                resultado = resultado & Chr$(92) & "b"
            Case 9
                resultado = resultado & Chr$(92) & "t"
            Case 10, 13
                resultado = resultado & Chr$(92) & "n"
            Case 32 To 126
                resultado = resultado & caractere
            Case Else
                resultado = resultado & Chr$(92) & "u" & Right$("0000" & Hex$(codigo), 4)
        End Select
    Next i

    Json = resultado
End Function

Private Sub GarantirPasta(ByVal pasta As String)
    If Dir(pasta, vbDirectory) = vbNullString Then MkDir pasta
End Sub

Private Function LocalizarEdge() As String
    Dim p As String
    p = Environ$("ProgramFiles(x86)") & "\Microsoft\Edge\Application\msedge.exe"
    If Dir(p) <> vbNullString Then LocalizarEdge = p: Exit Function
    p = Environ$("ProgramFiles") & "\Microsoft\Edge\Application\msedge.exe"
    If Dir(p) <> vbNullString Then LocalizarEdge = p
End Function

Private Function FileUrl(ByVal caminho As String) As String
    FileUrl = "file:///" & Replace(Replace(caminho, "\", "/"), " ", "%20")
End Function
