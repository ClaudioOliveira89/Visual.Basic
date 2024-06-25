option Explicit

sub GerarCartas2()

    Dim Wd As New Word.Application
    Dim Doc As Word.Document
    Dim col As Byte, lin As Byte
    Dim SplitText As Variant
    Dim i As Integer

    'inicia a leitura da lina
    lin = 2
    Do Until Planilha1.Cells(lin, 1) = ""

    'A abertura do documento
    Set Doc = Wd.Documents.Opne(ThisWorkbook.Path & "\cartas.docx")
    Wd.Visible = True

    'Substitua os placeholders com valores da planilha
    For col = 1 To 6

    celtext = Planilha1.Cells(lin, col).Value

    'adicionando quebra de parágrafos valor da célula
    If col = 8 Then 
        SplitText = Split(celText, " ")
        celText = Join(splitText, vbCrLf) '>>>> vbCrLf para fazer a quebra de parágrafo
    End If

    'Verifica o tamanho do texto e realiza a substituíções menores
    If Len(celText) > 255 Then 
        spliText = Split(celText, " ")
        For i = LBound(splitText) To UBound(splitText)
            With Doc.Content.Find
                .Text = Planilha1.Cells(1, col).Value
                .Replacement.Text = splicitText(i)
                .Wrap = wdFindContinue
                .Execute Replace:=wdReplaceAll
            End With
        Next i
    Else
        With Doc.Content.Find
            .Text = Planilha1.Cells(1, col).Value
            .Replacement.Text = celText
            .Wrap = wdFindContinue
            .Execute Replace:=wdReplaceAll
        End With
    End if

    Next col

    'Salvar o documento com nome baseado no valor da célula.
    Doc.SaveAs2 ThisWorkbook.Path & "\cartas" & Planilha1.Cells(lin, 1).Value & ".doc"

    Doc.Close 'para fechar o arquivo.

        lin = lin + 1 'para incrementar a linha

    Loop


    Wd.Quit 'para fechar.

'sessão que limpa as variaveis de objetos.
        Set Wd = Nothing
        Set Doc = Nothing

End Sub