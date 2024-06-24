'Importante é necessário importar a biblioteca word no excel, então aperte Alt+f11
'Insira um novo modulo da macro, vá em ferramentas, referências escolha a biblioteca 
'Microsofit Word 16.0 Object Library.

Option Explicit 

Sub GerarCartas()

    Dim Wd AS New Word.Application
    Dim Doc As Word.Document
    Dim col As Byte, lin AS Byte

    'Inicia a leitura da linha
    lin = 2
    Do Until Planilha1.Cells(lin, 1), = ""

    'Abra o documento
    set doc = wd.Documents.Open(ThisWorkbook.Path & "\cartas.doc")
    wd.Visible = True
    
    'f5 para visualizar

    'Substitua os placeholders com valores de sua planilha
    for col=1 to 15

        with doc.Content.Find
            .Excute Planilha1.Cells(1, col), _
            ReplaceWith:=Planilha1.Cells(lin, col),_
            replace:=wdReplaceAll
        End With

    Next 

        'Salva o documento com nome baseado no valor da célula.
        doc.SaveAs2_ThisWorkbook.Path & "\cartas" & Planilha1.Cells(lin, 1) & ".doc"

        'Aqui vai fechar o documento.
        Doc.Close

        'Incrementa a linha
        lin = lin + 1

    Loop

    'fecha o Word.
    Wd.Quit

    'Limpa as variáveis de objetos. 
    set wd = Nothing
    set Doc = Nothing

End Sub    