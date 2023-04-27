'Can be useful for someone (English Version)

'create the document
    Set objWord = CreateObject("Word.Application")

'makes the document visible
    objWord.Visible = True

'select document from folder
    Set arqProcuração = objWord.Documents.Open(ThisWorkbook.Path & "\Modelo.procuração.docx")
    Set conteudoDoc = arqProcuração.Application.Selection
    
'selects columns all to the right
    ultCol = Range("C4").End(xlToRight).Column

'diz que para cada coltab a coluna é igual ao 3 até a última
    For coltab = 3 To ultCol

'selects the top cells (find), and completes with the bottom ones (replacement) within the word.
    conteudoDoc.Find.Text = Cells(4, coltab).Value
    conteudoDoc.Find.Replacement.Text = Cells(5, coltab).Value
    conteudoDoc.Find.Execute Replace:=wdReplaceAll
    
Next

'back to ADI
    Sheets("ADI").Select

'select the columns all to the right
    ultCol = Range("C9").End(xlToRight).Column

'say that for each coltab the column is equal to 4 until the last
    For coltab = 4 To ultCol

'selects the top cells (find), and completes with the bottom ones (replacement) within the word.
    conteudoDoc.Find.Text = Cells(9, coltab).Value
    conteudoDoc.Find.Replacement.Text = Cells(10, coltab).Value
    conteudoDoc.Find.Execute Replace:=wdReplaceAll
    
Next
    
'go to the lawyer's area
    Sheets("Aréa do Advogado").Select
    
'selects columns all to the right
    ultCol = Range("D3").End(xlToRight).Column

'says that for each coltab the column is equal to 3 until the last
    For coltab = 3 To ultCol

'selects the top cells (find), and completes with the bottom ones (replacement) within the word.
    conteudoDoc.Find.Text = Cells(3, coltab).Value
    conteudoDoc.Find.Replacement.Text = Cells(4, coltab).Value
    conteudoDoc.Find.Execute Replace:=wdReplaceAll

Next

'back to ADI
    Sheets("ADI").Select

'save the file and give it as valid
    arqProcuração.SaveAs2 (ThisWorkbook.Path & "\Procuração - " & Cells(5, 3).Value & ".docx")

'Close the file
    arqProcuração.Close

'Close the word
    objWord.Quit

'bureaucracy
    Set arqProcuração = Nothing
    Set conteudoDoc = Nothing
    Set objWord = Nothing

'message that you are ready
MsgBox ("Procuração gerada com sucesso!")

End Sub