Sub GerarChamadosWord()

    ' ================================
    ' DECLARAÇÃO DAS VARIÁVEIS
    ' ================================
    
    Dim wdApp As Object              ' Aplicação Word
    Dim wdDoc As Object              ' Documento Word
    
    Dim ws As Worksheet              ' Planilha com os dados
    Dim ultimaLinha As Long          ' Última linha preenchida
    Dim i As Long                    ' Contador do loop
    
    Dim caminhoPasta As String       ' Pasta do projeto
    Dim caminhoModelo As String      ' Caminho do modelo Word
    Dim nomeArquivo As String        ' Nome final do arquivo Word
    
    Dim horarioFormatado As String   ' Horário convertido para texto
    Dim nomeLimpo As String          ' Nome sem caracteres inválidos
    
    ' ================================
    ' TRATAMENTO DE ERROS
    ' ================================
    
    On Error GoTo TrataErro
    
    ' ================================
    ' CONFIGURAÇÕES INICIAIS
    ' ================================
    
    ' Define a planilha com os dados (aba principal)
    Set ws = ThisWorkbook.Sheets("Gerar Chamados")
    
    ' Pasta base do projeto (ANONIMIZADA)
    caminhoPasta = "C:\Projetos\AutomacaoChamados\"
    
    ' Caminho completo do arquivo modelo Word
    caminhoModelo = caminhoPasta & "MODELO_PADRAO_CHAMADO.docx"
    
    ' Cria a aplicação Word
    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = False            ' Word roda em segundo plano
    
    ' Encontra a última linha preenchida na coluna NOME
    ultimaLinha = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' ================================
    ' LOOP PRINCIPAL
    ' ================================
    
    For i = 2 To ultimaLinha
        
        ' Executa somente se a coluna NOME estiver preenchida
        If ws.Cells(i, "A").Value <> "" Then
            
            ' Abre o documento modelo Word
            Set wdDoc = wdApp.Documents.Open(caminhoModelo)
            
            ' ================================
            ' TRATAMENTO DO HORÁRIO
            ' ================================
            
            ' Converte o valor da coluna HORÁRIO para HH:MM
            If IsDate(ws.Cells(i, "F").Value) Then
                horarioFormatado = Format(ws.Cells(i, "F").Value, "hh:mm")
            Else
                horarioFormatado = ws.Cells(i, "F").Value
            End If
            
            ' ================================
            ' SUBSTITUIÇÃO DOS MARCADORES
            ' ================================
            
            With wdDoc.Content
                
                ' GRUPO
                With .Find
                    .ClearFormatting
                    .Replacement.ClearFormatting
                    .Text = "#GRUPO"
                    .Replacement.Text = ws.Cells(i, "B").Value
                    .Wrap = 1
                    .Execute Replace:=2
                End With
                
                ' LINHA
                With .Find
                    .Text = "#LINHA"
                    .Replacement.Text = ws.Cells(i, "C").Value
                    .Execute Replace:=2
                End With
                
                ' PASSAGEIRO (NOME)
                With .Find
                    .Text = "#PASSAGEIRO"
                    .Replacement.Text = ws.Cells(i, "A").Value
                    .Execute Replace:=2
                End With
                
                ' CENTRO DE CUSTO
                With .Find
                    .Text = "#CC"
                    .Replacement.Text = ws.Cells(i, "K").Value
                    .Execute Replace:=2
                End With
                
                ' DATA DE INÍCIO
                With .Find
                    .Text = "#DATAINICIO"
                    .Replacement.Text = ws.Cells(i, "D").Text
                    .Execute Replace:=2
                End With
                
                ' DATA FINAL
                With .Find
                    .Text = "#DATAFINAL"
                    .Replacement.Text = ws.Cells(i, "E").Text
                    .Execute Replace:=2
                End With
                
                ' ENDEREÇO (Logradouro + Número)
                With .Find
                    .Text = "#ENDERECO"
                    .Replacement.Text = ws.Cells(i, "G").Value & ", " & ws.Cells(i, "H").Value
                    .Execute Replace:=2
                End With
                
                ' BAIRRO
                With .Find
                    .Text = "#BAIRRO"
                    .Replacement.Text = ws.Cells(i, "I").Value
                    .Execute Replace:=2
                End With
                
                ' CIDADE
                With .Find
                    .Text = "#CIDADE"
                    .Replacement.Text = ws.Cells(i, "J").Value
                    .Execute Replace:=2
                End With
                
                ' HORÁRIO DE EMBARQUE
                With .Find
                    .Text = "#EMBARQUE"
                    .Replacement.Text = horarioFormatado
                    .Execute Replace:=2
                End With
                
                ' TELEFONE
                With .Find
                    .Text = "#TELEFONE"
                    .Replacement.Text = ws.Cells(i, "L").Value
                    .Execute Replace:=2
                End With
                
                ' PERÍODO
                With .Find
                    .Text = "#PERIODO"
                    .Replacement.Text = ws.Cells(i, "M").Value
                    .Execute Replace:=2
                End With
                
                ' DETALHE / OBSERVAÇÃO
                With .Find
                    .Text = "#DETALHE"
                    .Replacement.Text = ws.Cells(i, "N").Value
                    .Execute Replace:=2
                End With
                
            End With
            
            ' ================================
            ' MONTAGEM DO NOME DO ARQUIVO
            ' <NOME> - G<GRUPO> - L<LINHA>.docx
            ' ================================
            
            ' Remove caracteres inválidos do nome
            nomeLimpo = LimparNomeArquivo(ws.Cells(i, "A").Value)
            
            nomeArquivo = caminhoPasta & _
                          nomeLimpo & _
                          " - G" & ws.Cells(i, "B").Value & _
                          " - L" & ws.Cells(i, "C").Value & ".docx"
            
            ' Salva o documento Word
            wdDoc.SaveAs nomeArquivo
            
            ' Fecha o documento
            wdDoc.Close
            
        End If
    Next i
    
    ' ================================
    ' FINALIZAÇÃO
    ' ================================
    
    wdApp.Quit
    
    Set wdDoc = Nothing
    Set wdApp = Nothing
    
    MsgBox "Chamados gerados com sucesso!", vbInformation
    Exit Sub

' ================================
' TRATAMENTO DE ERRO
' ================================

TrataErro:
    MsgBox "Erro ao gerar os chamados: " & Err.Description, vbCritical
    
    On Error Resume Next
    If Not wdDoc Is Nothing Then wdDoc.Close False
    If Not wdApp Is Nothing Then wdApp.Quit
    
    Set wdDoc = Nothing
    Set wdApp = Nothing

End Sub

' ================================
' FUNÇÃO PARA LIMPAR NOME DO ARQUIVO
' ================================

Function LimparNomeArquivo(texto As String) As String
    
    Dim caracteresInvalidos As Variant
    Dim i As Integer
    
    caracteresInvalidos = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    
    For i = LBound(caracteresInvalidos) To UBound(caracteresInvalidos)
        texto = Replace(texto, caracteresInvalidos(i), "")
    Next i
    
    LimparNomeArquivo = texto

End Function


