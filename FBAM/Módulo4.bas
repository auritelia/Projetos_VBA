Attribute VB_Name = "Módulo4"
Dim f As Worksheet
Dim P As Worksheet
Dim S As Worksheet
Dim NFicha As Integer
Dim Linhas As Long
Dim Categoria As String
Dim SubCategoria As String
Dim Material As String
Dim Cor As String
Dim TecProd As String
Dim Decora As String
Dim Integridade As String
Dim Estado As String
Dim Interv As String
Dim Acondici As String
Dim Armazena As String
Dim TipoAcervo As String
Dim ContColunas As Integer
Dim FBAM As Workbook

Sub Preencher_Fichas()

Application.ScreenUpdating = False

'teste de procedimento para preencher e copiar cada ficha

Set S = Sheets("Sitio")
Set P = Sheets("Preenchimento")
Set f = Sheets("FICHA")
Set FBAM = Application.ThisWorkbook

'Calcula a qtd de linhas da tabela
Linhas = P.Cells(Rows.Count, 1).End(xlUp).Row
Linhas = Linhas - 4
    
f.Visible = True
f.Unprotect ("zaza")

'Preenche e cola todas as fichas no arquivo:

For NFicha = 1 To Linhas
    
    'Para mudar o Nº da ficha na ficha e atualizar todos os procv:
    f.Range("D71") = P.Cells(NFicha + 4, 1).Value
    
    'Atribui o valor da célula de baixo para as variaveies de preenchimento:
    
    Categoria = P.Cells(NFicha + 4, 17).Value
    SubCategoria = P.Cells(NFicha + 4, 21).Value
    Material = P.Cells(NFicha + 4, 12).Value
    Cor = P.Cells(NFicha + 4, 47).Value
    TecProd = P.Cells(NFicha + 4, 23).Value
    Decora = P.Cells(NFicha + 4, 30).Value
    Integridade = P.Cells(NFicha + 4, 56).Value
    Estado = P.Cells(NFicha + 4, 57).Value
    Interv = P.Cells(NFicha + 4, 62).Value
    Acondici = P.Cells(NFicha + 4, 67).Value
    Armazena = P.Cells(NFicha + 4, 71).Value
    TipoAcervo = P.Cells(NFicha + 4, 6).Value
    
    
'--------------------------------------------------------------------------------------------------------

'Preenchimentos:

'Preenchimento 3. Categoria:

f.Select
f.Range("C14").Select


Do Until ActiveCell.Value = 0
    If ActiveCell = Categoria Then
    
            ActiveCell.Offset(0, -1).Value = "X"
            ActiveCell.Offset(1, 0).Select
    
        Else
    
            ActiveCell.Offset(1, 0).Select
    
    
    End If
Loop

'Preenchimento da segunda coluna:

f.Range("E14").Select


Do Until ActiveCell.Value = 0
    If ActiveCell = Categoria Then
    
            ActiveCell.Offset(0, -1).Value = "X"
            ActiveCell.Offset(1, 0).Select
    
        Else
    
            ActiveCell.Offset(1, 0).Select
    
    
    End If
Loop


'Preenchimento da terceira coluna:

f.Range("G14").Select


Do Until ActiveCell.Value = 0
    If ActiveCell = Categoria Then
    
            ActiveCell.Offset(0, -1).Value = "X"
            ActiveCell.Offset(1, 0).Select

    
        Else
    
            ActiveCell.Offset(1, 0).Select
    
    
    End If
Loop


'Procedimento para preencher "outros":

If f.Range("F15").Value = "X" Then

f.Range("G16") = P.Cells(NFicha + 4, 18).Value

Else

End If



            '----------------------------------------------------------------------
            
'Preenchimento 4. Subcategoria:


f.Select
f.Range("C18").Select


Do Until ActiveCell.Value = 0
    If ActiveCell = SubCategoria Then
    
            ActiveCell.Offset(0, -1).Value = "X"
            ActiveCell.Offset(1, 0).Select
    
        Else
    
            ActiveCell.Offset(1, 0).Select
    
    
    End If
Loop

'Preenchimento da segunda coluna:

f.Range("E18").Select


Do Until ActiveCell.Value = 0
    If ActiveCell = SubCategoria Then
    
            ActiveCell.Offset(0, -1).Value = "X"
            ActiveCell.Offset(1, 0).Select

    
        Else
    
            ActiveCell.Offset(1, 0).Select
    
    
    End If
Loop


'Preenchimento da terceira coluna:

f.Range("G18").Select


Do Until ActiveCell.Value = 0
    If ActiveCell = SubCategoria Then
    
           ActiveCell.Offset(0, -1).Value = "X"
            ActiveCell.Offset(1, 0).Select
    
        Else
    
            ActiveCell.Offset(1, 0).Select
    
    
    End If
Loop


'Preenchimento da quarta coluna:

f.Range("I18").Select


Do Until ActiveCell.Value = 0
    If ActiveCell = SubCategoria Then
    
            ActiveCell.Offset(0, -1).Value = "X"
            ActiveCell.Offset(1, 0).Select

    
        Else
    
            ActiveCell.Offset(1, 0).Select
    
    
    End If
Loop



'Procedimento para preencher "outros":

If f.Range("H20").Value = "X" Then

f.Range("I21") = P.Cells(NFicha + 4, 22).Value

Else

End If



            '-----------------------------------------------------------------------
            
            
'Preenchimento 5.Material:
   

ContColunas = 12

Do While ContColunas < 16

f.Select
f.Range("C23").Select


'Preenchimento da primeira coluna:

Do Until ActiveCell.Value = 0
    If ActiveCell = Material Then
    
            ActiveCell.Offset(0, -1).Value = "X"
            ActiveCell.Offset(1, 0).Select
    
        Else
    
            ActiveCell.Offset(1, 0).Select
    
    
    End If
Loop

'Preenchimento da segunda coluna:

f.Range("E23").Select


Do Until ActiveCell.Value = 0
    If ActiveCell = Material Then
    
            ActiveCell.Offset(0, -1).Value = "X"
            ActiveCell.Offset(1, 0).Select
    
        Else
    
            ActiveCell.Offset(1, 0).Select
    
    
    End If
Loop


'Preenchimento da terceira coluna:

f.Range("G23").Select


Do Until ActiveCell.Value = 0
    If ActiveCell = Material Then
    
            ActiveCell.Offset(0, -1).Value = "X"
            ActiveCell.Offset(1, 0).Select
    
        Else
    
            ActiveCell.Offset(1, 0).Select
    
    
    End If
Loop


'Preenchimento da quarta coluna:

f.Range("I23").Select


Do Until ActiveCell.Value = 0
    If ActiveCell = Material Then
    
            ActiveCell.Offset(0, -1).Value = "X"
            ActiveCell.Offset(1, 0).Select
    
        Else
    
            ActiveCell.Offset(1, 0).Select
    
    
    End If
Loop

'Atribuir novo valor de Material para a variável:

ContColunas = ContColunas + 1
Material = P.Cells(NFicha + 4, ContColunas).Value


Loop



'Procedimento para preencher "outros":

If f.Range("H25").Value = "X" Then

f.Range("I26") = P.Cells(NFicha + 4, 16).Value


Else

End If

            '----------------------------------------------------------------------------------
            

'Preenchimento 6. Cor:

ContColunas = 47

Do While ContColunas < 49

f.Select
f.Range("C30").Select


'Preenchimento da primeira coluna:

Do Until ActiveCell.Value = 0
    If ActiveCell = Cor Then
    
            ActiveCell.Offset(0, -1).Value = "X"
            ActiveCell.Offset(1, 0).Select
    
        Else
    
            ActiveCell.Offset(1, 0).Select
    
    
    End If
Loop

'Preenchimento da segunda coluna:

f.Range("E30").Select


Do Until ActiveCell.Value = 0
    If ActiveCell = Cor Then
    
            ActiveCell.Offset(0, -1).Value = "X"
            ActiveCell.Offset(1, 0).Select
    
        Else
    
            ActiveCell.Offset(1, 0).Select
    
    
    End If
Loop


'Preenchimento da terceira coluna:

f.Range("G30").Select


Do Until ActiveCell.Value = 0
    If ActiveCell = Cor Then
    
            ActiveCell.Offset(0, -1).Value = "X"
            ActiveCell.Offset(1, 0).Select
    
        Else
    
            ActiveCell.Offset(1, 0).Select
    
    
    End If
Loop


'Preenchimento da quarta coluna:

f.Range("I30").Select


Do Until ActiveCell.Value = 0
    If ActiveCell = Cor Then
    
            ActiveCell.Offset(0, -1).Value = "X"
            ActiveCell.Offset(1, 0).Select
    
        Else
    
            ActiveCell.Offset(1, 0).Select
    
    
    End If
Loop

'Atribuir novo valor de Cor para a variável:

ContColunas = ContColunas + 1
Cor = P.Cells(NFicha + 4, ContColunas).Value


Loop



'Procedimento para preencher "outros":

If f.Range("H30").Value = "X" Then

f.Range("I31") = P.Cells(NFicha + 4, 49).Value


Else

End If



            '--------------------------------------------------------------------

'Preenchimento 7. Técnicas de Produção

ContColunas = 23

Do While ContColunas < 29

f.Select
f.Range("C33").Select


'Preenchimento da primeira coluna:

Do Until ActiveCell.Value = 0
    If ActiveCell = TecProd Then
    
            ActiveCell.Offset(0, -1).Value = "X"
            ActiveCell.Offset(1, 0).Select
    
        Else
    
            ActiveCell.Offset(1, 0).Select
    
    
    End If
Loop

'Preenchimento da segunda coluna:

f.Range("E33").Select


Do Until ActiveCell.Value = 0
    If ActiveCell = TecProd Then
    
            ActiveCell.Offset(0, -1).Value = "X"
            ActiveCell.Offset(1, 0).Select
    
        Else
    
            ActiveCell.Offset(1, 0).Select
    
    
    End If
Loop


'Preenchimento da terceira coluna:

f.Range("G33").Select


Do Until ActiveCell.Value = 0
    If ActiveCell = TecProd Then
    
            ActiveCell.Offset(0, -1).Value = "X"
            ActiveCell.Offset(1, 0).Select
    
        Else
    
            ActiveCell.Offset(1, 0).Select
    
    
    End If
Loop


'Preenchimento da quarta coluna:

f.Range("I33").Select


Do Until ActiveCell.Value = 0
    If ActiveCell = TecProd Then
    
            ActiveCell.Offset(0, -1).Value = "X"
            ActiveCell.Offset(1, 0).Select
    
        Else
    
            ActiveCell.Offset(1, 0).Select
    
    
    End If
Loop

'Atribuir novo valor de Material para a variável:

ContColunas = ContColunas + 1
TecProd = P.Cells(NFicha + 4, ContColunas).Value


Loop



'Procedimento para preencher "outros":

If f.Range("H35").Value = "X" Then

f.Range("I36") = P.Cells(NFicha + 4, 29).Value


Else

End If


            '----------------------------------------------------------------------------

'Preenchimento 8. Decoração:

ContColunas = 30

Do While ContColunas < 46

f.Select
f.Range("C38").Select


'Preenchimento da primeira coluna:

Do Until ActiveCell.Value = 0
    If ActiveCell = Decora Then
    
            ActiveCell.Offset(0, -1).Value = "X"
            ActiveCell.Offset(1, 0).Select
    
        Else
    
            ActiveCell.Offset(1, 0).Select
    
    
    End If
Loop

'Preenchimento da segunda coluna:

f.Range("E38").Select


Do Until ActiveCell.Value = 0
    If ActiveCell = Decora Then
    
            ActiveCell.Offset(0, -1).Value = "X"
            ActiveCell.Offset(1, 0).Select
    
        Else
    
            ActiveCell.Offset(1, 0).Select
    
    
    End If
Loop


'Preenchimento da terceira coluna:

f.Range("G38").Select


Do Until ActiveCell.Value = 0
    If ActiveCell = Decora Then
    
            ActiveCell.Offset(0, -1).Value = "X"
            ActiveCell.Offset(1, 0).Select
    
        Else
    
            ActiveCell.Offset(1, 0).Select
    
    
    End If
Loop


'Preenchimento da quarta coluna:

f.Range("I38").Select


Do Until ActiveCell.Value = 0
    If ActiveCell = Decora Then
    
            ActiveCell.Offset(0, -1).Value = "X"
            ActiveCell.Offset(1, 0).Select
    
        Else
    
            ActiveCell.Offset(1, 0).Select
    
    
    End If
Loop

'Atribuir novo valor de Material para a variável:

ContColunas = ContColunas + 1
Decora = P.Cells(NFicha + 4, ContColunas).Value


Loop



'Procedimento para preencher "outros":

If f.Range("H41").Value = "X" Then

f.Range("I42") = P.Cells(NFicha + 4, 46).Value


Else

End If


            '--------------------------------------------------------------------------------

'Preenchimento 9. Integridade:


f.Select
f.Range("C44").Select


Do Until ActiveCell.Value = 0
    If ActiveCell = Integridade Then
    
            ActiveCell.Offset(0, -1).Value = "X"
            ActiveCell.Offset(1, 0).Select
    
        Else
    
            ActiveCell.Offset(1, 0).Select
    
    
    End If
Loop

'Preenchimento da segunda coluna:

f.Range("E44").Select


Do Until ActiveCell.Value = 0
    If ActiveCell = Integridade Then
    
            ActiveCell.Offset(0, -1).Value = "X"
            ActiveCell.Offset(1, 0).Select

    
        Else
    
            ActiveCell.Offset(1, 0).Select
    
    
    End If
Loop


'Preenchimento da terceira coluna:

f.Range("G44").Select


Do Until ActiveCell.Value = 0
    If ActiveCell = Integridade Then
    
           ActiveCell.Offset(0, -1).Value = "X"
            ActiveCell.Offset(1, 0).Select
    
        Else
    
            ActiveCell.Offset(1, 0).Select
    
    
    End If
Loop



            '--------------------------------------------------------------------------
            
            
'Preenchimento 10.Estado de conservação:


ContColunas = 57

Do While ContColunas < 61

f.Select
f.Range("C46").Select


'Preenchimento da primeira coluna:

Do Until ActiveCell.Value = 0
    If ActiveCell = Estado Then
    
            ActiveCell.Offset(0, -1).Value = "X"
            ActiveCell.Offset(1, 0).Select
    
        Else
    
            ActiveCell.Offset(1, 0).Select
    
    
    End If
Loop

'Preenchimento da segunda coluna:

f.Range("G46").Select


Do Until ActiveCell.Value = 0
    If ActiveCell = Estado Then
    
            ActiveCell.Offset(0, -1).Value = "X"
            ActiveCell.Offset(1, 0).Select
    
        Else
    
            ActiveCell.Offset(1, 0).Select
    
    
    End If
Loop


'Atribuir novo valor de Material para a variável:

ContColunas = ContColunas + 1
Estado = P.Cells(NFicha + 4, ContColunas).Value


Loop



'Procedimento para preencher "outros":

If f.Range("F46").Value = "X" Then

f.Range("G47") = P.Cells(NFicha + 4, 61).Value


Else

End If


            '--------------------------------------------------------------------------
            

'Preenchimento 11. Intervenções sofridas:

ContColunas = 62

Do While ContColunas < 65

f.Select
f.Range("C51").Select


'Preenchimento da primeira coluna:

Do Until ActiveCell.Value = 0
    If ActiveCell = Interv Then
    
            ActiveCell.Offset(0, -1).Value = "X"
            ActiveCell.Offset(1, 0).Select
    
        Else
    
            ActiveCell.Offset(1, 0).Select
    
    
    End If
Loop

'Preenchimento da segunda coluna:

f.Range("E51").Select


Do Until ActiveCell.Value = 0
    If ActiveCell = Interv Then
    
            ActiveCell.Offset(0, -1).Value = "X"
            ActiveCell.Offset(1, 0).Select
    
        Else
    
            ActiveCell.Offset(1, 0).Select
    
    
    End If
Loop


'Preenchimento da terceira coluna:

f.Range("G51").Select


Do Until ActiveCell.Value = 0
    If ActiveCell = Interv Then
    
            ActiveCell.Offset(0, -1).Value = "X"
            ActiveCell.Offset(1, 0).Select
    
        Else
    
            ActiveCell.Offset(1, 0).Select
    
    
    End If
Loop


'Preenchimento da quarta coluna:

f.Range("I51").Select


Do Until ActiveCell.Value = 0
    If ActiveCell = Interv Then
    
            ActiveCell.Offset(0, -1).Value = "X"
            ActiveCell.Offset(1, 0).Select
    
        Else
    
            ActiveCell.Offset(1, 0).Select
    
    
    End If
Loop

'Atribuir novo valor de Material para a variável:

ContColunas = ContColunas + 1
Interv = P.Cells(NFicha + 4, ContColunas).Value


Loop



'Procedimento para preencher "outros":

If f.Range("H51").Value = "X" Then

f.Range("I52") = P.Cells(NFicha + 4, 65).Value


Else

End If


            '------------------------------------------------------------------------
'Preenchimento 13. Invólucro/Acondicionamento:

ContColunas = 67

Do While ContColunas < 70

f.Select
f.Range("C57").Select


'Preenchimento da primeira coluna:

Do Until ActiveCell.Value = 0
    If ActiveCell = Acondici Then
    
            ActiveCell.Offset(0, -1).Value = "X"
            ActiveCell.Offset(1, 0).Select
    
        Else
    
            ActiveCell.Offset(1, 0).Select
    
    
    End If
Loop

'Preenchimento da segunda coluna:

f.Range("E57").Select


Do Until ActiveCell.Value = 0
    If ActiveCell = Acondici Then
    
            ActiveCell.Offset(0, -1).Value = "X"
            ActiveCell.Offset(1, 0).Select
    
        Else
    
            ActiveCell.Offset(1, 0).Select
    
    
    End If
Loop


'Preenchimento da terceira coluna:

f.Range("G57").Select


Do Until ActiveCell.Value = 0
    If ActiveCell = Acondici Then
    
            ActiveCell.Offset(0, -1).Value = "X"
            ActiveCell.Offset(1, 0).Select
    
        Else
    
            ActiveCell.Offset(1, 0).Select
    
    
    End If
Loop


'Preenchimento da quarta coluna:

f.Range("I57").Select


Do Until ActiveCell.Value = 0
    If ActiveCell = Acondici Then
    
            ActiveCell.Offset(0, -1).Value = "X"
            ActiveCell.Offset(1, 0).Select
    
        Else
    
            ActiveCell.Offset(1, 0).Select
    
    
    End If
Loop

'Atribuir novo valor de Material para a variável:

ContColunas = ContColunas + 1
Acondici = P.Cells(NFicha + 4, ContColunas).Value


Loop



'Procedimento para preencher "outros":

If f.Range("H57").Value = "X" Then

f.Range("I58") = P.Cells(NFicha + 4, 70).Value


Else

End If


            '--------------------------------------------------------------------------

'Preenchimento 14. Armazenamento

ContColunas = 71

Do While ContColunas < 73

f.Select
f.Range("C61").Select


'Preenchimento da primeira coluna:

Do Until ActiveCell.Value = 0
    If ActiveCell = Armazena Then
    
            ActiveCell.Offset(0, -1).Value = "X"
            ActiveCell.Offset(1, 0).Select
    
        Else
    
            ActiveCell.Offset(1, 0).Select
    
    
    End If
Loop

'Preenchimento da segunda coluna:

f.Range("E61").Select


Do Until ActiveCell.Value = 0
    If ActiveCell = Armazena Then
    
            ActiveCell.Offset(0, -1).Value = "X"
            ActiveCell.Offset(1, 0).Select
    
        Else
    
            ActiveCell.Offset(1, 0).Select
    
    
    End If
Loop


'Preenchimento da terceira coluna:

f.Range("G61").Select


Do Until ActiveCell.Value = 0
    If ActiveCell = Armazena Then
    
            ActiveCell.Offset(0, -1).Value = "X"
            ActiveCell.Offset(1, 0).Select
    
        Else
    
            ActiveCell.Offset(1, 0).Select
    
    
    End If
Loop


'Atribuir novo valor de Armazenamento para a variável:

ContColunas = ContColunas + 1
Armazena = P.Cells(NFicha + 4, ContColunas).Value


Loop

'Procedimento para preencher "outros":

If f.Range("F62").Value = "X" Then

f.Range("H61") = P.Cells(NFicha + 4, 73).Value

Else
End If


            '-----------------------------------------------------------------------

'Preenchimento 28. Tipo de Acervo

'TipoAcervo

If TipoAcervo = "Acervo Geral" Then

f.Range("D77") = "X"

Else

f.Range("F77") = "X"

End If


'------------------------------------------------------------------------------------------------------------------------

   'Copia cada ficha para nova aba:
     
   'Faz a cópia da planilha para nova aba
    f.Copy After:=ActiveWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    
    'Renomeia a nova aba com o número da Ficha
    ActiveSheet.Name = ActiveSheet.Range("D71").Value
    
    'Limpa os dados preenchidos no modelo da ficha
    f.Select
    'Union(Range( _
        "H33:H35,B38:B41,D38:D41,F38:F41,H38:H41,B44,D44,F44,B46:B49,F46,B51:B53,D51:D53,F51:F53,H51,B57:B59,D57:D59,F57:F59,H57,B61:B62,D61:D62,F61:F62,D77,F77,B14,B15,B16,D14,D15,D16,F14,F15,B18" _
        ), Range( _
        "B19,B20,B21,D18,D19,D20,D21,F18:F21,H18:H20,B23:B28,D23:D28,F23:F28,H23:H25,B30,D30,F30,H30,B33,B34,B35,B36,D33:D36,F33:F36" _
        )).Select
        
        
    f.Range("B14:B16,B18:B21,B23:B28,B30,B33:B36,B38:B41,B44,B46:B49,B51:B53,B57:B59,B61:B62").Select
    Selection.ClearContents
    f.Range("D14:D16,D18:D21,D23:D28,D30,D33:D36,D38:D41,D44,D51:D53,D57:D59,D61:D62,D77").Select
    Selection.ClearContents
    f.Range("F14:F15,F18:F21,F23:F28,F30,F33:F36,F38:F41,F44,F46,F51:F53,F57:F59,F61:F62,F77").Select
    Selection.ClearContents
    f.Range("H18:H20,H23:H25,H30,H33:H35,H38:H41,H51,H57,H61:H62,H77").Select
    Selection.ClearContents
    f.Range("G16,G47").Select
    Selection.ClearContents
    f.Range("I21,I26,I31,I36,I42,I52,I58").Select
    Selection.ClearContents
    f.Range("H61").Select
    Selection.ClearContents
        
    
            

Next

f.Protect ("zaza")
f.Visible = False

Application.ScreenUpdating = True
    
End Sub
