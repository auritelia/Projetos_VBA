Attribute VB_Name = "Módulo3"
Sub Inserir_Fotos()

Application.ScreenUpdating = False

Dim P As Worksheet
Dim Imagem() As Variant  'variável que armazena o caminho das imagens
Dim MatrizImg() As Variant 'Matriz de 2 colunas, a primeira com os caminhos de cada foto selecionada na pasta, a segunda com o numero da ficha onde a foto deve ser inserida
Dim AbaEmUso As Worksheet 'variável "movel" que armazena a aba que esta em uso
Dim FormatoImagem As String 'armazena os tipos de imagem que poderão ser inseridas
Dim NumAba As Integer 'Indice da aba ativa para colocar as fotos
Dim k As Integer 'contador do array MatrizImg
Dim f As Integer 'contador de fichas
Dim PInicial As Integer ' variavel para isolamento do numero da ficha na string do caminho da foto
Dim PFinal As Integer 'variavel para isolamento do numero da ficha na string do caminho da foto
Dim ContImg As Integer 'contador de qtas imagens por ficha
Dim CampoFotos As String 'variável que armazena o espaço das fotos na ficha "B68:I68"
Dim UltFicha As Long

Set P = Worksheets("Preenchimento")
UltFicha = P.Cells(Rows.Count, 1).End(xlUp).Row
FormatoImagem = "JPG (*.jpg),*.jpg,JPEG (*.jpeg),*.jpeg,PNG (*.png),*.jpg, GIF (*.gif),*.gif, BMP (*.bmp),*.bmp" 'formatos suportados
Imagem = Application.GetOpenFilename(FormatoImagem, False, False, False, True) 'armazena o caminho da pasta com imagens, e pode armazenar mais de um item

'-------------------------------------------------------//----------------------------------------------------------------------------

If IsArray(Imagem) Then 'Se varias imagens forem selecionadas

ReDim MatrizImg(1 To UBound(Imagem) + 1, 1 To UBound(Imagem) + 1) As Variant 'cria um array 2d com a segunda coluna para especificar em qual ficha vai a foto da primeira coluna

For k = LBound(Imagem) To UBound(Imagem)  'transferindo os dados do array(Imagem) para MatrizImg
    MatrizImg(k, 1) = Imagem(k)

    'isola o numero do lote no nome do arquivo e armazena ele na coluna 2 da matriz MatrizImg
    PInicial = Len(MatrizImg(k, 1))

    While PInicial >= 1
    If Mid(MatrizImg(k, 1), PInicial, 4) = "Lote" Then
    MatrizImg(k, 2) = Mid(MatrizImg(k, 1), PInicial + 4, 10)
        PFinal = 1
        While PFinal <= Len(MatrizImg(k, 2))
        If Mid(MatrizImg(k, 2), PFinal, 1) = "(" Then
        MatrizImg(k, 2) = Mid(MatrizImg(k, 2), 1, PFinal - 1)
        End If
        PFinal = PFinal + 1
        Wend
    End If
    PInicial = PInicial - 1
    Wend
    Trim (MatrizImg(k, 2))
    Next

Else: End If
'--------------------------------------MatrizImg pronta!

'-------------------------------------conferencia MatrizImg:
'ThisWorkbook.Worksheets("Planilha1").Activate
'ActiveSheet.Range("A1").Select
'For k = 1 To UBound(MatrizImg)
'ActiveCell.Value = MatrizImg(k, 2)
'ActiveCell.Offset(0, 1).Value = MatrizImg(k, 1)
'ActiveCell.Offset(1, 0).Select
'Next
'----------------------------------------fim da conferencia

NumAba = 6
Set AbaEmUso = Sheets(NumAba)
CampoFotos = "B68:I68"
k = 1

For f = 1 To UltFicha

'procedimento para contar quantas fotos serão inseridas em cada ficha:
AbaEmUso.Activate
ContImg = 1


While MatrizImg(k, 2) = MatrizImg(k + 1, 2)
ContImg = ContImg + 1
k = k + 1
Wend

'inserção e configuração das fotos no espaço da ficha:

If ContImg > 18 Then
MsgBox ("Esta planilha só suporta 18 imagens por ficha, por favor verifique a quantidade de fotos para cada ficha")
Exit Sub
Else: End If

If ContImg = 1 Then
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k, 1), msoFalse, msoTrue, Range(CampoFotos).Left + 180, Range(CampoFotos).Top + 4, Range(CampoFotos).Width - 350, Range(CampoFotos).Height - 8
Else: End If

If ContImg > 1 And ContImg < 3 Then
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - 1, 1), msoFalse, msoTrue, Range(CampoFotos).Left + 8, Range(CampoFotos).Top + 4, Range(CampoFotos).Width / 2.1, Range(CampoFotos).Height - 8 'imagem da esquerda
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k, 1), msoFalse, msoTrue, Range(CampoFotos).Left + 380, Range(CampoFotos).Top + 4, Range(CampoFotos).Width / 2.1, Range(CampoFotos).Height - 8 'imagem da direita
Else: End If

If ContImg = 3 Then
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 1), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 4, Range(CampoFotos).Top + 50, Range(CampoFotos).Width / 3.05, Range(CampoFotos).Height / 1.5 'imagem da esquerda
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 2), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 248, Range(CampoFotos).Top + 50, Range(CampoFotos).Width / 3.05, Range(CampoFotos).Height / 1.5 'imagem do centro
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k, 1), msoFalse, msoTrue, Range(CampoFotos).Left + 492, Range(CampoFotos).Top + 50, Range(CampoFotos).Width / 3.05, Range(CampoFotos).Height / 1.5 'imagem da direita
Else: End If

If ContImg = 4 Then
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 1), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 3, Range(CampoFotos).Top + 70, Range(CampoFotos).Width / 4.05, Range(CampoFotos).Height / 2
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 2), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 186, Range(CampoFotos).Top + 70, Range(CampoFotos).Width / 4.05, Range(CampoFotos).Height / 2
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 3), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 369, Range(CampoFotos).Top + 70, Range(CampoFotos).Width / 4.05, Range(CampoFotos).Height / 2
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k, 1), msoFalse, msoTrue, Range(CampoFotos).Left + 552, Range(CampoFotos).Top + 70, Range(CampoFotos).Width / 4.05, Range(CampoFotos).Height / 2
Else: End If

If ContImg = 5 Then
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 1), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 3, Range(CampoFotos).Top + 3, Range(CampoFotos).Width / 4.05, Range(CampoFotos).Height / 2.05
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 2), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 186, Range(CampoFotos).Top + 3, Range(CampoFotos).Width / 4.05, Range(CampoFotos).Height / 2.05
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 3), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 369, Range(CampoFotos).Top + 3, Range(CampoFotos).Width / 4.05, Range(CampoFotos).Height / 2.05
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 4), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 552, Range(CampoFotos).Top + 3, Range(CampoFotos).Width / 4.05, Range(CampoFotos).Height / 2.05
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k, 1), msoFalse, msoTrue, Range(CampoFotos).Left + 3, Range(CampoFotos).Top + 143, Range(CampoFotos).Width / 4.05, Range(CampoFotos).Height / 2.05
Else: End If

If ContImg = 6 Then
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 1), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 3, Range(CampoFotos).Top + 3, Range(CampoFotos).Width / 4.05, Range(CampoFotos).Height / 2.05
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 2), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 186, Range(CampoFotos).Top + 3, Range(CampoFotos).Width / 4.05, Range(CampoFotos).Height / 2.05
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 3), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 369, Range(CampoFotos).Top + 3, Range(CampoFotos).Width / 4.05, Range(CampoFotos).Height / 2.05
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 4), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 552, Range(CampoFotos).Top + 3, Range(CampoFotos).Width / 4.05, Range(CampoFotos).Height / 2.05
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 5), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 3, Range(CampoFotos).Top + 143, Range(CampoFotos).Width / 4.05, Range(CampoFotos).Height / 2.05
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k, 1), msoFalse, msoTrue, Range(CampoFotos).Left + 186, Range(CampoFotos).Top + 143, Range(CampoFotos).Width / 4.05, Range(CampoFotos).Height / 2.05
Else: End If

If ContImg = 7 Then
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 1), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 3, Range(CampoFotos).Top + 3, Range(CampoFotos).Width / 4.05, Range(CampoFotos).Height / 2.05
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 2), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 186, Range(CampoFotos).Top + 3, Range(CampoFotos).Width / 4.05, Range(CampoFotos).Height / 2.05
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 3), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 369, Range(CampoFotos).Top + 3, Range(CampoFotos).Width / 4.05, Range(CampoFotos).Height / 2.05
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 4), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 552, Range(CampoFotos).Top + 3, Range(CampoFotos).Width / 4.05, Range(CampoFotos).Height / 2.05
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 5), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 3, Range(CampoFotos).Top + 143, Range(CampoFotos).Width / 4.05, Range(CampoFotos).Height / 2.05
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 6), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 186, Range(CampoFotos).Top + 143, Range(CampoFotos).Width / 4.05, Range(CampoFotos).Height / 2.05
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k, 1), msoFalse, msoTrue, Range(CampoFotos).Left + 369, Range(CampoFotos).Top + 143, Range(CampoFotos).Width / 4.05, Range(CampoFotos).Height / 2.05
Else: End If

If ContImg = 8 Then
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 1), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 3, Range(CampoFotos).Top + 3, Range(CampoFotos).Width / 4.05, Range(CampoFotos).Height / 2.05
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 2), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 186, Range(CampoFotos).Top + 3, Range(CampoFotos).Width / 4.05, Range(CampoFotos).Height / 2.05
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 3), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 369, Range(CampoFotos).Top + 3, Range(CampoFotos).Width / 4.05, Range(CampoFotos).Height / 2.05
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 4), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 552, Range(CampoFotos).Top + 3, Range(CampoFotos).Width / 4.05, Range(CampoFotos).Height / 2.05
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 5), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 3, Range(CampoFotos).Top + 143, Range(CampoFotos).Width / 4.05, Range(CampoFotos).Height / 2.05
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 6), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 186, Range(CampoFotos).Top + 143, Range(CampoFotos).Width / 4.05, Range(CampoFotos).Height / 2.05
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 7), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 369, Range(CampoFotos).Top + 143, Range(CampoFotos).Width / 4.05, Range(CampoFotos).Height / 2.05
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k, 1), msoFalse, msoTrue, Range(CampoFotos).Left + 552, Range(CampoFotos).Top + 143, Range(CampoFotos).Width / 4.05, Range(CampoFotos).Height / 2.05
Else: End If

If ContImg = 9 Then
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 1), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 1.5, Range(CampoFotos).Top + 20, Range(CampoFotos).Width / 5.05, Range(CampoFotos).Height / 2.5
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 2), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 148, Range(CampoFotos).Top + 20, Range(CampoFotos).Width / 5.05, Range(CampoFotos).Height / 2.5
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 3), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 295, Range(CampoFotos).Top + 20, Range(CampoFotos).Width / 5.05, Range(CampoFotos).Height / 2.5
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 4), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 442, Range(CampoFotos).Top + 20, Range(CampoFotos).Width / 5.05, Range(CampoFotos).Height / 2.5
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 5), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 589, Range(CampoFotos).Top + 20, Range(CampoFotos).Width / 5.05, Range(CampoFotos).Height / 2.5
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 6), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 1.5, Range(CampoFotos).Top + 155, Range(CampoFotos).Width / 5.05, Range(CampoFotos).Height / 2.5
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 7), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 148, Range(CampoFotos).Top + 155, Range(CampoFotos).Width / 5.05, Range(CampoFotos).Height / 2.5
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 8), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 295, Range(CampoFotos).Top + 155, Range(CampoFotos).Width / 5.05, Range(CampoFotos).Height / 2.5
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k, 1), msoFalse, msoTrue, Range(CampoFotos).Left + 442, Range(CampoFotos).Top + 155, Range(CampoFotos).Width / 5.05, Range(CampoFotos).Height / 2.5
Else: End If

If ContImg = 10 Then
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 1), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 1.5, Range(CampoFotos).Top + 20, Range(CampoFotos).Width / 5.05, Range(CampoFotos).Height / 2.5
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 2), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 148, Range(CampoFotos).Top + 20, Range(CampoFotos).Width / 5.05, Range(CampoFotos).Height / 2.5
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 3), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 295, Range(CampoFotos).Top + 20, Range(CampoFotos).Width / 5.05, Range(CampoFotos).Height / 2.5
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 4), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 442, Range(CampoFotos).Top + 20, Range(CampoFotos).Width / 5.05, Range(CampoFotos).Height / 2.5
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 5), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 589, Range(CampoFotos).Top + 20, Range(CampoFotos).Width / 5.05, Range(CampoFotos).Height / 2.5
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 6), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 1.5, Range(CampoFotos).Top + 155, Range(CampoFotos).Width / 5.05, Range(CampoFotos).Height / 2.5
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 7), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 148, Range(CampoFotos).Top + 155, Range(CampoFotos).Width / 5.05, Range(CampoFotos).Height / 2.5
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 8), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 295, Range(CampoFotos).Top + 155, Range(CampoFotos).Width / 5.05, Range(CampoFotos).Height / 2.5
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 9), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 442, Range(CampoFotos).Top + 155, Range(CampoFotos).Width / 5.05, Range(CampoFotos).Height / 2.5
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k, 1), msoFalse, msoTrue, Range(CampoFotos).Left + 589, Range(CampoFotos).Top + 155, Range(CampoFotos).Width / 5.05, Range(CampoFotos).Height / 2.5
Else: End If

If ContImg = 11 Then
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 1), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 2, Range(CampoFotos).Top + 40, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.05
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 2), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 124, Range(CampoFotos).Top + 40, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.05
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 3), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 246, Range(CampoFotos).Top + 40, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.05
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 4), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 368, Range(CampoFotos).Top + 40, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.05
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 5), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 490, Range(CampoFotos).Top + 40, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.05
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 6), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 613, Range(CampoFotos).Top + 40, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.05
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 7), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 2, Range(CampoFotos).Top + 155, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.05
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 8), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 124, Range(CampoFotos).Top + 155, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.05
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 9), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 246, Range(CampoFotos).Top + 155, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.05
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 10), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 368, Range(CampoFotos).Top + 155, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.05
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k, 1), msoFalse, msoTrue, Range(CampoFotos).Left + 490, Range(CampoFotos).Top + 155, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.05
Else: End If

If ContImg = 12 Then
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 1), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 2, Range(CampoFotos).Top + 40, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.05
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 2), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 124, Range(CampoFotos).Top + 40, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.05
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 3), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 246, Range(CampoFotos).Top + 40, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.05
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 4), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 368, Range(CampoFotos).Top + 40, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.05
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 5), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 490, Range(CampoFotos).Top + 40, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.05
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 6), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 613, Range(CampoFotos).Top + 40, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.05
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 7), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 2, Range(CampoFotos).Top + 155, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.05
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 8), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 124, Range(CampoFotos).Top + 155, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.05
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 9), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 246, Range(CampoFotos).Top + 155, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.05
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 10), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 368, Range(CampoFotos).Top + 155, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.05
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 11), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 490, Range(CampoFotos).Top + 155, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.05
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k, 1), msoFalse, msoTrue, Range(CampoFotos).Left + 613, Range(CampoFotos).Top + 155, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.05
Else: End If

If ContImg = 13 Then
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 1), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 2, Range(CampoFotos).Top + 2, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.07
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 2), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 124, Range(CampoFotos).Top + 2, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.07
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 3), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 246, Range(CampoFotos).Top + 2, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.07
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 4), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 368, Range(CampoFotos).Top + 2, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.07
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 5), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 490, Range(CampoFotos).Top + 2, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.07
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 6), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 613, Range(CampoFotos).Top + 2, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.07
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 7), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 2, Range(CampoFotos).Top + 95.5, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.07
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 8), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 124, Range(CampoFotos).Top + 95.5, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.07
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 9), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 246, Range(CampoFotos).Top + 95.5, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.07
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 10), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 368, Range(CampoFotos).Top + 95.5, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.07
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 11), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 490, Range(CampoFotos).Top + 95.5, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.07
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 12), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 613, Range(CampoFotos).Top + 95.5, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.07
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k, 1), msoFalse, msoTrue, Range(CampoFotos).Left + 2, Range(CampoFotos).Top + 189, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.07
Else: End If

If ContImg = 14 Then
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 1), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 2, Range(CampoFotos).Top + 2, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.07
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 2), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 124, Range(CampoFotos).Top + 2, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.07
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 3), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 246, Range(CampoFotos).Top + 2, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.07
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 4), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 368, Range(CampoFotos).Top + 2, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.07
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 5), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 490, Range(CampoFotos).Top + 2, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.07
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 6), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 613, Range(CampoFotos).Top + 2, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.07
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 7), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 2, Range(CampoFotos).Top + 95.5, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.07
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 8), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 124, Range(CampoFotos).Top + 95.5, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.07
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 9), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 246, Range(CampoFotos).Top + 95.5, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.07
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 10), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 368, Range(CampoFotos).Top + 95.5, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.07
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 11), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 490, Range(CampoFotos).Top + 95.5, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.07
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 12), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 613, Range(CampoFotos).Top + 95.5, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.07
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 13), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 2, Range(CampoFotos).Top + 189, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.07
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k, 1), msoFalse, msoTrue, Range(CampoFotos).Left + 124, Range(CampoFotos).Top + 189, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.07
Else: End If

If ContImg = 15 Then
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 1), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 2, Range(CampoFotos).Top + 2, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.07
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 2), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 124, Range(CampoFotos).Top + 2, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.07
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 3), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 246, Range(CampoFotos).Top + 2, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.07
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 4), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 368, Range(CampoFotos).Top + 2, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.07
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 5), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 490, Range(CampoFotos).Top + 2, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.07
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 6), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 613, Range(CampoFotos).Top + 2, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.07
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 7), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 2, Range(CampoFotos).Top + 95.5, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.07
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 8), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 124, Range(CampoFotos).Top + 95.5, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.07
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 9), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 246, Range(CampoFotos).Top + 95.5, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.07
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 10), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 368, Range(CampoFotos).Top + 95.5, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.07
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 11), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 490, Range(CampoFotos).Top + 95.5, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.07
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 12), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 613, Range(CampoFotos).Top + 95.5, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.07
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 13), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 2, Range(CampoFotos).Top + 189, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.07
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 14), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 124, Range(CampoFotos).Top + 189, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.07
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k, 1), msoFalse, msoTrue, Range(CampoFotos).Left + 246, Range(CampoFotos).Top + 189, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.07
Else: End If

If ContImg = 16 Then
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 1), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 2, Range(CampoFotos).Top + 2, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.07
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 2), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 124, Range(CampoFotos).Top + 2, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.07
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 3), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 246, Range(CampoFotos).Top + 2, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.07
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 4), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 368, Range(CampoFotos).Top + 2, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.07
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 5), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 490, Range(CampoFotos).Top + 2, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.07
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 6), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 613, Range(CampoFotos).Top + 2, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.07
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 7), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 2, Range(CampoFotos).Top + 95.5, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.07
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 8), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 124, Range(CampoFotos).Top + 95.5, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.07
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 9), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 246, Range(CampoFotos).Top + 95.5, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.07
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 10), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 368, Range(CampoFotos).Top + 95.5, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.07
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 11), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 490, Range(CampoFotos).Top + 95.5, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.07
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 12), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 613, Range(CampoFotos).Top + 95.5, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.07
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 13), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 2, Range(CampoFotos).Top + 189, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.07
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 14), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 124, Range(CampoFotos).Top + 189, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.07
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 15), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 246, Range(CampoFotos).Top + 189, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.07
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k, 1), msoFalse, msoTrue, Range(CampoFotos).Left + 368, Range(CampoFotos).Top + 189, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.07
Else: End If

If ContImg = 17 Then
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 1), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 2, Range(CampoFotos).Top + 2, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.07
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 2), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 124, Range(CampoFotos).Top + 2, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.07
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 3), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 246, Range(CampoFotos).Top + 2, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.07
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 4), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 368, Range(CampoFotos).Top + 2, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.07
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 5), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 490, Range(CampoFotos).Top + 2, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.07
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 6), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 613, Range(CampoFotos).Top + 2, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.07
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 7), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 2, Range(CampoFotos).Top + 95.5, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.07
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 8), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 124, Range(CampoFotos).Top + 95.5, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.07
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 9), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 246, Range(CampoFotos).Top + 95.5, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.07
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 10), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 368, Range(CampoFotos).Top + 95.5, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.07
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 11), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 490, Range(CampoFotos).Top + 95.5, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.07
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 12), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 613, Range(CampoFotos).Top + 95.5, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.07
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 13), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 2, Range(CampoFotos).Top + 189, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.07
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 14), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 124, Range(CampoFotos).Top + 189, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.07
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 15), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 246, Range(CampoFotos).Top + 189, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.07
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 16), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 368, Range(CampoFotos).Top + 189, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.07
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k, 1), msoFalse, msoTrue, Range(CampoFotos).Left + 490, Range(CampoFotos).Top + 189, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.07
Else: End If

If ContImg = 18 Then
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 1), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 2, Range(CampoFotos).Top + 2, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.07
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 2), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 124, Range(CampoFotos).Top + 2, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.07
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 3), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 246, Range(CampoFotos).Top + 2, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.07
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 4), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 368, Range(CampoFotos).Top + 2, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.07
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 5), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 490, Range(CampoFotos).Top + 2, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.07
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 6), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 613, Range(CampoFotos).Top + 2, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.07
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 7), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 2, Range(CampoFotos).Top + 95.5, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.07
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 8), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 124, Range(CampoFotos).Top + 95.5, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.07
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 9), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 246, Range(CampoFotos).Top + 95.5, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.07
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 10), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 368, Range(CampoFotos).Top + 95.5, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.07
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 11), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 490, Range(CampoFotos).Top + 95.5, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.07
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 12), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 613, Range(CampoFotos).Top + 95.5, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.07
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 13), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 2, Range(CampoFotos).Top + 189, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.07
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 14), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 124, Range(CampoFotos).Top + 189, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.07
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 15), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 246, Range(CampoFotos).Top + 189, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.07
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 16), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 368, Range(CampoFotos).Top + 189, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.07
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k - (ContImg - 17), 1), msoFalse, msoTrue, Range(CampoFotos).Left + 490, Range(CampoFotos).Top + 189, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.07
Application.ActiveSheet.Shapes.AddPicture MatrizImg(k, 1), msoFalse, msoTrue, Range(CampoFotos).Left + 613, Range(CampoFotos).Top + 189, Range(CampoFotos).Width / 6.07, Range(CampoFotos).Height / 3.07

Else: End If

'atualização da aba:
k = k + 1

NumAba = NumAba + 1
If NumAba <= ThisWorkbook.Worksheets.Count Then
Set AbaEmUso = Sheets(NumAba)
Else
Application.ScreenUpdating = True
Exit Sub
End If

Next

Application.ScreenUpdating = True

End Sub
