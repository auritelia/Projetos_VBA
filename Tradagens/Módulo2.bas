Attribute VB_Name = "M�dulo2"
Sub InserirFoto()

Application.ScreenUpdating = False

Dim Imagem As Variant 'vari�vel que armazena o caminho da pasta com as imagens
Dim Intervencoes As Worksheet 'vari�vel "movel" que armazena a aba que esta em uso
Dim CelulaEmUso As String 'vari�vel m�vel que armazena a c�lula onde a imagem ser� inserida - talvez essa tenha q virar uma string e ser trabalhada com Chr()
Dim Linhas As Long 'Numero de linhas na aba em uso
Dim UltLinha As Range 'armazena a �ltima c�lula em que a ultima foto da tabela ser� inserida
Dim FormatoImagem As String 'armazena os tipos de imagem que poder�o ser inseridas

Linhas = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row 'calcula a qtd de linhas da tabela

FormatoImagem = "JPEG (*.jpeg),*.jpeg, JPG (*.jpg),*.jpg,PNG (*.png),*.jpg, GIF (*.gif),*.gif, BMP (*.bmp),*.bmp" 'formatos suportados
Set UltLinha = Cells(Linhas, 6)
CelulaEmUso = "F2" 'F2 para todas as tabelas
Set Intervencoes = Worksheets("Tradagens_Realizadas")

Imagem = Application.GetOpenFilename(FormatoImagem, False, False, False, True) 'armazena o caminho da pasta com imagens, e pode armazenar mais de um item

'If Imagem = False Then End 'caso a pessoa cancele o "bot�o"

If IsArray(Imagem) Then 'Se varias imagens forem selecionadas

Intervencoes.Activate
j = 2 'contador de linhas da tabela

For i = LBound(Imagem) To UBound(Imagem) 'vai pegar cada imagem, desde a primeira selecionada at� a ultima e fazer as instru��es a seguir


'insere a imagem na CelulaEmUso:
Application.ActiveSheet.Shapes.AddPicture Imagem(i), msoFalse, msoTrue, Range(CelulaEmUso).Left + 4, Range(CelulaEmUso).Top + 4, Range(CelulaEmUso).Width - 8, Range(CelulaEmUso).Height - 8

'Atualiza o endere�o armazenado em CelulaEmUso:
j = j + 1
CelulaEmUso = "F" & j

Next

End If

Application.ScreenUpdating = True

End Sub
