Attribute VB_Name = "Módulo5"
Sub Limpar_Fotos()

Application.ScreenUpdating = False
Dim P As Worksheet
Dim AbaEmUso As Worksheet 'variável "movel" que armazena a aba que esta em uso
Dim NumAba As Integer 'Indice da aba ativa para colocar as fotos
Dim f As Integer 'contador de fichas
Dim UltFicha As Long

Set P = Worksheets("Preenchimento")
UltFicha = P.Cells(Rows.Count, 1).End(xlUp).Row
NumAba = 6
Set AbaEmUso = Sheets(NumAba)

For f = 1 To UltFicha

AbaEmUso.Pictures.Delete
NumAba = NumAba + 1

If NumAba <= ThisWorkbook.Worksheets.Count Then
Set AbaEmUso = Sheets(NumAba)
Else: End If

Next

Application.ScreenUpdating = True

End Sub
