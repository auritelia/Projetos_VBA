Attribute VB_Name = "Módulo6"
Dim Lote As String
Dim P As Worksheet
Dim S As Worksheet
Dim E As Worksheet
Dim FBAM As Workbook
Dim Linhas As Integer
Dim NFicha As Integer
Dim QtAbas As Integer
Dim AbaEtq As Integer
Dim ContColunas As Integer
Dim ContLinhas As Integer
Dim ContC As Integer
Dim ContL As Integer

Sub Gerar_Etiquetas()

Application.ScreenUpdating = False

Set P = Worksheets("Preenchimento")
Set S = Worksheets("Sitio")
Set E = Worksheets("ETIQUETA")
Set FBAM = Application.ThisWorkbook

'identificando a quantidade de linhas em "Preenchimento":
Linhas = P.Cells(Rows.Count, 1).End(xlUp).Row
Linhas = Linhas - 4

'trazendo a aba das etiquetas à tona:
E.Visible = True
E.Unprotect ("zaza")

'Descobrindo quantas abas de etiquetas vai precisar gerar:
QtdAbas = 1
AbaEtq = 1
QtdAbas = Linhas / 240
QtdAbas = Application.WorksheetFunction.RoundUp(QtdAbas, 0)
NFicha = 1

For AbaEtq = 1 To QtdAbas

ContColunas = 2
ContLinhas = 5

While ContLinhas <= 418

While ContColunas < 21

Lote = P.Cells(NFicha + 4, 1).Value

'Preenchendo etiquetas:

E.Activate
E.Cells(ContLinhas, ContColunas).Select
ActiveCell.Value = Lote
ContColunas = ContColunas + 6
NFicha = NFicha + 1

Wend

ContLinhas = ContLinhas + 7
ContColunas = 2

Wend

'Duplicando Aba e renomeando para poder preencher de novo:
E.Activate
Sheets("ETIQUETA").Copy After:=Sheets(ThisWorkbook.Worksheets.Count)
ActiveSheet.Name = "Etq " & AbaEtq

'Limpando células do numero do lote para dar sequencia ao preenchimento da proxima aba

E.Activate

'Coluna 1
E.Range("B5").Select
ContL = 5

While ContL < 418
ActiveCell.Value = ""
ActiveCell.Offset(7, 0).Select
ContL = ContL + 7
Wend
ActiveCell.Value = ""

'Coluna 2
E.Range("H5").Select
ContL = 5

While ContL < 418
ActiveCell.Value = ""
ActiveCell.Offset(7, 0).Select
ContL = ContL + 7
Wend
ActiveCell.Value = ""

'Coluna 3
E.Range("N5").Select
ContL = 5

While ContL < 418
ActiveCell.Value = ""
ActiveCell.Offset(7, 0).Select
ContL = ContL + 7
Wend
ActiveCell.Value = ""

'Coluna 4
E.Range("T5").Select
ContL = 5

While ContL < 418
ActiveCell.Value = ""
ActiveCell.Offset(7, 0).Select
ContL = ContL + 7
Wend
ActiveCell.Value = ""

Next

'protegendo de volta a aba das etiquetas:
E.Protect ("zaza")
E.Visible = False

Application.ScreenUpdating = True

'TA BASICAMENTO PRONTO, AGORA PRECISA FAZER OS REFINAMENTOS PARA USO!
End Sub
