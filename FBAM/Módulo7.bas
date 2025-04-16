Attribute VB_Name = "Módulo7"
Sub Preencher()
Dim Etq As Worksheet
Dim ContLinhas As Integer
Dim E As Range
Set Etq = Worksheets("ETIQUETA")

ContLinhas = 7

Etq.Activate
Etq.Range("B7").Select
Etq.Range("B7").Copy

While ContLinhas <= 420
ActiveCell.Offset(7, 0).Select
ActiveCell.PasteSpecial (xlPasteFormulas)
ContLinhas = ContLinhas + 7
Wend

'Coluna 2

ContLinhas = 7
Etq.Range("H7").Select
Etq.Range("H7").Copy

While ContLinhas <= 420
ActiveCell.Offset(7, 0).Select
ActiveCell.PasteSpecial (xlPasteFormulas)
ContLinhas = ContLinhas + 7
Wend

'Coluna 3

ContLinhas = 7
Etq.Range("N7").Select
Etq.Range("N7").Copy

While ContLinhas <= 420
ActiveCell.Offset(7, 0).Select
ActiveCell.PasteSpecial (xlPasteFormulas)
ContLinhas = ContLinhas + 7
Wend

'Coluna 4

ContLinhas = 7
Etq.Range("T7").Select
Etq.Range("T7").Copy

While ContLinhas <= 420
ActiveCell.Offset(7, 0).Select
ActiveCell.PasteSpecial (xlPasteFormulas)
ContLinhas = ContLinhas + 7
Wend


End Sub
