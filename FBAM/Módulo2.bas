Attribute VB_Name = "Módulo2"
Sub Gerar_Inventario()

Application.ScreenUpdating = False

Dim f As Worksheet
Dim P As Worksheet
Dim S As Worksheet
Dim I As Worksheet
Dim Linhas As Long
Dim FBAM As Workbook
Dim UltLinha As Integer

'Procedimento para gerar o Inventário exatamente com o número de linhas da tabela "preenchimento"

Set I = Sheets("Inventário")
Set S = Sheets("Sitio")
Set P = Sheets("Preenchimento")
Set f = Sheets("FICHA")
Set FBAM = Application.ThisWorkbook

I.Visible = True
I.Unprotect ("zaza")

'Calcula a qtd de linhas da tabela
Linhas = P.Cells(Rows.Count, 1).End(xlUp).Row
Linhas = Linhas - 4


I.Activate
I.Range("A3:M3").Copy
I.Range("A3").Select

Do While ActiveCell.Row <= Linhas + 1
ActiveCell.Offset(1, 0).PasteSpecial (xlPasteAll)

Loop


'copia os números das Fichas de "preenchimento" para "Inventário"

P.Activate
P.Range("A5").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy

'cola Numeros das fichas na aba "inventario"

I.Activate
I.Range("E3").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False



'Faz a cópia da planilha para nova aba
I.Copy After:=ActiveWorkbook.Sheets(3)
ActiveSheet.Name = ActiveSheet.Range("A3")


I.Protect ("zaza")
I.Visible = False

Application.ScreenUpdating = True

End Sub
