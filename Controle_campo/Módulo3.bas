Attribute VB_Name = "Módulo3"
Sub IrParaHoje()

Application.ScreenUpdating = False

Dim T As Worksheet
Dim C As Worksheet
Dim CAC As Workbook
Dim Hoje As Date
Dim DataCalendario As Date

Set T = Sheets("Tabelao")
Set C = Sheets("Calendario")
Set CAC = Application.ThisWorkbook
Hoje = Date

C.Activate
Range("C6").Select
DataCalendario = ActiveCell.Value

While Hoje > DataCalendario

ActiveCell.Offset(0, 1).Select
DataCalendario = ActiveCell.Value

Wend

Application.ScreenUpdating = True

ActiveCell.Select

End Sub
