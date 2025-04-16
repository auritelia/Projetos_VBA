Attribute VB_Name = "Módulo4"
Sub LimparTradagensRealizadas()

Application.ScreenUpdating = False

Dim Intervencoes As Worksheet
Set Intervencoes = Worksheets("Tradagens_Realizadas")

Intervencoes.Activate
Intervencoes.Range("A2:F2").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.ClearContents

Intervencoes.Pictures.Delete

Application.ScreenUpdating = True

End Sub
