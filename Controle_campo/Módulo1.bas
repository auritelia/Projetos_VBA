Attribute VB_Name = "Módulo1"
Sub Atualizar_Calendario()

'Application.ScreenUpdating = False

Dim T As Worksheet
Dim C As Worksheet
Dim CAC As Workbook

Dim Nome As String
Dim CodProj As String
Dim MobPrev As Date
Dim MobReal As Date
Dim DesmobPrev As Date
Dim DesmobReal As Date

Dim Data As Date
Dim Mob As Date
Dim Desmob As Date

Dim Linhas As Integer
Dim NColuna As Integer
Dim Entrada As Integer
Dim Arqueologos As Range


Set T = Sheets("Tabelao")
Set C = Sheets("Calendario")
Set CAC = Application.ThisWorkbook

Linhas = T.Range("A1048576").End(xlUp).Row
Linhas = Linhas - 2
Entrada = 1

C.Range("C9:NC200").ClearContents
C.Range("C9:NC200").Interior.ColorIndex = xlColorIndexNone
C.Range("C9:NC200").Font.Color = RGB(0, 0, 0)


For Entrada = 1 To Linhas

T.Activate
Nome = Cells(Entrada + 2, 4).Value
CodProj = Cells(Entrada + 2, 6).Value
MobPrev = Cells(Entrada + 2, 9).Value
MobReal = Cells(Entrada + 2, 10).Value
DesmobPrev = Cells(Entrada + 2, 11).Value
DesmobReal = Cells(Entrada + 2, 12).Value

If DesmobReal <> 0 Then

Desmob = DesmobReal

Else

Desmob = DesmobPrev

End If


If MobReal <> 0 Then

Mob = MobReal

Else

Mob = MobPrev

End If


C.Activate
Data = C.Range("C6").Value
C.Range("B9:B200").Find(Nome).Select
ActiveCell.Offset(0, 1).Select

While Data <= Desmob

If Data >= Mob Then

ActiveCell.Value = CodProj

    If DesmobReal <> 0 Then

    ActiveCell.Interior.Color = RGB(153, 255, 153)
    ActiveCell.Font.Color = RGB(0, 84, 0)

    End If

    If MobReal <> 0 And DesmobReal = 0 Then

    ActiveCell.Interior.Color = RGB(254, 195, 88)
    ActiveCell.Font.Color = RGB(69, 50, 1)

    End If
    
    If MobReal = 0 And DesmobReal = 0 Then

    ActiveCell.Interior.Color = RGB(255, 204, 204)
    ActiveCell.Font.Color = RGB(150, 54, 52)

    End If

End If

Data = Data + 1

ActiveCell.Offset(0, 1).Select

Wend

Next


'Application.ScreenUpdating = True

End Sub
