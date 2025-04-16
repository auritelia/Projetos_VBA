Attribute VB_Name = "Módulo3"
Sub PreencherEstratigrafia()

Application.ScreenUpdating = False

Dim FichasTrad As Worksheet
Dim Intervencoes As Worksheet
Dim Estratigrafia As String 'variavel que irá armazenar o texto da estratigrafia, concatenar tudo
Dim StatusInterv As String ' variável que armazena o status da intervenção
Dim Coordenadas As String   'variavel que armazena as coordenadas da intervenção
Dim ProfFinal As String 'variável que armazena a profundidade final da tradagem
Dim CodTrad As String 'variável que armazena o codigo da tradagem
Dim Interrupcao As String
Dim i As Integer 'contador de linhas
Dim j As Integer 'contador de colunas
Dim k As Integer 'contador do array TextoNivel
Dim UltLinha As Long
Dim A(1 To 4) As Variant ' A,B,C e D são os arrays com a tabela de solos
Dim B(1 To 4) As Variant
Dim C(1 To 4) As Variant
Dim D(1 To 9) As Variant
Dim N As String, NN As String, P1 As String, P2 As String, P3 As String, P4 As String 'N e NN são as variáveis que vão conferir qtos niveis diferentes de solo tem em cada tradagem
Dim Nivel(21, 21) As Variant 'array que irá armazenar os niveis com diferentes tipos de solo que aparecem em cada tradagem


Set FichasTrad = Worksheets("Fichas_Tradagem")
Set Intervencoes = Worksheets("Tradagens_Realizadas")
CodTrad = FichasTrad.Range("A4").Value 'ajustar com coordenadas
Coordenadas = FichasTrad.Range("C4").Value 'ajustar com coordenadas
UltLinha = FichasTrad.Cells(Rows.Count, 1).End(xlUp).Row 'calcula numero de linhas na tabela
k = 0
j = 4
i = 4

A(1) = "arenosa"
A(2) = "areno argilosa"
A(3) = "argilo arenosa"
A(4) = "argilosa"
B(1) = "sem compactação"
B(2) = "pouco compacto"
B(3) = "compacto"
B(4) = "muito compacto"
C(1) = "pouco friável"
C(2) = "friável"
C(3) = "muito friável"
C(4) = "solo solto"
D(1) = "marrom-clara"
D(2) = "marrom-média"
D(3) = "marrom-escura"
D(4) = "cinza-clara"
D(5) = "cinza-escura"
D(6) = "marrom-amarelada"
D(7) = "marrom-alaranjada"
D(8) = "marrom-avermelhada"

'---------------------------------------------------------------------------------------------
 Nivel(k, 2) = "0 - "

For i = 4 To UltLinha 'percorre as linhas

    'Armazenando os números que indicam tipo de solo no array TextoNivel:
    FichasTrad.Activate
    N = FichasTrad.Cells(i, j).Value
    NN = N
    FichasTrad.Cells(i, j).Select
   Nivel(k, 2) = "0 - "
    
    While N = NN And j <= 24 And ActiveCell.Offset.Value <> 0 And k <= 20
    
    While N = NN
    
    ActiveCell.Offset(0, 1).Select
    NN = ActiveCell.Value
    j = j + 1
    Wend
    
    Nivel(k, 2) = Nivel(k, 2) & FichasTrad.Cells(3, j - 1).Value & "0 cm: "
        If Left(Nivel(k, 2), 2) = "00" Then
        Nivel(k, 2) = Right(Nivel(k, 2), 500)
        Else: End If
        
    Nivel(k, 1) = N
    k = k + 1
    N = NN
    
   If FichasTrad.Cells(3, j - 1).Value = "0" Then
    Nivel(k, 2) = "0 - "
    Else
    Nivel(k, 2) = FichasTrad.Cells(3, j - 1).Value & "0 - "
    End If
    
    Wend
    
    'se o tipo de solo da superfície for único, troca o texto de Nivel(0,2) por "superfície":
    If Nivel(0, 2) = "0 - 00 cm: " Then
    Nivel(0, 2) = "Superfície: "
    Else: End If
    
    '----------------------------//----------------------------------------//----------------------------
    
    'conferencia se Nivel está armazenando direito as infos:
    'k = 0
    'Worksheets("Planilha1").Activate
    'Worksheets("Planilha1").Range("A1").Select
    'For k = 0 To 20
    'ActiveCell.Value = Nivel(k, 1)
    'ActiveCell.Offset(0, 1).Select
    'ActiveCell.Value = Nivel(k, 2)
    'ActiveCell.Offset(1, -1).Select
    'Next
    
    '------------------------------//-------------------------------//-----------------------------------------
    
    
    'atribuição dos 4 valores A B C e D contidos em um nivel da tradagem:
    
    k = 0
    Estratigrafia = ""
    
    
    Do Until Nivel(k, 1) = "" Or k = 21
    
    P1 = Left(Nivel(k, 1), 1)
    P2 = Mid(Nivel(k, 1), 2, 1)
    P3 = Mid(Nivel(k, 1), 3, 1)
    P4 = Right(Nivel(k, 1), 1)
    
    'Atribuição do texto contido nos arrays às variáveis "P"
    P1 = A(P1)
    P2 = B(P2)
    P3 = C(P3)
    P4 = D(P4)
    
    Estratigrafia = Estratigrafia & Nivel(k, 2) & "sedimento com textura " & P1 & ", " & P2 & ", " & P3 & " e de coloração " & P4 & ". "
    
    k = k + 1
   
    
    Loop


'armazenar os dados na tabela de tradagens:
ProfFinal = Right(Nivel(k - 1, 2), 8)
ProfFinal = Left(ProfFinal, 6)
If ProfFinal = "rfície" Then
ProfFinal = "Superfície"
Else: End If
CodTrad = FichasTrad.Cells(i, 1).Value
StatusInterv = FichasTrad.Cells(i, 2).Value
Coordenadas = FichasTrad.Cells(i, 3).Value
Interrupcao = FichasTrad.Cells(i, 25).Value
Estratigrafia = Estratigrafia & " Tradagem interrompida ao alcançar " & Interrupcao & "."

Intervencoes.Activate
Intervencoes.Cells(i - 2, 1).Value = CodTrad
Intervencoes.Cells(i - 2, 2).Value = Coordenadas
Intervencoes.Cells(i - 2, 3).Value = StatusInterv
Intervencoes.Cells(i - 2, 4).Value = Estratigrafia
Intervencoes.Cells(i - 2, 5).Value = ProfFinal

j = 4
k = 0
For k = 0 To 21
Nivel(k, 1) = Empty
Nivel(k, 2) = Empty
Next

k = 0

Next

Application.ScreenUpdating = True

End Sub
