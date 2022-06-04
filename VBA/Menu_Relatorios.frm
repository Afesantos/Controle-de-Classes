VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Menu_Relatorios 
   Caption         =   "Impressão de Relatórios"
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8970
   OleObjectBlob   =   "Menu_Relatorios.frx":0000
   StartUpPosition =   1  'CenterOwner
   Tag             =   "0"
End
Attribute VB_Name = "Menu_Relatorios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim Planilha As String

Private Sub CmdFichCad_Click()
Dim strPDF_File_Name As String
strPDF_File_Name = ActiveWorkbook.Path & "\Cadastro de Alunos EBD.pdf"
'strPDF_File_Name = Application.GetOpenFilename (". . Arquivos PDF *, pdf , todos os arquivos * * " , 1, " Open File ", False)
If Dir(strPDF_File_Name) <> vbNullString Then ActiveWorkbook.FollowHyperlink strPDF_File_Name Else MsgBox ("Arquivo: <" & strPDF_File_Name & "> -  não foi encontrado.")
End Sub

Private Sub ComboBox1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
TextBox1.Text = "Escolha uma Classe."
End Sub



Private Sub CommandButton1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
TextBox1.Text = "Selecione uma Classe e imprima os participantes."
End Sub

Private Sub CommandButton2_Click()
Planilha = "Presença_" & ComboBox1.Text & "-" & Month(Me.MonthView1.Value) & "-" & Year(Me.MonthView1.Value)

If ComboBox1.Text <> "Escolha uma Classe" Then
    If Not PlanExiste(Planilha) Then
        Resp = MsgBox("Planilha não existe, deseja cria-la agora?", vbOKCancel, "Gerando Lista de Presença")
        If Resp = vbOK Then
            pNovaClasse = ComboBox1.Text
            CriaListaPres
        Else
        Exit Sub
        End If
    End If
Application.ScreenUpdating = True
Sheets(Planilha).Select
Calcula_Domingos
With ActiveSheet   'Preenche Domingos no Calendário
    .Range("C1") = "Classe: " & Me.ComboBox1.Text
    .Range("E1") = MonthView1.Value
    .Range("E2") = CalenDomingos(1) & "/" & MonthView1.Month
    .Range("F2") = CalenDomingos(2) & "/" & MonthView1.Month
    .Range("G2") = CalenDomingos(3) & "/" & MonthView1.Month
    .Range("H2") = CalenDomingos(4) & "/" & MonthView1.Month
    If CalenDomingos(5) <> "" Then
        .Range("I2") = CalenDomingos(5) & "/" & MonthView1.Month
    End If
    If CalenDomingos(5) <> NullString Then
        .Range("J2") = CalenDomingos(5) & "/" & MonthView1.Month
    Else
        .Range("J2") = ""
    End If
End With

PreencheListaPresençaComAlunos
Application.Visible = True
Menu_Relatorios.Hide
ActiveSheet.PrintPreview
Menu_Relatorios.Show

Else
    MsgBox "Escolha uma Classe.", vbInformation, "Impressão"
End If
End Sub

Private Sub CommandButton2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
TextBox1.Text = "Escolha uma Classe e imprima a folha de presença"
End Sub

Private Sub CommandButton3_Click()
Menu_Relatorios.Hide
Menu_Principal.Show
End Sub


Private Sub CommandButton3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
TextBox1.Text = "Sai do Menu Relatórios"
End Sub

Private Sub CmdAniversario_Click()
'Application.Visible = True
BuscaAniver = False
Menu_Relatorios.Hide
'Menu_Principal.Hide
PreencheListaAniversariantes
LimpaAniver
Menu_Relatorios.Show
Application.Visible = False

End Sub



Private Sub CmdAniversario_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
TextBox1.Text = "Click para Visualizar os aniversariantes do mês"
End Sub




Private Sub MonthView1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As stdole.OLE_XPOS_PIXELS, ByVal Y As stdole.OLE_YPOS_PIXELS)
TextBox1.Text = "Selecione o Mês para diversas consultas."
End Sub

Private Sub TextBox1_Change()
CommandButton3.SetFocus
End Sub

Private Sub TextBox1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = 0
CommandButton3.SetFocus
End Sub

Private Sub TextBox1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
CommandButton3.SetFocus
End Sub

Private Sub UserForm_Activate()
Dim i As Integer
Application.Visible = False
BackColor = RGB(137, 250, 250)
MonthView1.MonthBackColor = RGB(137, 250, 250)
Label1.BackColor = RGB(137, 250, 250)
ComboBox1.Clear
preenche_Combo_Classes
For i = 2 To UBound(todasAsClasses)
     Me.ComboBox1.AddItem (todasAsClasses(i))
 Next i
 ComboBox1.Text = "Escolha uma Classe"
End Sub

Sub PreencheListaPresençaComAlunos()
Dim linhaPresença As Integer
Dim k As Integer
linha = 2
linhaPresença = 3
While Sheets("Alunos").Cells(linha, 1) <> Empty
If InStr(UCase(Sheets("Alunos").Cells(linha, 4).Value), UCase(ComboBox1.Text)) > 0 Then
With Sheets("Presença_" & ComboBox1.Text & "-" & Month(Me.MonthView1.Value) & "-" & Year(Me.MonthView1.Value))
For k = 1 To 9
If (linhaPresença Mod 2) = 1 Then Cells(linhaPresença, k).Interior.Color = RGB(183, 225, 251)
Next
.Cells(linhaPresença, 2) = Sheets("Alunos").Cells(linha, 1).Value
.Cells(linhaPresença, 3) = Sheets("Alunos").Cells(linha, 2).Value
.Cells(linhaPresença, 4) = Sheets("Alunos").Cells(linha, 3).Value

linhaPresença = linhaPresença + 1

End With
End If
linha = linha + 1
Wend
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
TextBox1.Text = ""
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
Menu_Relatorios.Hide
Menu_Principal.Show

End Sub

Function PlanExiste(Nome As String) As Boolean
Dim i As Integer
PlanExiste = False
For i = 1 To ActiveWorkbook.Worksheets.Count
If ActiveWorkbook.Worksheets(i).Name = Nome Then
PlanExiste = True
Exit For
End If
Next i
End Function
