VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Menu_Principal 
   Caption         =   "   Menu Principal - EBD PIBJI   "
   ClientHeight    =   7395
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9090
   OleObjectBlob   =   "Menu_Principal.frx":0000
   StartUpPosition =   2  'CenterScreen
   Tag             =   "0"
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "Menu_Principal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()
If OptionButton1 Then
pCadastrarAlunos = True
Alunos.Show
End If
If OptionButton2 Then
pConsultaProfessores = True
Professores.Show
End If
If OptionButton3 Then
pCadastrarClasses = True
Classes.Show
End If
End Sub

Private Sub CommandButton1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
TextBox1.Text = "Click para cadastrar Alunos, Professores ou Classes da EBD."
End Sub

Private Sub CommandButton2_Click()
Me.CommandButton2.Tag = 1
Application.Visible = True
ActiveSheet.AutoFilterMode = False
Sheets("Alunos").Select
Unload Me
'Menu_Principal.Hide
End Sub

Private Sub CommandButton2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
TextBox1.Text = "Sai desta tela e abre a Planilha no Excel"
End Sub

Private Sub CommandButton4_Click()
Application.Visible = True
Sheets("Index").Select
ActiveWorkbook.Save
Application.Quit
End Sub

Private Sub CommandButton4_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
TextBox1.Text = "Sai do Aplicativo e do Excel"
End Sub

Private Sub CommandButton5_Click()
If Me.OptionButton1.Value = True Then
        pSala = "todas as Classes"
        pConsulta = True
        Alunos.Show
    End If
If Me.OptionButton2.Value = True Then
        'pSala = Me.ComboBox1.Value
        Professores.Show
        Exit Sub
    End If
    If Me.OptionButton3.Value = True Then
        'pSala = Me.ComboBox1.Value
        pConsulta = True
        Classes.Show
    End If
End Sub


Private Sub CommandButton5_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
TextBox1.Text = "Faça consulta de Alunos, Professores ou Classes disponíveis."
End Sub

Private Sub CommandButton7_Click()
Menu_Principal.Hide
Menu_Relatorios.Show
End Sub


Private Sub CommandButton7_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
TextBox1.Text = "Imprima a Lista de Chamadas ou Aniversariantes do mês."
End Sub





Private Sub Frame1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
TextBox1.Text = ""
End Sub

Private Sub Frame2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
TextBox1.Text = "Escolha uma opção"
End Sub

Private Sub Label2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Dim link As String
link = "http:\\" & Me.Label2.Caption
    On Error GoTo NoCanDo
    ActiveWorkbook.FollowHyperlink Address:=link, NewWindow:=True
    Exit Sub
NoCanDo:
    MsgBox "Não foi possível abrir a página", vbCritical, "Site da Igreja"

End Sub

Private Sub Label2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
TextBox1.Text = "Duplo click para abrir o site da igreja."
End Sub



Private Sub TextBox1_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Frame1.SetFocus
End Sub

Private Sub UserForm_Initialize()
 Dim linha As Integer
Application.Visible = False
pFiltrado = False
BuscaAniver = True
linha = 2
pConsulta = False
pConsultaProfessores = False
pCadastrarClasses = False
'Do Until Sheets("Classes").Range("a" & linha).Value = ""
'Me.ComboBox1.AddItem Sheets("Classes").Range("A" & linha).Value
'linha = linha + 1
'Loop
'Image1.Picture = LoadPicture("")
If BuscaAniver = True Then
    PreencheListaAniversariantes
    BuscaAniver = False
End If
End Sub


Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
TextBox1.Text = ""
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
Application.Visible = True
ActiveSheet.AutoFilterMode = False
GeraXML
If Me.CommandButton2.Tag = 0 Then Application.Quit
End Sub
