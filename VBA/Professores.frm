VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Professores 
   Caption         =   "Professores"
   ClientHeight    =   9480
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11610
   OleObjectBlob   =   "Professores.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Professores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Public linha As String
 Dim UltimaLinha As Long
 Dim IncluirProfessor As Boolean
 Dim EditarProfessor As Boolean
 Dim FileName As String
 

Private Sub cmdAlterar_Click()
EditarProfessor = True
cmdLimpaImg.Visible = True
Label4.Caption = "Editar Professor"
CorEdicao
Me.BotaoVisivel
End Sub

Private Sub cmdAlterar_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Me.Label5.Caption = Me.cmdAlterar.ControlTipText
End Sub

Private Sub cmdAnterior_Click()
linha = linha - 1
If linha < 2 Then linha = 2 Else
Procura
End Sub

Private Sub cmdAnterior_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Me.Label5.Caption = Me.cmdAnterior.ControlTipText
End Sub

Private Sub CmdEmail_Click()
Dim todosProf As String, link As String
Dim linha As Integer
linha = 2
Sheets("Professores").Select
Do While Cells(linha, 1) <> vbNullString
    todosProf = Cells(linha, 4) & ";" & todosProf
    linha = linha + 1
Loop

link = "mailto:" & todosProf
    On Error GoTo NoCanDo
    ActiveWorkbook.FollowHyperlink Address:=link, NewWindow:=True
    Exit Sub
NoCanDo:
    MsgBox "Não foi possível enviar o e-mail para todos os professores.", vbCritical, "Enviando e-mail para todos os Professores"
End Sub

Private Sub CmdEmail_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Me.Label5.Caption = Me.CmdEmail.ControlTipText
End Sub

Private Sub cmdExcluir_Click()
Dim resultado As VbMsgBoxResult
resultado = MsgBox("Tem certeza que deseja EXCLUIR esse Professor?", vbYesNo, "EXCLUSÃO DE PROFESSOR")
If resultado = vbYes Then
ActiveCell.EntireRow.Delete
linha = 2
Procura
Else
MsgBox "Não será feito nada"
End If
End Sub

Private Sub cmdIncluir_Click()
Sheets("Professores").Select
IncluirProfessor = True
cmdLimpaImg.Visible = True
Label4.Caption = "Cadastro de Professores"
CorEdicao
Me.BotaoVisivel
'Cells(UltimaLinha + 1, 1).Select
ImgNovo.Visible = True
Me.TextBox1.Text = ""
Me.TextBox2.Text = ""
Me.TextBox3.Text = ""
Me.TextBox5.Text = ""
Me.Image1.Picture = LoadPicture(ThisWorkbook.Path & "\fotos\add_foto.bmp")
Me.TextBox1.SetFocus
End Sub

Private Sub cmdIncluir_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Label5.Caption = Me.cmdIncluir.ControlTipText
End Sub

Private Sub cmdLimpaImg_Click()
Cells(ActiveCell.Row, 5).Value = ""
Image1.Picture = LoadPicture("")
Me.Repaint
End Sub

Private Sub cmdPrimeiro_Click()
linha = 2
Procura
End Sub

Private Sub cmdPrimeiro_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Me.Label5.Caption = Me.cmdPrimeiro.ControlTipText
End Sub

Private Sub cmdProcurar_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Me.Label5.Caption = Me.cmdProcurar.ControlTipText
End Sub

Private Sub cmdProximo_Click()
linha = linha + 1
If linha > UltimaLinha Then linha = UltimaLinha Else
Procura
End Sub

Private Sub cmdProximo_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Me.Label5.Caption = Me.cmdProximo.ControlTipText
End Sub

Private Sub cmdSair_Click()
cmdLimpaImg.Visible = False
Unload Me
End Sub

Private Sub cmdProcurar_Click()
pQuemChamouFiltro = ActiveSheet.Name
Filtrar.Show
End Sub

Private Sub cmdSair_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Me.Label5.Caption = Me.cmdSair.ControlTipText
End Sub

Private Sub cmdUltimo_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Me.Label5.Caption = Me.cmdUltimo.ControlTipText
End Sub


Private Sub CommandButton1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Label5.Caption = Me.CommandButton1.ControlTipText
End Sub

Private Sub CommandButton6_Click()
Label4.Caption = "Professores"
cmdLimpaImg.Visible = False
If (IncluirProfessor Or EditarProfessor) Then
CorPadrao
Me.BotaoVisivel
If IncluirProfessor Then Cells(UltimaLinha + 1, 1).Select
With Sheets("Professores")
    .Cells(ActiveCell.Row, 1) = Me.TextBox1.Text
    .Cells(ActiveCell.Row, 2) = Me.TextBox2.Text
    .Cells(ActiveCell.Row, 3) = Me.TextBox3.Text
    .Cells(ActiveCell.Row, 4) = Me.TextBox5.Text
    .Cells(ActiveCell.Row, 6) = Me.TextBox6.Text
Me.ImgNovo.Visible = False
MsgBox FileName
If Not IsEmpty(FileName) Then
    .Cells(ActiveCell.Row, 5) = Dir(FileName)
'    If Dir(FileName) = vbNullString Then
        FileCopy FileName, ThisWorkbook.Path & "\fotos\" & Dir(FileName)
'    End If
End If
End With
MsgBox FileName
fUltimaLinha
End If
End Sub

Private Sub CommandButton6_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Me.Label5.Caption = Me.CommandButton6.ControlTipText
End Sub

Private Sub CommandButton7_Click()
If EditarProfessor Or IncluirProfessor Then
Label4.Caption = "Professores"
Me.ImgNovo.Visible = False
CorPadrao
Me.BotaoVisivel
IncluirProfessor = False
EditarProfessor = False
cmdLimpaImg.Visible = False
linha = 2
 UltimaLinha = Sheets("Professores").Range("A" & Rows.Count).End(xlUp).Row
Label10.Caption = "Registro: " & ActiveWorkbook.Sheets("Professores").Cells.Row & " / " & UltimaLinha - 1
 Procura
End If
Me.Repaint
End Sub

Private Sub cmdUltimo_Click()
linha = UltimaLinha
Procura
End Sub


Private Sub CommandButton7_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Me.Label5.Caption = Me.CommandButton7.ControlTipText
End Sub

Private Sub Image1_Click()
ChangeFoto
End Sub

Sub ChangeFoto()
Dim fd As FileDialog
Dim Dirdestino As String
Dim FileChosen As Integer
Dim FileName As String
Set fd = Application.FileDialog(msoFileDialogOpen)
Set FSO = CreateObject("Scripting.FileSystemObject")
If EditarProfessor Or IncluirProfessor Then
'the number of the button chosen
FileChosen = fd.Show
fd.Title = "Selecione a foto"
'fd.InitialFileName = "c:\wise owl\"
fd.InitialView = msoFileDialogViewList
'show Excel workbooks and macro workbooks
fd.Filters.Clear
fd.Filters.Add "Imagens BMP", "*.bmp"
fd.Filters.Add "Imagens JPG", "*.jpg"
fd.FilterIndex = 1
fd.ButtonName = "Escolha o Arquivo"
If FileChosen <> -1 Then
'didn't choose anything (clicked on CANCEL)
MsgBox "Nenhuma foto selecionada!", vbInformation, "Foto do Professor"
Else
'get file, and open it (NAME property
'includes path, which we need)
FileName = fd.SelectedItems(1)
TxtFoto = FileName
ImagemNome = FileName
 Dirdestino = ActiveWorkbook.Path & "\fotos\"
 FSO.CopyFile FileName, Dirdestino

Me.Image1.Picture = LoadPicture(FileName)

End If
End If
Me.Repaint

End Sub

Private Sub Image1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Label5.Caption = Image1.ControlTipText
End Sub

Private Sub TextBox1_Change()

End Sub

Private Sub TextBox1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If Not (IncluirProfessor Or EditarProfessor) Then
KeyAscii = 0
End If
End Sub


Private Sub TextBox2_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If Not (IncluirProfessor Or EditarProfessor) Then
KeyAscii = 0
End If
End Sub



Private Sub TextBox3_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If Not (IncluirProfessor Or EditarProfessor) Then
KeyAscii = 0
End If
End Sub

Private Sub TextBox5_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If Not (IncluirProfessor Or EditarProfessor) Then
KeyAscii = 0
End If
End Sub

Private Sub TextBox5_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Label5.Caption = Me.TextBox5.ControlTipText
End Sub

Private Sub TextBox5_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Dim link As String
If IncluirProfessor = True Then MsgBox "Incluindo"
If (IncluirProfessor = False Or EditarProfessor = False) Then
link = "mailto:" & Me.TextBox5.Text
    On Error GoTo NoCanDo
    ActiveWorkbook.FollowHyperlink Address:=link, NewWindow:=True
    Exit Sub
NoCanDo:
    MsgBox "Não foi possível enviar o e-mail para: " & Me.TextBox5.Text
    Professores.Show
End If
End Sub

Private Sub TextBox6_Change()

End Sub

Private Sub TextBox6_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If Not (IncluirProfessor Or EditarProfessor) Then
KeyAscii = 0
End If
End Sub

Private Sub UserForm_Activate()
Sheets("Professores").Select
IncluirProfessor = False
EditarProfessor = False
Sheets("Professores").Select
Me.Frame1.SetFocus
 If pConsultaProfessores = True Then
     pConsultaProfessores = False
    Call cmdIncluir_Click
    Me.TextBox1.SetFocus
 End If
Me.TextBox5.MousePointer = 99
'Me.TextBox5.MouseIcon = ThisWorkbook.Path & "\fotos\link.cur"
linha = 2
 UltimaLinha = Sheets("Professores").Range("A" & Rows.Count).End(xlUp).Row
 Label10.Caption = "Registro: " & ActiveWorkbook.Sheets("Professores").Cells.Row & " / " & UltimaLinha - 1
 Procura

End Sub

Private Sub UserForm_Initialize()
Me.Frame1.BackColor = RGB(241, 237, 56)

End Sub
Sub Procura()
Dim foto As Scripting.FileSystemObject
Set foto = New Scripting.FileSystemObject
Cells(linha, 1).Select
TextBox1.Text = ActiveWorkbook.Sheets("Professores").Cells(linha, 1).Value
TextBox2.Text = ActiveWorkbook.Sheets("Professores").Cells(linha, 2).Value
TextBox3.Text = ActiveWorkbook.Sheets("Professores").Cells(linha, 3).Value
TextBox5.Text = ActiveWorkbook.Sheets("Professores").Cells(linha, 4).Value
Label10.Caption = "Registro: " & ActiveCell.Row - 1 & " / " & UltimaLinha - 1

If ActiveWorkbook.Sheets("Professores").Cells(linha, 5).Value = "" Then
Me.Image1.Picture = LoadPicture(ThisWorkbook.Path & "\fotos\ndisp.bmp")
TxtFoto = "Sem Imagem"
Else
foto = ThisWorkbook.Path & "\fotos\" & ActiveWorkbook.Sheets("Professores").Cells(linha, 5).Value
TxtFoto = foto
If foto.FileExists Then
    Me.Image1.Picture = LoadPicture(foto)
    Else
    Me.Image1.Picture = LoadPicture(ThisWorkbook.Path & "\fotos\naoEncont.bmp")
End If
Me.Repaint
End If

End Sub
Sub CorPadrao()
Dim cCont As Control
    For Each cCont In Controls
        If TypeName(cCont) = "TextBox" Then
            cCont.BackColor = RGB(255, 255, 255)
            cCont.ForeColor = RGB(37, 108, 2)
        End If
     Next cCont
'Me.TextBox4.ForeColor = &H80000008

End Sub

Sub CorEdicao()
Dim cCont As Control
    For Each cCont In Controls
        If TypeName(cCont) = "TextBox" Then
            cCont.BackColor = RGB(255, 254, 0)
            cCont.ForeColor = RGB(0, 0, 0)
        End If
     Next cCont

End Sub

Function fUltimaLinha() As Integer
UltimaLinha = Sheets("Professores").Range("A" & Rows.Count).End(xlUp).Row
End Function

Sub BotaoVisivel()
Me.cmdPrimeiro.Visible = Not Me.cmdPrimeiro.Visible
Me.cmdAnterior.Visible = Not Me.cmdAnterior.Visible
Me.cmdProximo.Visible = Not Me.cmdProximo.Visible
Me.cmdUltimo.Visible = Not Me.cmdUltimo.Visible
Me.cmdIncluir.Visible = Not Me.cmdIncluir.Visible
Me.cmdAlterar.Visible = Not Me.cmdAlterar.Visible
Me.CmdEmail.Visible = Not Me.CmdEmail.Visible
Me.cmdExcluir.Visible = Not Me.cmdExcluir.Visible
Me.cmdProcurar.Visible = Not Me.cmdProcurar.Visible
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Label5.Caption = ""
End Sub
