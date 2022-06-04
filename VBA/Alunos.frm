VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Alunos 
   Caption         =   "EBD PIBJI"
   ClientHeight    =   10050
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13095
   OleObjectBlob   =   "Alunos.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Alunos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

 Public linha As String
 Dim UltimaLinha As Long
 Dim IncluirAluno As Boolean
 Dim AlterarAluno As Boolean
 Dim fotoAntiga As String
 Dim AlunoMemorizado As Integer

 
Private Sub cmdAlterar_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Me.Label5.Caption = Me.cmdAlterar.ControlTipText
End Sub

Private Sub cmdAnterior_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Me.Label5.Caption = Me.cmdAnterior.ControlTipText
End Sub

Private Sub cmdDesfazer_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Me.Label5.Caption = Me.cmdDesfazer.ControlTipText
End Sub

Private Sub cmdExcluir_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Label5.Caption = Me.cmdIncluir.ControlTipText
End Sub

Private Sub cmdIncluir_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Label5.Caption = Me.cmdIncluir.ControlTipText
End Sub

Private Sub cmdLimpaImg_Click()
ImagemNome = ""
Cells(ActiveCell.Row, 7).Value = ""
Image1.Picture = LoadPicture("")
Me.TxtFoto.Text = "Sem imagem"
Me.Repaint
End Sub

Private Sub cmdPrimeiro_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Me.Label5.Caption = Me.cmdPrimeiro.ControlTipText
End Sub

Private Sub cmdProcurar_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Me.Label5.Caption = Me.cmdProcurar.ControlTipText
End Sub

Private Sub cmdProximo_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Me.Label5.Caption = Me.cmdProximo.ControlTipText
End Sub

Private Sub cmdSair_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Me.Label5.Caption = Me.cmdSair.ControlTipText
End Sub

Private Sub cmdSalvar_Click()
Dirdestino = ActiveWorkbook.Path & "\fotos\"
If (Me.TextBox8.Text <> "Selecione uma Classe") Or (Me.TextBox1.Text <> vbNullString) Then
    If IncluirAluno Or AlterarAluno And Me.ComboBox1.Text <> "Selecione uma Classe" Then
       With Sheets("Alunos")
        .Cells(ActiveCell.Row, 1) = Me.TextBox1.Text
        .Cells(ActiveCell.Row, 2) = Me.TextBox2.Text
        .Cells(ActiveCell.Row, 3) = Me.TextBox3.Text
        .Cells(ActiveCell.Row, 4) = Me.TextBox8.Text
        .Cells(ActiveCell.Row, 5) = Me.TextBox5.Text
        .Cells(ActiveCell.Row, 6) = Me.TextBox6.Text
        .Cells(ActiveCell.Row, 8) = Me.TextBox7.Text
        .Cells(ActiveCell.Row, 9) = Me.TelPai.Text
        .Cells(ActiveCell.Row, 10) = Me.TelMae.Text
       If Not ArqExiste(ThisWorkbook.Path & "\fotos\" & ImagemNome) Then
        .Cells(ActiveCell.Row, 7) = Dir(FileName)
        FileCopy FileName, ThisWorkbook.Path & "\fotos\" & ImagemNome
        End If
       End With
        Label4.Caption = "Alunos"
        CorPadrao
        If IncluirAluno Or AlterarAluno Then BotaoVisivel
    OrderPlanilhaAlunos
    Cells(1, 1).Select
    IncluirAluno = False
    AlterarAluno = False
    cmdLimpaImg.Visible = False
    End If

Else
MsgBox "Escolha uma Classe", vbInformation, "Cadastro de Alunos"
End If
End Sub

Private Sub cmdAlterar_Click()
Label4.Caption = "Edição de Alunos"
CorEdicao
BotaoVisivel
AlterarAluno = True
cmdLimpaImg.Visible = True
AlunoMemorizado = ActiveCell.Row
Alunos.Repaint
End Sub

Private Sub cmdAnterior_Click()
sUltimaLinha
Do
linha = linha - 1
If linha < 2 Then Exit Do
Cells(linha, 1).Select
Loop Until ActiveCell.Rows.Hidden = False
If linha < 2 Then
linha = 2
Exit Sub
End If
Procura
Alunos.Repaint
End Sub

Private Sub cmdExcluir_Click()
Dim resultado As VbMsgBoxResult
resultado = MsgBox("Tem certeza que deseja EXCLUIR esse ALUNO?", vbYesNo, "EXCLUSÃO DE ALUNO")
If resultado = vbYes Then
 Selection.EntireRow.Delete
 linha = linha - 1
 If linha <= 2 Then linha = 2
Procura
End If
End Sub

Private Sub cmdIncluir_Click()
sUltimaLinha
AlunoMemorizado = ActiveCell.Row
Label4.Caption = "Cadastro de Alunos"
Call CorEdicao
BotaoVisivel
IncluirAluno = True
cmdLimpaImg.Visible = True
Me.ImgNovo.Visible = True
Me.TextBox1.Text = ""
Me.TextBox2.Text = ""
Me.TextBox3.Text = ""
Me.TextBox5.Text = ""
Me.TextBox6.Text = ""
Me.TextBox7.Text = ""
Me.TextBox8.Text = ""
Me.TelMae.Text = ""
Me.TelPai.Text = ""

Me.ComboBox1.Text = "Selecione uma Classe"
Image1.Picture = LoadPicture(ThisWorkbook.Path & "/fotos/add_foto.bmp")
Cells(UltimaLinha + 1, 1).Select
Me.Label10.Caption = "Registro: " & ActiveCell.Row - 1 & " / " & UltimaLinha
Alunos.Repaint
Me.TextBox1.SetFocus
End Sub

Private Sub cmdPrimeiro_Click()
linha = 1
Do
linha = linha + 1
Cells(linha, 1).Select
If linha > UltimaLinha Then Exit Do

Loop Until ActiveCell.Rows.Hidden = False
If linha >= UltimaLinha Then Exit Sub
Procura
Alunos.Repaint
End Sub

Private Sub cmdProximo_Click()
sUltimaLinha
Do

If linha >= UltimaLinha + 1 Then Exit Do
Cells(linha, 1).Select
linha = linha + 1
Loop Until ActiveCell.Rows.Hidden = False
If linha >= UltimaLinha + 1 Then Exit Sub
Procura
Alunos.Repaint
End Sub

Private Sub cmdSair_Click()
IgnoraIncluir
cmdLimpaImg.Visible = False
ActiveSheet.AutoFilterMode = False
Unload Me
End Sub

Private Sub cmdSalvar_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Me.Label5.Caption = Me.cmdSalvar.ControlTipText
End Sub

Private Sub cmdUltimo_Click()
sUltimaLinha
linha = UltimaLinha
Procura
Alunos.Repaint
End Sub

Private Sub cmdUltimo_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Me.Label5.Caption = Me.cmdUltimo.ControlTipText
End Sub

Private Sub ComboBox1_Change()
Dim cCont As Control
If IncluirAluno = False And AlterarAluno = False Then
linha = 2
If ComboBox1.Text <> "todas as Classes" Then
    For Each cCont In Me.Controls   'limpa todos os registros
        If TypeName(cCont) = "TextBox" Then
            cCont.Text = ""
        End If
     Next cCont

With ActiveSheet      'filtra pela classe
        .AutoFilterMode = False
        .Range("D1").AutoFilter
        .Range("D1").AutoFilter Field:=4, Criteria1:=Me.ComboBox1.Text
End With
sUltimaLinha
linha = 1
CountVisRows
Do
linha = linha + 1
Cells(linha, 1).Select

If linha > UltimaLinha Then Exit Sub
Loop Until ActiveCell.Rows.Hidden = False
End If
If ComboBox1.Value = "todas as Classes" Then
Sheets("Alunos").AutoFilterMode = False
End If
Procura
Else: TextBox8.Text = ComboBox1.Value
End If
End Sub

Private Sub cmdProcurar_Click()
pQuemChamouFiltro = "Alunos"
Filtrar.Show
End Sub



Sub CorPadrao()
Dim cCont As Control
    For Each cCont In Me.Controls
        If TypeName(cCont) = "TextBox" Then
            cCont.BackColor = RGB(255, 255, 255)
            cCont.ForeColor = RGB(37, 108, 2)
        End If
     Next cCont

ComboBox1.BackColor = RGB(255, 255, 255)
End Sub

Sub CorEdicao()
Dim cCont As Control
    For Each cCont In Me.Controls
        If TypeName(cCont) = "TextBox" Then
            cCont.BackColor = RGB(255, 254, 0)
            cCont.ForeColor = RGB(0, 0, 0)
        End If
     Next cCont
     ComboBox1.BackColor = RGB(255, 254, 0)
End Sub



Private Sub cmdDesfazer_Click()
If IncluirAluno Or AlterarAluno Then BotaoVisivel
IgnoraIncluir
cmdLimpaImg.Visible = False
End Sub

Sub IgnoraIncluir()
Label4.Caption = "Alunos"
CorPadrao
If AlterarAluno Or IncluirAluno Then
AlterarAluno = False
IncluirAluno = False
Me.ImgNovo.Visible = False
'If fotoAntiga <> "" Then Me.Image1.Picture = LoadPicture(fotoAntiga)
Cells(AlunoMemorizado, 1).Select
Procura
End If
Alunos.Repaint
End Sub

Sub CountVisRows()
'by Tom Ogilvy
Dim rng As Range
Set rng = Sheets("Alunos").AutoFilter.Range

Label10.Caption = rng.Columns(1). _
   SpecialCells(xlCellTypeVisible).Count - 1 _
   & " / " & rng.Rows.Count - 1 & " Registros"
End Sub




Private Sub CommandButton1_Click()
ChangeFoto
End Sub

Private Sub CommandButton1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Label5.Caption = Me.CommandButton1.ControlTipText
End Sub

Private Sub cmdLipmaImg_Click()
Image1.Picture = LoadPicture("")
Cells(ActiveCell.Row, 7) = ""
End Sub






Private Sub Image1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
ChangeFoto
End Sub

Sub ChangeFoto()
If AlterarAluno Or IncluirAluno Then
SelecionaFoto
TxtFoto.Text = FileName
Me.Image1.Picture = LoadPicture(FileName)
Alunos.Repaint
End If

End Sub

Private Sub Image1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Label5.Caption = Image1.ControlTipText
End Sub

Private Sub Label13_Click()

End Sub

Private Sub TextBox1_Change()

End Sub

Private Sub TextBox1_Exit(ByVal Cancel As MSForms.ReturnBoolean)
If IncluirAluno And TextBox1.Text <> vbNullString Then
VerificaNomeExiste

If NomeExiste Then
resultado = MsgBox("Esse ALUNO já existe, verifique.", , "CADASTRO DE ALUNO")
'If resultado = vbNo Then TelPai.SetFocus
End If
End If
End Sub

Private Sub TextBox1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If Not (IncluirAluno Or AlterarAluno) Then
KeyAscii = 0
End If
End Sub

Private Sub TextBox2_AfterUpdate()

End Sub

Private Sub TextBox2_Change()
'Formata : dd/mm/aa
    If Len(TextBox2) = 2 Or Len(TextBox2) = 5 Then
        TextBox2.Text = TextBox2.Text & "/"
        SendKeys "{End}", True
    End If
End Sub



Private Sub TextBox2_Exit(ByVal Cancel As MSForms.ReturnBoolean)
If Not IsDate(Me.TextBox2.Value) Then
MsgBox "data inválida, tente novamente.", vbCritical, "Cadastro de Alunos"
TextBox2.SetFocus
TextBox2.Text = ""

Exit Sub
End If
If Me.TextBox2.Text <> "" Then TextBox3.Text = Year(Now) - (Year(Me.TextBox2.Value))

End Sub

Private Sub TextBox2_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
'If KeyCode = 46 Then Key = 47
End Sub

Private Sub TextBox2_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
'para permitir que apenas números sejam digitados
    If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
        KeyAscii = 0
    End If
End Sub

Private Sub TextBox3_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = 0
    Application.SendKeys ("{TAB}")
End Sub

Private Sub TextBox3_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    KeyAscii = 0
    Application.SendKeys ("{TAB}")
End Sub



Private Sub TextBox5_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If Not (IncluirAluno Or AlterarAluno) Then
KeyAscii = 0
End If
End Sub



Private Sub TextBox6_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If Not (IncluirAluno Or AlterarAluno) Then
KeyAscii = 0
End If
End Sub


Private Sub TextBox7_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If Not (IncluirAluno Or AlterarAluno) Then
KeyAscii = 0
End If
End Sub

Private Sub TextBox8_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If Not (IncluirAluno Or AlterarAluno) Then
KeyAscii = 0
End If
End Sub

Private Sub UserForm_Activate()
Dim i As Integer
linha = 2
Me.Frame1.BackColor = RGB(0, 162, 132)
CorPadrao
Me.Label10.Caption = "Registro: " & ActiveCell.Row - 1 & " / " & UltimaLinha - 1
Sheets("Alunos").Select
fotoAntiga = ""
Procura
If pConsulta Then
Me.ComboBox1.Text = pSala
pConsulta = False
End If
preenche_Combo_Classes
Me.ComboBox1.AddItem ("todas as Classes")
For i = 2 To UBound(todasAsClasses)
 Me.ComboBox1.AddItem (todasAsClasses(i))
Next i
If pCadastrarAlunos Then
pCadastrarAlunos = False
'Me.ComboBox1.Text = Menu_Principal.ComboBox1.Text
Call cmdIncluir_Click
End If

End Sub

Private Sub UserForm_Initialize()
IncluirAluno = False
AlterarAluno = False
 linha = 2
 sUltimaLinha
 Me.ComboBox1.Clear
 
End Sub

Sub sUltimaLinha()
UltimaLinha = Sheets("Alunos").Range("A" & Rows.Count).End(xlUp).Row
End Sub


Sub Procura()
Dim foto As String
Cells(linha, 1).Select
Me.TextBox1.Text = ActiveWorkbook.Sheets("Alunos").Cells(linha, 1).Value
Me.TextBox2.Text = ActiveWorkbook.Sheets("Alunos").Cells(linha, 2).Value
Me.TextBox3.Text = ActiveWorkbook.Sheets("Alunos").Cells(linha, 3).Value
Me.TextBox5.Text = ActiveWorkbook.Sheets("Alunos").Cells(linha, 5).Value
Me.TextBox6.Text = ActiveWorkbook.Sheets("Alunos").Cells(linha, 6).Value
Me.TextBox7.Text = ActiveWorkbook.Sheets("Alunos").Cells(linha, 8).Value
Me.TelPai.Text = ActiveWorkbook.Sheets("Alunos").Cells(linha, 9).Value
Me.TelMae.Text = ActiveWorkbook.Sheets("Alunos").Cells(linha, 10).Value
Me.Label10.Caption = "Registro: " & ActiveCell.Row - 1 & " / " & UltimaLinha - 1
Me.TextBox8.Text = ActiveWorkbook.Sheets("Alunos").Cells(linha, 4).Value
If pConsulta Then Me.ComboBox1.Text = pSala

If IsEmpty(ActiveWorkbook.Sheets("Alunos").Cells(linha, 7).Value) Then
    Me.Image1.Picture = LoadPicture(ThisWorkbook.Path & "\fotos\ndisp.bmp")
    TxtFoto.Text = "Sem imagem"
    fotoAntiga = ThisWorkbook.Path & "\fotos\ndisp.bmp"
Else
    foto = ActiveWorkbook.Sheets("Alunos").Cells(linha, 7).Value
    TxtFoto.Text = ThisWorkbook.Path & "\fotos\" & foto
    If Dir(ThisWorkbook.Path & "\fotos\" & foto) <> "" Then
        Me.Image1.Picture = LoadPicture(ThisWorkbook.Path & "\fotos\" & foto)
        fotoAntiga = ThisWorkbook.Path & "\fotos\" & foto
    Else
        MsgBox "A imagem:" & vbCrLf & "'" & ThisWorkbook.Path & "\fotos\" & foto & "'" & vbCrLf & "não existe, selecione outra foto!", vbExclamation
        Sheets("Alunos").Cells(ActiveCell.Row, 7) = ""
        TxtFoto.Text = "Imagem perdida"
    End If
End If
Me.Frame1.SetFocus

End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Me.Label5.Caption = ""
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
ActiveSheet.AutoFilterMode = False
End Sub

Private Sub UserForm_Resize()

If pFiltrado Then
linha = pItemFiltrado
Procura
pFiltrado = False
End If
End Sub

Sub BotaoVisivel()
Me.cmdPrimeiro.Visible = Not Me.cmdPrimeiro.Visible
Me.cmdAnterior.Visible = Not Me.cmdAnterior.Visible
Me.cmdProximo.Visible = Not Me.cmdProximo.Visible
Me.cmdUltimo.Visible = Not Me.cmdUltimo.Visible
Me.cmdIncluir.Visible = Not Me.cmdIncluir.Visible
Me.cmdAlterar.Visible = Not Me.cmdAlterar.Visible
'Me.cmdSalvar.Visible = Not Me.cmdSalvar.Visible
Me.cmdExcluir.Visible = Not Me.cmdExcluir.Visible
Me.cmdProcurar.Visible = Not Me.cmdProcurar.Visible
End Sub

Sub VerificaNomeExiste()
Dim Cont As Integer
'código que irá filtrar os nomes
Dim linha As Integer
Dim TextoCelula As String
Cont = 0
linha = 2
'limpa os dados do formulário
NomeExiste = False
'Irá executar até o último nome
While ActiveSheet.Cells(linha, 1).Value <> Empty
'pega o nome atual
TextoCelula = ActiveSheet.Cells(linha, 1).Value
'quebra a palavra atual pela esquerda conforme a quantidade de letras digitadas e compara com o texto digitado
If InStr(UCase(TextoCelula), UCase(TextBox1.Text)) > 0 Then
'se a comparação for igual será adicionado no formulario
NomeExiste = True
Cont = Cont + 1
End If
linha = linha + 1
Wend

End Sub
