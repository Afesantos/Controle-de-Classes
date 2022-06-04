VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Classes 
   Caption         =   "Classes"
   ClientHeight    =   9480
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11610
   OleObjectBlob   =   "Classes.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Classes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim linha As Integer
 Dim UltimaLinha As Long
 Dim IncluirClasse As Boolean
 Dim EditarClasse As Boolean
 Dim lRow As Long
Sub Procura()
Cells(linha, 1).Select
Me.TextBox1.Text = ActiveWorkbook.Sheets("Classes").Cells(linha, 1).Value
Me.TextBox2.Text = ActiveWorkbook.Sheets("Classes").Cells(linha, 2).Value
Me.TextBox3.Text = ActiveWorkbook.Sheets("Classes").Cells(linha, 3).Value
Me.TextBox5.Text = ActiveWorkbook.Sheets("Classes").Cells(linha, 4).Value
Me.TextBox6.Text = ActiveWorkbook.Sheets("Classes").Cells(linha, 5).Value
Me.Label12.Caption = "Registro: " & ActiveCell.Row - 1 & " / " & UltimaLinha - 1
ComboBox1.Text = ActiveCell.Value
fUltimaLinha
Me.Label12.Caption = "Registro: " & ActiveCell.Row - 1 & " / " & UltimaLinha - 1
End Sub



Private Sub cmdAlterar_Click()
ComboBox1.Visible = False
CorEdicao
BotaoVisivel
Label4.Caption = "Edição de Classe"
EditarClasse = True
TextBox1.SetFocus
End Sub

Private Sub cmdAnterior_Click()
linha = linha - 1
If linha <= 2 Then linha = 2 Else
Procura

End Sub

Private Sub cmdExcluir_Click()
Dim resultado As VbMsgBoxResult
Dim Plan_a_Excluir As String
If ExisteAlunos(TextBox1.Text) Then
MsgBox "Impossível Excluir." & vbCrLf & "Exclua os Alunos cadastrados nesta Classe primeiro.", vbCritical, "Exclussão de Classe"
Exit Sub
End If
Sheets("Classes").Select
resultado = MsgBox("Tem certeza que deseja EXCLUIR essa Classe?", vbYesNo, "EXCLUSÃO DE CLASSE")
If resultado = vbYes Then
Plan_a_Excluir = TextBox1.Text
 Selection.EntireRow.Delete
 fUltimaLinha
 linha = linha - 1
If linha <= 2 Then linha = 2 Else
Procura
preenche_Combo_Classes
On Error GoTo Final
If Len(Sheets("Presença_" & Plan_a_Excluir)) > 0 Then
Resp = MsgBox("Excluir também a Lista de Presença?", vbYesNo, "Mensagem")
If Resp = vbYes Then
Sheets("Presença_" & Plan_a_Excluir).Delete

End If
End If
End If
Final:
Exit Sub
End Sub


Private Sub cmdIncluir_Click()
IncluirClasse = True
ComboBox1.Visible = False
BotaoVisivel
CorEdicao
Label4.Caption = "Inclusão de Classe"
Me.ImgNovo.Visible = True
Me.TextBox1.Text = ""
Me.TextBox2.Text = ""
Me.TextBox3.Text = ""
Me.TextBox5.Text = ""
Me.TextBox6.Text = ""
Cells(UltimaLinha + 1, 1).Select
Me.TextBox1.SetFocus
End Sub

Private Sub cmdPrimeiro_Click()
linha = 2
Procura
End Sub

Private Sub cmdProximo_Click()
linha = linha + 1
Cells(linha, 1).Select
If linha >= UltimaLinha Then linha = UltimaLinha Else
Procura
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub CommandButton5_Click()
Me.TextBox1.Text = ""
Me.TextBox2.Text = ""
Me.TextBox3.Text = ""
Me.TextBox1.SetFocus
End Sub

Private Sub cmdUltimo_Click()
linha = UltimaLinha
Procura
End Sub

Private Sub ComboBox1_Change()
fUltimaLinha
For lRow = 2 To UltimaLinha
Cells(lRow, 1).Select
If ComboBox1.Value = Cells(lRow, 1).Value Then PreencheTela: Exit Sub
Next
End Sub

Private Sub cmdProcurar_Click()
pQuemChamouFiltro = "Classes"
Filtrar.Show
End Sub

Private Sub CommandButton7_Click()
CorPadrao
If IncluirClasse = True Or EditarClasse = True Then
    If IsEmpty(TextBox1.Value) And IsEmpty(TextBox2.Value) And IsEmpty(TextBox3.Value) Then
    Label4.Caption = "Classes"
    
    With Sheets("Classes") '<- troque o nome da planilha, se necessário
        'Obtém última linha e soma 1:
        lRow = .Cells(.Rows.Count, "A").End(xlUp).Row + 1
        If EditarClasse Then lRow = ActiveCell.Row
         'Preenche dados na planilha:
        .Cells(lRow, "A") = TextBox1.Value
        .Cells(lRow, "B") = TextBox2.Value
        .Cells(lRow, "C") = TextBox3.Value
        .Cells(lRow, "D") = TextBox5.Value
        .Cells(lRow, "E") = TextBox6.Value
        .Cells(lRow, "F") = TextBox7.Value
    End With
pNovaClasse = Me.TextBox1.Text
If IncluirClasse = True Then CriaListaPres
fUltimaLinha
preenche_Combo_Classes
If IncluirClasse Then
MsgBox "Classe Criada com sucesso!", vbExclamation
Else
BotaoVisivel
MsgBox "Classe Alterada com sucesso!", vbExclamation
End If
       IncluirClasse = False
       EditarClasse = False
Else: MsgBox "Favor Preencher todos os campos.", vbCritical, "Cadastro de Classes"
End If

End If
preenche_Combo_Classes
ComboBox1.Visible = True
End Sub

Private Sub CommandButton8_Click()
Me.ImgNovo.Visible = False
IncluirClasse = False
EditarClasse = False
BotaoVisivel
Label4.Caption = "Classes"
CorPadrao
Cells(linha, 1).Select
If linha <= 2 Then linha = 2
ComboBox1.Visible = True
Procura
End Sub





Private Sub TextBox2_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TextBox3_Change()
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub UserForm_Activate()
Dim i As Integer
CorPadrao
Sheets("Classes").Select
Procura
Me.Frame1.SetFocus
If pCadastrarClasses Then
pCadastrarClasses = False
Call cmdIncluir_Click
End If
'TextBox1.BackColor = RGB(181, 230, 29)
Me.Frame1.BackColor = RGB(0, 64, 128)
preenche_Combo_Classes
For i = 2 To UBound(todasAsClasses)
 Me.ComboBox1.AddItem (todasAsClasses(i))
 Next i
End Sub

Private Sub UserForm_Initialize()
 linha = 2
 fUltimaLinha
Me.Label12.Caption = "Registro: " & ActiveCell.Row - 1 & " / " & UltimaLinha - 1

End Sub

Function fUltimaLinha() As Integer
UltimaLinha = Sheets("Classes").Range("A" & Rows.Count).End(xlUp).Row
End Function

Sub CorPadrao()
Dim cCont As Control
    For Each cCont In Controls
        If TypeName(cCont) = "TextBox" Then
            cCont.BackColor = RGB(255, 255, 255)
            cCont.ForeColor = RGB(171, 14, 21)
        End If
     Next cCont

ComboBox1.BackColor = RGB(255, 255, 255)
End Sub

Sub CorEdicao()
Dim cCont As Control
    For Each cCont In Controls
        If TypeName(cCont) = "TextBox" Then
            cCont.BackColor = RGB(255, 254, 0)
            cCont.ForeColor = RGB(0, 0, 0)
        End If
     Next cCont
     ComboBox1.BackColor = RGB(255, 254, 0)
End Sub

Sub PreencheTela()
 With Sheets("Classes") '<- troque o nome da planilha, se necessário
         'Preenche dados na planilha:
        TextBox1.Value = .Cells(lRow, "A")
        TextBox2.Value = .Cells(lRow, "B")
        TextBox3.Value = .Cells(lRow, "C")
        TextBox5.Value = .Cells(lRow, "D")
        TextBox6.Value = .Cells(lRow, "E")
        TextBox7.Value = .Cells(lRow, "F")
    End With
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



