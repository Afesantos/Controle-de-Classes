VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Filtrar 
   Caption         =   "Filtrar"
   ClientHeight    =   4965
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7545
   OleObjectBlob   =   "Filtrar.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Filtrar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private textoDigitado As String






Private Sub ListBox1_Click()
TextBox2.Text = ListBox1.List(ListBox1.ListIndex)
TextBox1.SetFocus
End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

On Error GoTo ErrHandler

pItemFiltrado = ListBox1.List(ListBox1.ListIndex, 1)

    pRespFiltro = pItemFiltrado
    pFiltrado = True
    Select Case pQuemChamouFiltro
        Case Is = "Alunos"
        Alunos.linha = pItemFiltrado
        Alunos.Procura
        Case Is = "Professores"
        Professores.linha = pItemFiltrado
        Professores.Procura
    End Select
Unload Me
Exit Sub
ErrHandler:
Me.TextBox1.SetFocus
End Sub

Private Sub TextBox1_Change()
textoDigitado = TextBox1.Text
Call PreencheLista
End Sub

Private Sub UserForm_Activate()
TextBox1.BackColor = RGB(181, 230, 29)
ListBox1.ForeColor = RGB(49, 74, 108)
If pQuemChamouFiltro <> "" Then
Sheets(pQuemChamouFiltro).Select
Caption = pQuemChamouFiltro

ListaTodos

 Me.TextBox1.SetFocus
 'pQuemChamouFiltro = ""
 End If
End Sub

Sub ListaTodos()
Dim i As Integer
Dim Aluno As String
i = 2
Aluno = Sheets(pQuemChamouFiltro).Range("A" & i)
 While Aluno <> Empty
 ListBox1.AddItem (Aluno)
Me.ListBox1.List(i - 2, 0) = Aluno
Me.ListBox1.List(i - 2, 1) = Sheets(pQuemChamouFiltro).Range("A" & i).Row
i = i + 1
Aluno = Sheets(pQuemChamouFiltro).Range("A" & i)
Wend
End Sub



Private Sub PreencheLista()
Dim Cont As Integer
'código que irá filtrar os nomes
Dim linha As Integer
Dim TextoCelula As String
Cont = 0
linha = 2
'limpa os dados do formulário
ListBox1.Clear
'Irá executar até o último nome
While ActiveSheet.Cells(linha, 1).Value <> Empty
'pega o nome atual
TextoCelula = ActiveSheet.Cells(linha, 1).Value
'quebra a palavra atual pela esquerda conforme a quantidade de letras digitadas e compara com o texto digitado
If InStr(UCase(TextoCelula), UCase(textoDigitado)) > 0 Then
'se a comparação for igual será adicionado no formulario
ListBox1.AddItem ActiveSheet.Cells(linha, 1)
ListBox1.List(Cont, 0) = ActiveSheet.Cells(linha, 1)
ListBox1.List(Cont, 1) = linha
Cont = Cont + 1
End If
linha = linha + 1
Wend
End Sub

