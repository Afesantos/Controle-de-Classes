Attribute VB_Name = "Módulo1"
Public pSala As String
Public pConsultaProfessores As Boolean
Public pCadastrarAlunos As Boolean
Public pCadastrarClasses As Boolean
Public pConsulta As Boolean
Public todasAsClasses() As String
Public pItemFiltrado As Integer 'Célula encontrada no Filtro
Public pQuemChamouFiltro As String
Public pRespFiltro As Integer
Public pFiltrado As Boolean
Public linha As Integer
Public pNovaClasse As String
Public CalenDomingos(5) As String
Public BuscaAniver As Boolean
Public ArquivoExiste As Boolean

Public Function ArquivoExiste(Arquivo)
Dim Arquivo As String
ArquivoExiste = False
If Dir(Arquivo) <> vbNullString Then ArquivoExiste = True
End If
End Function


Sub preenche_Combo_Classes()
Dim i, UltimaLinha As Integer
UltimaLinha = Sheets("Classes").Range("A" & Rows.Count).End(xlUp).Row
ReDim todasAsClasses(2 To UltimaLinha)
For i = 2 To UltimaLinha
todasAsClasses(i) = Sheets("Classes").Cells(i, 1).Value
Next i
End Sub


Sub CriaListaPres()
Dim NewNamePlan As String
NewNamePlan = "Presença_" & pNovaClasse & "-" & Month(Menu_Relatorios.MonthView1.Value) & "-" & Year(Menu_Relatorios.MonthView1.Value)
ActiveWorkbook.Sheets("Presença_Padrão").Copy _
       after:=ActiveWorkbook.Sheets(Sheets.Count)
ActiveWorkbook.Sheets(Sheets.Count).Select
ActiveWorkbook.ActiveSheet.Name = NewNamePlan
Sheets(NewNamePlan).Range("C1") = "Classe: " & pNovaClasse

End Sub

Sub Calcula_Domingos()
Dim dateValue As Date

Dim i, d As Integer
Dim Ddata As String
d = 0
For i = 1 To 31
On Error Resume Next
Ddata = i & "/" & Month(Menu_Relatorios.MonthView1.Value) & "/" & Year(Menu_Relatorios.MonthView1.Value)
dateValue = CDate(Ddata)
If Weekday(dateValue) = vbSunday Then
d = d + 1
CalenDomingos(d) = i
End If
Next
End Sub


Public Function ExisteAlunos(pClasse As String) As Boolean
Dim result As Boolean
Dim linha As Integer
result = False
linha = 2
Sheets("Alunos").Select
Do While Cells(linha, 4).Value <> ""
If Cells(linha, 4).Value = pClasse Then result = True
linha = linha + 1
Loop
ExisteAlunos = result
End Function

Public Function PlanilhaExiste(NomePlan As String) As Boolean
Dim i As Integer
Dim result As Boolean
For i = 1 To Worksheets.Count
If UCase(Worksheets(1).Name) = UCase(NomePlan) Then
result = True
Exit For
End If
Next
End Function

Sub Retângulo1_Clique()
Menu_Principal.Show
End Sub

Sub LimpaAniver()
Dim j As Integer
j = 3
Do While Sheets("Aniversariantes").Cells(j, 2) <> Empty
' limpa após imprimir
Sheets("Aniversariantes").Cells(j, 1).Value = ""
Sheets("Aniversariantes").Cells(j, 2).Value = ""
Sheets("Aniversariantes").Cells(j, 3).Value = ""
Sheets("Aniversariantes").Cells(j, 4).Value = ""
Sheets("Aniversariantes").Cells(j, 5).Value = ""
j = j + 1
Loop
Sheets("Aniversariantes").Range("C1") = "Mês"
End Sub

Sub PreencheListaAniversariantes()
Dim i As Integer
Dim DtNasc As Date
Dim linha As Integer
Dim ContaLinha As Integer
ContaLinha = 1
linha = 3
i = 2
DtNasc = CDate(Sheets("Alunos").Cells(i, 2))

Do While Sheets("Alunos").Cells(i, 2).Value <> Empty
    DtNasc = CDate(Sheets("Alunos").Cells(i, 2))
    If Month(DtNasc) = Month(Menu_Relatorios.MonthView1.Value) Then
     
    Select Case Menu_Relatorios.ComboBox1
    Case Is = "Classe"
        Sheets("Aniversariantes").Cells(linha, 1).Value = linha - 2
        Sheets("Aniversariantes").Cells(linha, 2).Value = Sheets("Alunos").Cells(i, 1)
        Sheets("Aniversariantes").Cells(linha, 3).Value = Sheets("Alunos").Cells(i, 2)
        Sheets("Aniversariantes").Cells(linha, 4).Value = Sheets("Alunos").Cells(i, 3)
        Sheets("Aniversariantes").Cells(linha, 5).Value = Sheets("Alunos").Cells(i, 4)
        'if mod(linha) = 0 then Sheets("Aniversariantes").Range(
        linha = linha + 1
    Case Is <> "Classe"
        If Sheets("Alunos").Cells(i, 4).Value = Menu_Relatorios.ComboBox1.Value Then
            Sheets("Aniversariantes").Cells(linha, 1).Value = linha - 2
            Sheets("Aniversariantes").Cells(linha, 2).Value = Sheets("Alunos").Cells(i, 1)
            Sheets("Aniversariantes").Cells(linha, 3).Value = Sheets("Alunos").Cells(i, 2)
            Sheets("Aniversariantes").Cells(linha, 4).Value = Sheets("Alunos").Cells(i, 3)
            Sheets("Aniversariantes").Cells(linha, 5).Value = Sheets("Alunos").Cells(i, 4)
            linha = linha + 1
        End If
    End Select
        
    End If
i = i + 1
Loop



Select Case linha
Case Is > 3
        If BuscaAniver = False Then
            Sheets("Aniversariantes").Range("C1") = Menu_Relatorios.MonthView1.Value
            Application.Visible = True
            Sheets("Aniversariantes").PrintPreview
            Application.Visible = False
        End If
        If BuscaAniver = True Then
        MsgBox "Temos aniversariantes em " & MonthName(Month(Menu_Relatorios.MonthView1.Value)) & "!", vbInformation, "Aniversariantes"
        End If

Case Is <= 3
         If BuscaAniver = False Then
            MsgBox "Não há aniversariantes em " _
                   & MonthName(Month(Menu_Relatorios.MonthView1.Value)) _
                   & ".", vbInformation, "Aniversariantes"
         End If
End Select
End Sub


Sub GeraXML()
Dim linha As Integer

Pathxml = ActiveWorkbook.Path & "\EBD.xml"
Open Pathxml For Output As #1

Print #1, "<?xml " & "version=" & """" & "1.0" & """" & " encoding=" _
& """" & "ISO-8859-1" & """" & " standalone=" & """" & "yes" & """" & "?>"


Print #1, "<EBDPIBJI>"

Print #1, "<Alunos>"
linha = 2
    Do While Sheets("Alunos").Cells(linha, 1) <> Empty
                    Print #1, Spc(5); "<Aluno Nome='" & Sheets("Alunos").Cells(linha, 1) & "'>"
                    Print #1, Spc(10); "<DtNasc>" & Sheets("Alunos").Cells(linha, 2) & "</DtNasc>"
                    Print #1, Spc(10); "<Idade>" & Sheets("Alunos").Cells(linha, 3) & "</Idade>"
                    Print #1, Spc(10); "<Classe>" & Sheets("Alunos").Cells(linha, 4) & "</Classe>"
                    Print #1, Spc(10); "<Pai>" & Sheets("Alunos").Cells(linha, 5) & "</Pai>"
                    Print #1, Spc(10); "<Mae>" & Sheets("Alunos").Cells(linha, 6) & "</Mae>"
                    Print #1, Spc(10); "<Foto>" & Sheets("Alunos").Cells(linha, 7) & "</Foto>"
                    Print #1, Spc(10); "<Obs>" & Sheets("Alunos").Cells(linha, 8) & "</Obs>"
                    Print #1, Spc(5); "</Aluno>"
    linha = linha + 1
    Loop
Print #1, "</Alunos>"

linha = 2
Print #1, "<Professores>"
    Do While Sheets("Professores").Cells(linha, 1) <> Empty


                    Print #1, Spc(5); "<Professor Nome='" & Sheets("Professores").Cells(linha, 1) & "'>"
                    Print #1, Spc(10); "<Telefone>" & Sheets("Professores").Cells(linha, 2) & "</Telefone>"
                    Print #1, Spc(10); "<Celular>" & Sheets("Professores").Cells(linha, 3) & "</Celular>"
                    Print #1, Spc(10); "<email>" & Sheets("Professores").Cells(linha, 4) & "</email>"
                    Print #1, Spc(10); "<Foto>" & Sheets("Professores").Cells(linha, 5) & "</Foto>"
                    Print #1, Spc(5); "</Professor>"
    linha = linha + 1
    Loop
Print #1, "</Professores>"


linha = 2
Print #1, "<Classes>"
Do While Sheets("Classes").Cells(linha, 1) <> Empty


                    Print #1, Spc(5); "<Classe Nome='" & Sheets("Classes").Cells(linha, 1) & "'>"
                    Print #1, Spc(10); "<IdadeMin>" & Sheets("Classes").Cells(linha, 2) & "</IdadeMin>"
                    Print #1, Spc(10); "<IdadeMax>" & Sheets("Classes").Cells(linha, 3) & "</IdadeMax>"
                    Print #1, Spc(10); "<Prof1>" & Sheets("Classes").Cells(linha, 4) & "</Prof1>"
                    Print #1, Spc(10); "<Prof2>" & Sheets("Classes").Cells(linha, 5) & "</Prof2>"
                    Print #1, Spc(10); "<Obs>" & Sheets("Classes").Cells(linha, 6) & "</Obs>"
                    Print #1, Spc(5); "</Classe>"
    linha = linha + 1
    Loop
Print #1, "</Classes>"

Print #1, "</EBDPIBJI>"


Close #1
End Sub

































