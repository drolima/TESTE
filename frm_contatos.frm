VERSION 5.00
Begin VB.Form frm_contatos 
   Caption         =   "RH - CONTATOS"
   ClientHeight    =   2580
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   5388
   LinkTopic       =   "Form1"
   ScaleHeight     =   2580
   ScaleWidth      =   5388
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Btn_sair 
      Caption         =   "Sair"
      Height          =   372
      Left            =   4320
      TabIndex        =   10
      Top             =   2160
      Width           =   852
   End
   Begin VB.CommandButton Btn_pesq 
      Caption         =   "Pesquisar"
      Height          =   372
      Left            =   2400
      TabIndex        =   9
      Top             =   2160
      Width           =   972
   End
   Begin VB.CommandButton Btn_Cad 
      Caption         =   "Cadastrar"
      Height          =   372
      Left            =   1440
      TabIndex        =   8
      Top             =   2160
      Width           =   972
   End
   Begin VB.TextBox txtdatacont 
      Height          =   288
      Left            =   1440
      MaxLength       =   10
      TabIndex        =   7
      Top             =   1320
      Width           =   1212
   End
   Begin VB.TextBox txtFornecel 
      Height          =   288
      Left            =   3840
      MaxLength       =   14
      TabIndex        =   5
      Top             =   960
      Width           =   1452
   End
   Begin VB.TextBox txtFones 
      Height          =   288
      Left            =   1440
      MaxLength       =   14
      TabIndex        =   3
      Top             =   960
      Width           =   1572
   End
   Begin VB.TextBox txtCandidato 
      Height          =   288
      Left            =   1440
      TabIndex        =   1
      Top             =   480
      Width           =   3852
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Data_Contato"
      Height          =   192
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   984
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Celular "
      Height          =   192
      Left            =   3120
      TabIndex        =   4
      Top             =   960
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Fone Res."
      Height          =   192
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   744
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Candidato"
      Height          =   192
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   744
   End
End
Attribute VB_Name = "frm_contatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Para desabilitar o botão Fechar
 
Private Declare Function GetSystemMenu Lib "user32" _
(ByVal hwnd As Long, ByVal bRevert As Long) As Long

Private Declare Function GetMenuItemCount Lib "user32" _
(ByVal hMenu As Long) As Long

Private Declare Function DrawMenuBar Lib "user32" _
(ByVal hwnd As Long) As Long

Private Declare Function RemoveMenu Lib "user32" _
(ByVal hMenu As Long, ByVal nPosition As Long, _
ByVal wFlags As Long) As Long
Const MF_BYPOSITION = &H400&
Const MF_REMOVE = &H1000&

Function FormataTelefone(ByVal text As String) As String
Dim i As Long

' ignora vazio
If Len(text) = 0 Then Exit Function
 'verifica valores invalidos
  For i = Len(text) To 1 Step -1
    If InStr("0123456789", Mid$(text, i, 1)) = 0 Then
       text = Left$(text, i - 1) & Mid$(text, i + 1)
    End If
  Next
  ' ajusta a posicao correta
  If Len(text) <= 7 Then
     FormataTelefone = Format$(text, "!@@@-@@@@")
  ElseIf Len(text) > 7 And Len(text) <= 9 Then
     FormataTelefone = Format$(text, "!(@@) @@@-@@@@")
  ElseIf Len(text) > 9 Then
     FormataTelefone = Format$(text, "!(@@) @@@@-@@@@")
  End If
End Function
 
 Private Sub cadastrar()

  Dim lstRSQL As String
  Dim cnnTeste As New ADODB.Connection ' Banco
  Dim cnnComando As New ADODB.Command
  Dim rs As New ADODB.Recordset 'Recordset
  Dim vOk As Integer

  cnnTeste.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
  "Data Source=" & App.Path & "\contatos.mdb;"

   cnnTeste.Open
   With cnnComando
   .ActiveConnection = cnnTeste
   .CommandText = "tab_contatos"

      'Chama a classe de banco
      'Set objData = New restful
      'objData.Abrir_BD "C:\douglas\contatos.mdb"
      'Set cnnTeste = objData.Abrir_Dados

      'Set objData = Nothing
      'Set RS = db.OpenRecordset("AuthoRS", dbOpenDynaset)

      'Do Until RS.EOF
      'List1.AddItem RS("Author")
      'RS.MoveNext
      'Loop
      'db.Close

      If Trim$(txtCandidato.text) = "" Or Trim$(txtFones.text) = "" Then
         MsgBox "Obrigatório Informar todos os dados", vbExclamation, Me.Caption
         txtCandidato.SetFocus
         GoTo Fim_Cmd_Confirma_Click
      End If

        .CommandText = "SELECT candidato,fone_celular,fone_residencial,Data_contato FROM tab_contatos WHERE candidato = '" & txtCandidato.text & "'"
         Set rs = .Execute

      If rs.EOF Then
          txtCandidato.text = UCase(RTrim(txtCandidato.text))

        .CommandText = "insert into tab_contatos (candidato, fone_celular,fone_residencial,Data_contato) VALUES ('" & txtCandidato.text & "', '" & txtFones.text & "', '" & txtFornecel.text & "', '" & txtdatacont.text & "')"
        Set rs = .Execute
        MsgBox " Cadastro Realizado com sucesso", vbInformation
        limpa
      Else
        MsgBox "Candidato já cadastrado!!!"
        txtCandidato.SetFocus
        limpa
      End If

Fim_Cmd_Confirma_Click:
    Set rs = Nothing
        Exit Sub
Erro_Cmd_Confirma_Click:
   
    MsgBox "Erro " & Err.Description & " na correção", vbExclamation, "ERRO"

End With

End Sub

Private Sub limpa()
txtCandidato.text = ""
txtFones.text = ""
txtFornecel.text = ""
txtdatacont.text = ""
End Sub

Private Sub Btn_Cad_Click()
cadastrar
End Sub

Private Sub Btn_pesq_Click()
FomPesquisa.Show
End Sub

Private Sub Btn_sair_Click()
End
End Sub

Private Sub Form_Load()
'Funcao para inibir o botao fechar
Dim hSysMenu As Long
Dim nCnt As Long

' FiRSt, show the form
Me.Show

' Get handle to our form's system menu
' (Restore, Maximize, Move, close etc.)
hSysMenu = GetSystemMenu(Me.hwnd, False)

If hSysMenu Then
' Get System menu's menu count
nCnt = GetMenuItemCount(hSysMenu)

If nCnt Then
' Menu count is based on 0 (0, 1, 2, 3...)

RemoveMenu hSysMenu, nCnt - 1, MF_BYPOSITION Or MF_REMOVE
RemoveMenu hSysMenu, nCnt - 2, MF_BYPOSITION Or MF_REMOVE
' Remove the seperator

DrawMenuBar Me.hwnd
' Force caption bar's refresh. Disabling X button

'Me.Caption = "Try to close me!"
End If
End If
End Sub

Private Sub txtdatacont_Change()
'Funcao Data
If Len(txtdatacont) = 2 Then
    txtdatacont = txtdatacont + "/"
    txtdatacont.SelStart = 4
End If
    If Len(txtdatacont) = 5 Then
       txtdatacont = txtdatacont + "/"
       txtdatacont.SelStart = 7
    End If
End Sub

Private Sub txtFones_Validate(keepfocus As Boolean)
'Funcao Fone
If Not IsNumeric(txtFones.text) Or Len(txtFones.text) < 4 Then
   keepfocus = True
   MsgBox "Informe um valor valido !", vbInformation, "Formatando telefone"
   Exit Sub
End If
    txtFones.text = FormataTelefone(txtFones.text)
End Sub

Private Sub txtFornecel_Validate(keepfocus As Boolean)
If Not IsNumeric(txtFornecel.text) Or Len(txtFornecel.text) < 4 Then
   keepfocus = True
   MsgBox "Informe um valor valido !", vbInformation, "Formatando telefone"
   Exit Sub
End If
txtFornecel.text = FormataTelefone(txtFornecel.text)
End Sub
