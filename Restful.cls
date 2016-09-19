VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "restful"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_ClsDb As New ADODB.Connection
Private m_clsInsert  As String
Private m_clsConsulta As String
Public Property Get Abrir_Dados() As ADODB.Connection
  ' Set Abrir_Dados = m_ClsDb
End Property

Public Property Get inserir_dados() As String
  ' Set inserir_dados = clsInsert
End Property

Public Property Get Consultar_dados() As String
  ' Set consultar_dados  = clsConsulta
End Property
'-------------------------------------------x---------------------------------------------
'Public Sub Abrir_BD(dbName As String)
'On Error GoTo DBErrHandle
'Dim lstRSQL As String
'Dim cnnTeste As New ADODB.Connection
'Dim cnnComando As New ADODB.Command
'Dim RS As New ADODB.Recordset
'Dim vOk As Integer

'  Set m_ClsDb = cnnTeste.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=C:\douglas\contatos.mdb;"


  
 ' On Error GoTo 0
 ' Exit Sub
'DBErrHandle:
'errnum = Err
'End Sub

'-------------------------------------------*-----------------------------------------------------
'Public Sub inserir_dados(candidato As String, fone_celular As Integer, fone_residencial As Integer, Data_contato As Data)

 'Set m_clsInsert = .CommandText = "insert into tab_contatos (candidato, fone_celular,fone_residencial,Data_contato) VALUES ('" & txtCandidato.Text & "', '" & txtFones.Text & "', '" & txtFornecel.Text & "', '" & txtdatacont.Text & "')"
 'Set RS = .Execute
 'MsgBox " Cadastro Realizado com sucesso", vbInformation

'End Sub
'------------------------------------------------------------------------------------------
'Public Sub cONSULTA(candidato As String, fone_celular As Integer, fone_residencial As Integer, Data_contato As Data)
'On Error GoTo trataerro
'Conexao
'sConnString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\contatos.mdb;Persist Security Info=False"
'conn.Open sConnString
'Pesquisa o Nome do candidato
'rs.Open "SELECT * FROM tab_contatos WHERE candidato Like '%" & txt_busca.Text & "%'", conn, adOpenStatic, adLockOptimistic

'If rs.EOF Then
  ' MsgBox "Candidato inexistente"
   'txt_busca.SetFocus
'Else
     'Chama funcao para preencher o grid
     'preencheFlexGrid MSFlexGrid1, rs
'End If
'rs.Close
'conn.Close
'Exit Sub

'trataerro:
'MsgBox Err.Number & vbCrLf & Err.Description

'End Sub
