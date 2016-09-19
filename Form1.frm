VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FomPesquisa 
   Caption         =   "Pesquisar "
   ClientHeight    =   3528
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   6264
   LinkTopic       =   "Form1"
   ScaleHeight     =   5000
   ScaleMode       =   0  'User
   ScaleWidth      =   6264
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnOrg 
      Caption         =   "Organizar"
      Height          =   492
      Left            =   3600
      TabIndex        =   5
      Top             =   2880
      Width           =   972
   End
   Begin VB.CommandButton BtnCancelar 
      Caption         =   "Cancelar"
      Height          =   492
      Left            =   5280
      TabIndex        =   4
      Top             =   2880
      Width           =   852
   End
   Begin VB.CommandButton Btnok 
      Caption         =   "Consultar"
      Height          =   492
      Left            =   2520
      TabIndex        =   3
      Top             =   2880
      Width           =   972
   End
   Begin VB.TextBox txt_busca 
      BackColor       =   &H0080FFFF&
      Height          =   288
      Left            =   2280
      TabIndex        =   2
      Top             =   120
      Width           =   2892
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2172
      Left            =   720
      TabIndex        =   0
      Top             =   720
      Width           =   5484
      _ExtentX        =   9673
      _ExtentY        =   3831
      _Version        =   393216
      Rows            =   4
      Cols            =   5
      AllowUserResizing=   1
      FormatString    =   ""
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Digite o Candidato"
      Height          =   192
      Left            =   720
      TabIndex        =   1
      Top             =   120
      Width           =   1332
   End
End
Attribute VB_Name = "FomPesquisa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim conn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim sql As String
Dim sConnString As String
Dim nomeDB As String

'Funcao preenche grid
Public Function preencheFlexGrid(FlexGrid As Object, rs As Object) As Boolean
           
On Error GoTo flex

'verifica os tipos dos objetos
If Not TypeOf FlexGrid Is MSFlexGrid Then Exit Function
If Not TypeOf rs Is ADODB.Recordset Then Exit Function

Dim i As Integer
Dim J As Integer
   
   'define linha e coluna do msflexgrid
   FlexGrid.FixedRows = 1
   FlexGrid.FixedCols = 0
    
   'se o recordset tiver dados então...
   If Not rs.EOF Then
    
       FlexGrid.Rows = rs.RecordCount + 1
       FlexGrid.Cols = rs.Fields.Count
    
       'preenche o msflexgrid com cabeçalho
       For i = 0 To rs.Fields.Count - 1
           FlexGrid.TextMatrix(0, i) = rs.Fields(i).Name
          Next
    
       i = 1
        
       Do While Not rs.EOF
    
           For J = 0 To rs.Fields.Count - 1
               If Not IsNull(rs.Fields(J).Value) Then
                   FlexGrid.TextMatrix(i, J) = rs.Fields(J).Value
               End If
           Next
    
       i = i + 1
       rs.MoveNext
       Loop
    
   End If
preencheFlexGrid = True

flex:
   preencheFlexGrid = False
   Exit Function
End Function

Private Sub BtnCancelar_Click()
End
End Sub

Private Sub Btnok_Click()
On Error GoTo trataerro
'Conexao
sConnString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\contatos.mdb;Persist Security Info=False"
conn.Open sConnString
'Pesquisa o Nome do candidato
rs.Open "SELECT * FROM tab_contatos WHERE candidato Like '%" & txt_busca.text & "%'", conn, adOpenStatic, adLockOptimistic

If rs.EOF Then
   MsgBox "Candidato inexistente"
   txt_busca.SetFocus
Else
     'Chama funcao para preencher o grid
     preencheFlexGrid MSFlexGrid1, rs
     
End If
rs.Close
conn.Close
Exit Sub

trataerro:
MsgBox Err.Number & vbCrLf & Err.Description
End Sub

Private Sub btnOrg_Click()
On Error GoTo trataerro
'Conexao
sConnString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\contatos.mdb;Persist Security Info=False"
conn.Open sConnString
'Pesquisa o Nome do candidato
rs.Open "SELECT * FROM tab_contatos WHERE candidato order by candidato", conn, adOpenStatic, adLockOptimistic

     preencheFlexGrid MSFlexGrid1, rs

rs.Close
conn.Close
Exit Sub

trataerro:
MsgBox Err.Number & vbCrLf & Err.Description
End Sub
