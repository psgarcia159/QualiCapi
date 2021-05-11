VERSION 5.00
Begin VB.Form FiltroCaracter 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Filtro"
   ClientHeight    =   1290
   ClientLeft      =   1335
   ClientTop       =   3045
   ClientWidth     =   7320
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H0000C0C0&
   Icon            =   "FiltroCaracter.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1290
   ScaleWidth      =   7320
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Botao_Conf 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Con&firma"
      Default         =   -1  'True
      Height          =   345
      Left            =   6015
      TabIndex        =   1
      Top             =   345
      Width           =   1185
   End
   Begin VB.CommandButton Botao_canc 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&Desiste"
      Height          =   345
      Left            =   6015
      TabIndex        =   2
      Top             =   825
      Width           =   1185
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   165
      TabIndex        =   0
      Top             =   825
      Width           =   5685
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Campo: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   150
      TabIndex        =   4
      Top             =   135
      Width           =   1125
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   1305
      TabIndex        =   3
      Top             =   135
      Width           =   4560
   End
End
Attribute VB_Name = "FiltroCaracter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim controledadosfiltro As Object
'Dim GridFiltro As TDBGrid
'*Dim GridFiltro As TrueOleDBGrid70.Columns
Dim GridFiltro As TrueDBGrid70.Columns
Dim xColuna As Integer

Private Sub botao_canc_Click()
    Filtrou = False
    Unload Me
End Sub

Private Sub botao_conf_Click()
    Dim a As Integer, xcampo As String
    Dim Filtro As String
    Dim ResTeste As Object
    
    'Apenas para o formulario de Fonecedores (Serviço/Produto/Produto e Serviço)
     
    If Trim$(Text1.Text) <> "" Then
       FiltroCaracter.MousePointer = vbHourglass
       GridFiltro.Item(xColuna).HeadFont.Italic = True
       GridFiltro.Item(xColuna).HeadFont.Bold = True
       Filtrou = True
       Guarda_Filtro GridFiltro.Item(xColuna).DataField, Text1.Text
       'Guarda_Filtro "focl_tx_RazaoSocial", Text1.Text
       Filtro = FiltroOrig
       xcampo = Filtros(1, 1)
       If Tem_Filtro = False Then
          If Filtros(3, 1) = "1" Then    'Tipo String
              If filtra_especial Then
'                 Modificado para dar certo com ADO
'                 Filtro = Filtro + " AND " + xcampo + " LIKE '" + Filtros(2, 1) + "*'"
                 Filtro = Filtro + " AND " + xcampo + " LIKE '" + Filtros(2, 1) + "%'"
              Else
'                  Modificado para dar certo com ADO
'                 Filtro = Filtro + " WHERE " + xcampo + " LIKE '" + Filtros(2, 1) + "*'"
                 Filtro = Filtro + " WHERE " + xcampo + " LIKE '" + Filtros(2, 1) + "%'"
                 
              End If
          ElseIf Filtros(3, 1) = "2" Then 'Tipo Data ou Número
              If filtra_especial Then
                 Filtro = Filtro + " AND " + xcampo + Filtros(2, 1)
              Else
                 Filtro = Filtro + " WHERE " + xcampo + Filtros(2, 1)
              End If
          End If
       Else
          If Filtros(3, 1) = "1" Then    'Tipo String
'             Modificado para dar certo com ADO
'             Filtro = Filtro + " AND " + xcampo + " LIKE '" + Filtros(2, 1) + "*'"
             Filtro = Filtro + " AND " + xcampo + " LIKE '" + Filtros(2, 1) + "%'"
          ElseIf Filtros(3, 1) = "2" Then 'Tipo Data ou Número
             Filtro = Filtro + " AND " + xcampo + Filtros(2, 1)
          End If
       End If
       For a = 2 To ContFil
          xcampo = Filtros(1, a)
          If Filtros(3, a) = "1" Then    'Tipo String
'             Modificado para dar certo com ADO
'             Filtro = Filtro + " AND " + xcampo + " LIKE '" + Filtros(2, a) + "*'"
             Filtro = Filtro + " AND " + xcampo + " LIKE '" + Filtros(2, a) + "%'"
          ElseIf Filtros(3, a) = "2" Then 'Tipo Data ou Número
             Filtro = Filtro + " AND " + xcampo + Filtros(2, a)
          End If
       Next
       
       FiltroAtual = Filtro
       
'       If ConTelaOrdem <> 0 Then
'            controledadosfiltro.RecordSource = funTrataSql(Filtro + " ORDER BY " + Ordem)
'       Else
'            controledadosfiltro.RecordSource = funTrataSql(Filtro)
'       End If

'       controledadosfiltro.Refresh
       
       FiltroCaracter.MousePointer = vbDefault
       Unload Me
    Else
       MsgBox "Valor deve ser informado.", 32, "ERRO"
       Text1.SetFocus
    End If
End Sub

Private Sub Guarda_Filtro(ByVal Campo1 As String, ByVal Campo2 As String)
    Dim CampoTratado As String
    CampoTratado = Campo1
    ContFil = ContFil + 1
    ReDim Preserve Filtros(3, ContFil)
    Filtros(1, ContFil) = CampoTratado   'Nome do campo a ser filtrado
    Filtros(2, ContFil) = Campo2   'Valor do filtro digitado
    Filtros(3, ContFil) = "1"      'Tipo = Caracter
End Sub
Public Property Let ControleDados(ByVal vNewValue As Variant)
    Set controledadosfiltro = vNewValue
End Property
Public Property Let Grid(ByVal vNewValue As Variant)
    Set GridFiltro = vNewValue
End Property
Public Property Let ColunaGrid(ByVal Valor As Integer)
    xColuna = Valor
End Property







