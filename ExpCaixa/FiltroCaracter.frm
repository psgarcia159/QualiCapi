VERSION 5.00
Begin VB.Form FiltroCaracter 
   Caption         =   "Filtro"
   ClientHeight    =   1260
   ClientLeft      =   60
   ClientTop       =   3060
   ClientWidth     =   7290
   ClipControls    =   0   'False
   LinkTopic       =   "FiltroCaracter"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   508.732
   ScaleMode       =   0  'User
   ScaleWidth      =   7320
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtCampo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   135
      TabIndex        =   2
      Top             =   810
      Width           =   5685
   End
   Begin VB.CommandButton btnDesiste 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&Desiste"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5985
      TabIndex        =   1
      Top             =   810
      Width           =   1185
   End
   Begin VB.CommandButton btnConfirmar 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Con&firma"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5985
      TabIndex        =   0
      Top             =   330
      Width           =   1185
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
      Left            =   1275
      TabIndex        =   4
      Top             =   120
      Width           =   4560
   End
   Begin VB.Label lblCampo 
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
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1125
   End
End
Attribute VB_Name = "FiltroCaracter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim controledadosfiltro As Object
Dim GridFiltro As TrueDBGrid70.Columns
Dim xColuna As Integer

Private Sub btnDesiste_Click()
    Filtrou = False
    Unload Me
End Sub

Private Sub btnConfirmar_Click()
    Dim a As Integer, xcampo As String
    Dim Filtro As String
    Dim ResTeste As Object
    
    'Apenas para o formulario de Fonecedores (Serviço/Produto/Produto e Serviço)
     
    If Trim$(txtCampo.Text) <> "" Then
       FiltroCaracter.MousePointer = vbHourglass
       GridFiltro.Item(xColuna).HeadFont.Italic = True
       GridFiltro.Item(xColuna).HeadFont.Bold = True
       Filtrou = True
       Guarda_Filtro GridFiltro.Item(xColuna).DataField, txtCampo.Text
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
       
       FiltroCaracter.MousePointer = vbDefault
       Unload Me
    Else
       MsgBox "Valor deve ser informado.", 32, "ERRO"
       txtCampo.SetFocus
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
Public Property Let ColunaGrid(ByVal valor As Integer)
    xColuna = valor
End Property
