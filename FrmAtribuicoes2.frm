VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form FrmAtribuicoes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Alteração do HH Real"
   ClientHeight    =   3330
   ClientLeft      =   30
   ClientTop       =   1335
   ClientWidth     =   9285
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   9285
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraProjeto 
      Height          =   915
      Left            =   60
      TabIndex        =   6
      Top             =   0
      Width           =   9135
      Begin VB.TextBox TxtSubprojeto 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2280
         TabIndex        =   10
         Top             =   540
         Width           =   5415
      End
      Begin VB.TextBox TxtProjeto 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2280
         TabIndex        =   9
         Top             =   180
         Width           =   5415
      End
      Begin VB.Label LblProjeto 
         Caption         =   "Projeto:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1560
         TabIndex        =   8
         Top             =   240
         Width           =   735
      End
      Begin VB.Label LblSubProjeto 
         Caption         =   "SubProjeto:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1260
         TabIndex        =   7
         Top             =   600
         Width           =   1035
      End
   End
   Begin VB.Frame FraDadosGerais 
      Height          =   1035
      Left            =   60
      TabIndex        =   2
      Top             =   960
      Width           =   9135
      Begin VB.TextBox TxtDia 
         Enabled         =   0   'False
         Height          =   315
         Left            =   780
         TabIndex        =   13
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox TxtRecurso 
         Enabled         =   0   'False
         Height          =   315
         Left            =   3360
         TabIndex        =   12
         Top             =   600
         Width           =   5655
      End
      Begin VB.TextBox TxtTarefas 
         Enabled         =   0   'False
         Height          =   315
         Left            =   780
         TabIndex        =   11
         Top             =   180
         Width           =   8235
      End
      Begin VB.Label LblDia 
         Alignment       =   1  'Right Justify
         Caption         =   "Dia:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   180
         TabIndex        =   5
         Top             =   660
         Width           =   525
      End
      Begin VB.Label LblTarefa 
         Caption         =   "Tarefa:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   615
      End
      Begin VB.Label LblRecurso 
         Caption         =   "Recurso:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   3
         Top             =   660
         Width           =   795
      End
   End
   Begin Threed.SSCommand CmdDesistir 
      Height          =   330
      Left            =   5400
      TabIndex        =   1
      Top             =   2880
      Width           =   1365
      _Version        =   65536
      _ExtentX        =   2408
      _ExtentY        =   582
      _StockProps     =   78
      Caption         =   "&Sair"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSCommand CmdGravar 
      Height          =   330
      Left            =   2400
      TabIndex        =   0
      Top             =   2880
      Width           =   1365
      _Version        =   65536
      _ExtentX        =   2408
      _ExtentY        =   582
      _StockProps     =   78
      Caption         =   "&Gravar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame FraHH 
      Height          =   675
      Left            =   60
      TabIndex        =   14
      Top             =   2040
      Width           =   9135
      Begin VB.TextBox TxtHHPrevisto 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   2400
         TabIndex        =   18
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox TxtHHReal 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5880
         TabIndex        =   15
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label LblHHReal 
         Caption         =   "HH Real:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5040
         TabIndex        =   17
         Top             =   300
         Width           =   795
      End
      Begin VB.Label LblHHPrevisto 
         Alignment       =   1  'Right Justify
         Caption         =   "HH Previsto:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1200
         TabIndex        =   16
         Top             =   300
         Width           =   1155
      End
   End
End
Attribute VB_Name = "FrmAtribuicoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim XFI_PROJETO As Integer 'Código do projeto
Dim XFI_SUBPROJETO As Integer  'Código do Subprojeto
Dim XFL_TAREFA As Long
Dim XFL_RECURSO As Long
Dim XFL_ATRIBUICAO As Long
Dim XFD_DIA As Date

Sub subPreencheCampos()
   
  Dim XLT_SQL As String
  Dim XLO_ATRIBUICAO As New ADODB.Recordset 'Pega os dados do Imóvel
    
  SubAbrirRecordset XLO_ATRIBUICAO, "SELECT * FROM ConsSASAtribuicoes WHERE " & _
    " proj_cd_Projeto = " & XFI_PROJETO & _
      " AND subp_cd_SubProjeto = " & XFI_SUBPROJETO & _
      " AND tare_cd_UID = " & XFL_TAREFA & _
      " AND recu_cd_UID = " & XFL_RECURSO & _
      " AND atri_cd_UIDAtribuicao = " & XFL_ATRIBUICAO & _
      " AND atri_dt_Dia = " & FunNuloData(XFD_DIA, NomeSgbd), adOpenStatic
  
    TxtProjeto.Text = XLO_ATRIBUICAO!proj_tx_Nome
    TxtSubprojeto.Text = FunNuloString(XLO_ATRIBUICAO!Subprojeto)
    TxtTarefas.Text = XLO_ATRIBUICAO!tare_tx_Nome
    TxtRecurso.Text = XLO_ATRIBUICAO!recu_tx_Nome
    TxtDia.Text = Format(XLO_ATRIBUICAO!atri_dt_dia, "dd/mm/yy")
    TxtHHPrevisto.Text = Format(XLO_ATRIBUICAO!atri_vl_HHPrevisto, "standard")
    TxtHHReal.Text = Format(XLO_ATRIBUICAO!atri_vl_HHReal, "standard")
    
  XLO_ATRIBUICAO.Close
  Set XLO_ATRIBUICAO = Nothing
    
End Sub

Private Sub CmdDesistir_Click()
    Unload Me
End Sub

Private Sub cmdGravar_Click()
      
  Dim XLT_SQL As String
  
  'On Error GoTo RotuloErro
    
  XLT_SQL = "UPDATE Atribuicoes SET atri_vl_HHReal = " & FunTrataFloat(TxtHHReal.Text) & " WHERE " & _
    " proj_cd_Projeto = " & XFI_PROJETO & _
    " AND subp_cd_SubProjeto = " & XFI_SUBPROJETO & _
    " AND tare_cd_UID = " & XFL_TAREFA & _
    " AND recu_cd_UID = " & XFL_RECURSO & _
    " AND atri_cd_UIDAtribuicao = " & XFL_ATRIBUICAO & _
    " AND atri_dt_Dia = " & FunNuloData(XFD_DIA, NomeSgbd)
                     
    Conexao.Execute XLT_SQL
    
    Unload Me
    
  Exit Sub
    
'RotuloErro:
   ' funTrataErros (ComMensagem)
   ' Resume Next
   
End Sub

Private Sub Form_Activate()
    Set Formulario = Me
End Sub

Private Sub Form_Load()
    
    Dim XLT_SQL  As String
      
    Set Formulario = Me
    
    XFI_PROJETO = funDecodificarChave(Chave, Chr(10), 0)
    XFI_SUBPROJETO = funDecodificarChave(Chave, Chr(10), 1)
    XFL_TAREFA = funDecodificarChave(Chave, Chr(10), 2)
    XFL_RECURSO = funDecodificarChave(Chave, Chr(10), 3)
    XFL_ATRIBUICAO = funDecodificarChave(Chave, Chr(10), 4)
    XFD_DIA = funDecodificarChave(Chave, Chr(10), 5)
    
    subPreencheCampos
    
End Sub

Private Sub TxtHHReal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then KeyAscii = 44
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 44 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtHHReal_LostFocus()
    If IsNumeric(TxtHHReal.Text) Then
        TxtHHReal.Text = Format$(TxtHHReal.Text, "##,##0.00")
    Else
        TxtHHReal.Text = ""
    End If
End Sub
