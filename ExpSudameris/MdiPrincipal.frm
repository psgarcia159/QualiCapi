VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{6FBA474E-43AC-11CE-9A0E-00AA0062BB4C}#1.0#0"; "SYSINFO.OCX"
Begin VB.MDIForm MdiPrincipal 
   AutoShowChildren=   0   'False
   BackColor       =   &H00C0C0C0&
   Caption         =   "Pagamento Eletrônico - Banco Itaú"
   ClientHeight    =   4440
   ClientLeft      =   870
   ClientTop       =   2895
   ClientWidth     =   10515
   Icon            =   "MdiPrincipal.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   NegotiateToolbars=   0   'False
   Picture         =   "MdiPrincipal.frx":08D2
   WindowState     =   2  'Maximized
   Begin ComctlLib.StatusBar BarraStatus 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   1
      Top             =   4110
      Width           =   10515
      _ExtentX        =   18547
      _ExtentY        =   529
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   6
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   10585
            MinWidth        =   10585
            Object.Tag             =   ""
            Object.ToolTipText     =   "Mensagens do Sistema"
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   3069
            MinWidth        =   3069
            Object.Tag             =   ""
            Object.ToolTipText     =   "Empresa Atual"
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Width           =   2822
            MinWidth        =   2822
            Object.Tag             =   ""
            Object.ToolTipText     =   "Informações sobre o Banco de Dados"
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   970
            MinWidth        =   970
            TextSave        =   "CAPS"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Caps Lock"
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   2
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   882
            MinWidth        =   882
            TextSave        =   "NUM"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Num Lock"
         EndProperty
         BeginProperty Panel6 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   3
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   882
            MinWidth        =   882
            TextSave        =   "INS"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Modo Insert"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   30
      Left            =   0
      TabIndex        =   0
      Top             =   4410
      Width           =   10515
      _ExtentX        =   18547
      _ExtentY        =   53
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CdbImpressora 
      Left            =   150
      Top             =   1950
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin SysInfoLib.SysInfo SysInfo1 
      Left            =   150
      Top             =   630
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin ComctlLib.ImageList ListaImagens 
      Left            =   165
      Top             =   1290
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   6
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MdiPrincipal.frx":59CAB
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MdiPrincipal.frx":59FC5
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MdiPrincipal.frx":5A2DF
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MdiPrincipal.frx":5A5F9
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MdiPrincipal.frx":5A70B
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MdiPrincipal.frx":5A81D
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "MdiPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub MDIForm_Activate()
    
    ' Carrega formulario para entrada da senha do usuário
'''    If func_cd_operador = 0 Then
'''       func_cd_operador = 9999
'''       FrmSenha.Show 1
'''    End If
    If XGT_TIPO = "E" Then
        TelaExpBancaria.Show
    Else
        FrmMovImportacao.Show
    End If

End Sub


Private Sub MDIForm_Load()
    ReDim VetorJanelas(0) As Variant
    
    'TestaProtecao
    XGT_NomeAplic = "ADMFIN"
    XGT_IdAplic = "QUALIADMFIN-NIFMDAILAUQ"
    subTestaProtecao XGT_NomeAplic, XGT_IdAplic, XGT_DirPadrao
    XGT_NomeAplic = "QUALIFIN"
    MdiPrincipal.BarraStatus.Panels(2).Text = PFantasia
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload Me
End Sub



Private Sub MnuFinalizar_Click()
    Unload Me
End Sub
