VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form FrmSobre 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sobre QualiComp"
   ClientHeight    =   5460
   ClientLeft      =   2265
   ClientTop       =   1665
   ClientWidth     =   5595
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "System"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "FrmSobre.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5460
   ScaleWidth      =   5595
   Begin Threed.SSPanel SSPanel2 
      Height          =   120
      Left            =   225
      TabIndex        =   22
      Top             =   3180
      Width           =   5145
      _Version        =   65536
      _ExtentX        =   9075
      _ExtentY        =   212
      _StockProps     =   15
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "System"
         Size            =   9.77
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   0
      BevelInner      =   1
   End
   Begin VB.PictureBox VSElastic1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5460
      Left            =   0
      ScaleHeight     =   5400
      ScaleWidth      =   5535
      TabIndex        =   9
      Top             =   0
      Width           =   5595
      Begin Threed.SSPanel SSPanel1 
         Height          =   120
         Left            =   195
         TabIndex        =   21
         Top             =   1710
         Width           =   5145
         _Version        =   65536
         _ExtentX        =   9075
         _ExtentY        =   212
         _StockProps     =   15
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "System"
            Size            =   9.77
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         BevelInner      =   1
      End
      Begin VB.PictureBox VSElastic2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   240
         ScaleHeight     =   825
         ScaleWidth      =   5010
         TabIndex        =   5
         Top             =   2115
         Width           =   5070
         Begin VB.Label lblUsuario 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "QUALIDADOS - Consultoria e Sistemas Ltda."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Left            =   165
            TabIndex        =   20
            Top             =   60
            Width           =   4365
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Nº de Série:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   180
            TabIndex        =   6
            Top             =   315
            Width           =   1335
         End
         Begin VB.Label lblNroSerie 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "1111"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   1515
            TabIndex        =   7
            Top             =   330
            Width           =   1230
         End
         Begin VB.Label label2 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Licenças de rede:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   180
            TabIndex        =   8
            Top             =   540
            Width           =   1335
         End
         Begin VB.Label lblNroLic 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "9999"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   1500
            TabIndex        =   10
            Top             =   555
            Width           =   630
         End
      End
      Begin VB.CommandButton Cmd_OK 
         Appearance      =   0  'Flat
         Caption         =   "OK"
         Height          =   270
         Left            =   4320
         TabIndex        =   0
         Top             =   5040
         Width           =   1035
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "QUALIDADOS - Consultoria e Sistemas Ltda."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   675
         TabIndex        =   4
         Top             =   3630
         Width           =   4365
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         Caption         =   "Esta cópia é registrada para:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   210
         TabIndex        =   19
         Top             =   1845
         Width           =   3765
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Fone:  (71) 3339-2600"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   720
         TabIndex        =   18
         Top             =   5010
         Width           =   1860
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "CEP: 40140-340   -   Salvador - Bahia - Brasil"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   1530
         TabIndex        =   17
         Top             =   4770
         Width           =   3270
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Endereço:  Rua Ari Barroso, 12 - Chame-Chame"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   705
         TabIndex        =   16
         Top             =   4545
         Width           =   3915
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Fax:  (71) 3339-2643"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   705
         TabIndex        =   15
         Top             =   4335
         Width           =   2460
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "suporte@qualidados.com.br"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   255
         Left            =   1200
         TabIndex        =   14
         Top             =   4110
         Width           =   4110
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "EMail:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   690
         TabIndex        =   13
         Top             =   4110
         Width           =   555
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Maurício Netto Simões"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   690
         TabIndex        =   12
         Top             =   3855
         Width           =   3330
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Suporte Técnico:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   270
         Left            =   240
         TabIndex        =   11
         Top             =   3330
         Width           =   2010
      End
      Begin VB.Label Lbl_Microsoft 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   525
         Left            =   885
         TabIndex        =   1
         Top             =   735
         Width           =   3495
      End
      Begin VB.Label Lbl_Version 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Versão 1.13"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   1860
         TabIndex        =   2
         Top             =   1350
         Width           =   1470
      End
      Begin VB.Label Lbl_IconWorks 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "QualiCOMP"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   900
         TabIndex        =   3
         Top             =   15
         Width           =   3585
      End
   End
End
Attribute VB_Name = "FrmSobre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

DefInt A-Z
'Private Declare Function GetFreeSpace Lib "Kernel" (ByVal wFlags) As Long
'Private Declare Function GetWinFlags Lib "Kernel" () As Long
Const WF_STANDARD = &H10
Const WF_ENHANCED = &H20
Const WF_80x87 = &H400

Private Sub Cmd_OK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  Dim WinFlags As Long
  Dim Mode As String, processor As String
  FrmSobre.Caption = "Sobre " & App.ProductName
  Lbl_IconWorks.Caption = App.ProductName
  Lbl_Version.Caption = "Versão " & Format(App.Major, "#0") & "." & Format(App.Revision, "00")
  Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
End Sub

