VERSION 5.00
Begin VB.Form FrmImpressora 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Impressoras Ativas"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5730
   Icon            =   "FrmImpressora.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   5730
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   60
      TabIndex        =   2
      Top             =   -60
      Width           =   5655
      Begin VB.ListBox LstImpressora 
         Height          =   2010
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   180
         Width           =   5415
      End
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3180
      TabIndex        =   1
      Top             =   2400
      Width           =   1155
   End
   Begin VB.CommandButton CmdConfirmar 
      Caption         =   "Confirmar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1020
      TabIndex        =   0
      Top             =   2400
      Width           =   1155
   End
End
Attribute VB_Name = "FrmImpressora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancelar_Click()
  
  Unload Me
  
End Sub

Private Sub CmdConfirmar_Click()
  
  Dim XLO_IMPRESSORA As Printer
  
  For Each XLO_IMPRESSORA In Printers
    If XLO_IMPRESSORA.DeviceName = LstImpressora.Text Then
      Set Printer = XLO_IMPRESSORA
      Exit For
    End If
  Next
  
  Unload Me
  
End Sub

Private Sub Form_Load()
  
  Dim XLO_IMPRESSORA As Printer
    
  For Each XLO_IMPRESSORA In Printers
    LstImpressora.AddItem XLO_IMPRESSORA.DeviceName
  Next
  
  If LstImpressora.ListCount > 0 Then
    LstImpressora.ListIndex = 0
  End If
  
End Sub
