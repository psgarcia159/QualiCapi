VERSION 5.00
Object = "{6FBA474E-43AC-11CE-9A0E-00AA0062BB4C}#1.0#0"; "SYSINFO.OCX"
Begin VB.Form TelaInicial 
   Caption         =   "Pagamento Eletrônico - ITAÚ"
   ClientHeight    =   5370
   ClientLeft      =   765
   ClientTop       =   1800
   ClientWidth     =   7500
   Icon            =   "TelaInicial.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "TelaInicial.frx":08CA
   ScaleHeight     =   5370
   ScaleWidth      =   7500
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdDesistir 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6780
      Picture         =   "TelaInicial.frx":20A3A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4725
      Width           =   615
   End
   Begin VB.CommandButton CmdExporta 
      Height          =   840
      Left            =   2250
      Picture         =   "TelaInicial.frx":20E7C
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1470
      Width           =   1100
   End
   Begin VB.CommandButton CmdImporta 
      Height          =   840
      Left            =   2265
      Picture         =   "TelaInicial.frx":21746
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2790
      Width           =   1100
   End
   Begin SysInfoLib.SysInfo SysInfo1 
      Left            =   6825
      Top             =   1455
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Versão"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6720
      TabIndex        =   4
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label LblEmpresa 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Empresa"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   720
      TabIndex        =   3
      Top             =   5040
      Width           =   5895
   End
End
Attribute VB_Name = "TelaInicial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdDesistir_Click()
    Conexao.Close
    Unload Me
    End
End Sub

Private Sub CmdExporta_Click()
    XGT_TIPO = "E"
    TelaExpBancaria.Show
End Sub

Private Sub CmdImporta_Click()
    XGT_TIPO = "I"
    FrmImportacao.Show
End Sub

Private Sub Form_Activate()
    Set Formulario = TelaInicial
End Sub

Private Sub Form_Load()
    Dim XLO_VERIFICA As ADODB.Recordset
    Dim res As VbMsgBoxResult
    Dim Diretorio As String
    
    XGT_NomeAplic = "EXPITAU"
    XGT_IdAplic = "EXPITAU-UATIPXE"
    subTestaProtecao XGT_NomeAplic, XGT_IdAplic, XGT_DirPadrao
    XGT_NomeAplic = "EXPITAU"
    LblEmpresa.Caption = PFantasia
    TelaInicial.Label1.Caption = "R " & Format(App.Major, "#0") & "." & Format(App.Revision, "00")
    
    '------------------------------------------------------------------------
    'ATUALIZAÇÃO DO SISTEMA
    
    Diretorio = Mid(App.Path, 1, (Len(App.Path) - Len("\QualiCapi\ExpItau")))
    
    'Verifica a versão do cliente no banco de dados
    SubQOpenRecordset XLO_VERIFICA, "SELECT sist_vl_versao FROM Sistema WHERE sist_tx_nome = 'QUALICAPI - EXPITAU'", Estatico
        
    If Not XLO_VERIFICA.BOF Then
        'Se a versão atual for maior que a existente atualiza o banco
        If CDbl(Format(App.Major, "#0") & "," & Format(App.Revision, "00")) > CDbl(XLO_VERIFICA!sist_vl_versao) Then
            
            Conexao.Execute _
                "UPDATE Sistema SET sist_vl_versao = " & _
                Format(App.Major, "#0") & "." & Format(App.Revision, "00") & _
                " WHERE sist_tx_nome = 'QUALICAPI - EXPITAU'"
        
        ElseIf CDbl(Format(App.Major, "#0") & "," & Format(App.Revision, "00")) < CDbl(XLO_VERIFICA!sist_vl_versao) Then
            
            XLO_VERIFICA.Close
            Set XLO_VERIFICA = Nothing
            
            res = MsgBox("Máquina desatualizada! Deseja atualizar o ExpItau nesta máquina?", VbMsgBoxStyle.vbYesNo + vbQuestion, "Aviso")
            
            If (res = vbYes) Then
                Call Shell(Diretorio & "\Atualizador\Atualizador.exe QualiCapi ExpItau", vbNormalFocus)
                End
            Else
                MsgBox "A aplicação não será executada até ser atualizada", VbMsgBoxStyle.vbInformation, "Aviso"
                End
            End If
        End If
    Else
        XLO_VERIFICA.Close
        Set XLO_VERIFICA = Nothing
    End If
    
    'Verifica se há atualização para o Atualizador Qualipro:
    If (Dir(App.Path & "\Atualizador.ini") = "Atualizador.ini") Then
        If (Dir(Diretorio & "\Atualizador\Atualizador.ini") = "Atualizador.ini") Then
            Kill Diretorio & "\Atualizador\Atualizador.ini"
        End If
        
        'Copia o arquivo Atualizador.ini para a pasta do Atualizador
        Name App.Path & "\Atualizador.ini" As Diretorio & "\Atualizador\Atualizador.ini"
        
        'Exclui a versão antiga do atualizador
        Kill Diretorio & "\Atualizador\Atualizador.exe"
        
        'Muda a extenção do arquivo Atualizador.ini para .exe
        Name Diretorio & "\Atualizador\Atualizador.ini" As Diretorio & "\Atualizador\Atualizador.exe"
    
    ElseIf (Dir(Diretorio & "\Atualizador\Atualizador.ini") = "Atualizador.ini") Then
        Kill Diretorio & "\Atualizador\Atualizador.exe"
        Name Diretorio & "\Atualizador\Atualizador.ini" As Diretorio & "\Atualizador\Atualizador.exe"
    End If
    '------------------------------------------------------------------------

End Sub


Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub Label2_Click()

End Sub
