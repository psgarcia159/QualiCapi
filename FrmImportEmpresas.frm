VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form FrmImportEmpresas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Importação Entre Empresas"
   ClientHeight    =   2490
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7425
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   166
   ScaleMode       =   0  'User
   ScaleWidth      =   495
   StartUpPosition =   1  'CenterOwner
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
      Height          =   255
      Left            =   5280
      TabIndex        =   12
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton CmdSair 
      Caption         =   "Sair"
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
      Left            =   6480
      TabIndex        =   11
      Top             =   2160
      Width           =   855
   End
   Begin MSMask.MaskEdBox MskConta 
      Height          =   330
      Left            =   1410
      TabIndex        =   0
      Top             =   960
      Width           =   840
      _ExtentX        =   1482
      _ExtentY        =   582
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   7
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MskPcr 
      Height          =   330
      Left            =   1410
      TabIndex        =   1
      Top             =   1350
      Width           =   840
      _ExtentX        =   1482
      _ExtentY        =   582
      _Version        =   393216
      MaxLength       =   20
      PromptChar      =   " "
   End
   Begin MSDataListLib.DataCombo CboPcr 
      Bindings        =   "FrmImportEmpresas.frx":0000
      Height          =   330
      Left            =   2250
      TabIndex        =   2
      Top             =   1350
      Width           =   5115
      _ExtentX        =   9022
      _ExtentY        =   582
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      ListField       =   "desccombo"
      BoundColumn     =   "cere_cd_estruturado"
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo CboPlanoContas 
      Bindings        =   "FrmImportEmpresas.frx":0015
      Height          =   330
      Left            =   2250
      TabIndex        =   3
      Top             =   960
      Width           =   5115
      _ExtentX        =   9022
      _ExtentY        =   582
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      ListField       =   "desccombo"
      BoundColumn     =   "plco_cd_codreduzido"
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo CboCCorrente 
      Bindings        =   "FrmImportEmpresas.frx":002C
      Height          =   330
      Left            =   1395
      TabIndex        =   4
      Top             =   1740
      Width           =   5985
      _ExtentX        =   10557
      _ExtentY        =   582
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      ListField       =   "desccombo"
      BoundColumn     =   "coco_cd_codigo"
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc DatContaCorrente 
      Height          =   330
      Left            =   2280
      Top             =   2160
      Visible         =   0   'False
      Width           =   2025
      _ExtentX        =   3572
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "DatContaCorrente"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc DatConta 
      Height          =   330
      Left            =   1080
      Top             =   2160
      Visible         =   0   'False
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "DatConta"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc DatPcr 
      Height          =   330
      Left            =   240
      Top             =   2160
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "DatPcr"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataListLib.DataCombo CboEmpreendimentos 
      Bindings        =   "FrmImportEmpresas.frx":004B
      Height          =   315
      Left            =   2040
      TabIndex        =   8
      Top             =   240
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      ListField       =   "empd_tx_Nome"
      BoundColumn     =   "empd_tx_Nome"
      Text            =   ""
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   240
      TabIndex        =   9
      Top             =   0
      Width           =   7095
      Begin VB.Label LblEmpreendimento 
         Alignment       =   1  'Right Justify
         Caption         =   "Empreendimento:"
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
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1485
      End
   End
   Begin MSAdodcLib.Adodc DatEmpreendimentos 
      Height          =   330
      Left            =   2880
      Top             =   2160
      Visible         =   0   'False
      Width           =   2280
      _ExtentX        =   4022
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "DatEmpreendimento"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog DlgAbrirFigura 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label LblPcr 
      Alignment       =   1  'Right Justify
      Caption         =   "CR Destino:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   180
      TabIndex        =   7
      Top             =   1410
      Width           =   1200
   End
   Begin VB.Label LblConta 
      Alignment       =   1  'Right Justify
      Caption         =   "Conta Destino:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   60
      TabIndex        =   6
      Top             =   1020
      Width           =   1320
   End
   Begin VB.Label LblCCorrente 
      Alignment       =   1  'Right Justify
      Caption         =   "C/C Destino:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   1800
      Width           =   1365
   End
End
Attribute VB_Name = "FrmImportEmpresas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CboPcr_Change()

If CboPcr.Text <> "" Then
    DatPcr.Recordset.bookmark = CboPcr.SelectedItem
    If DatPcr.Recordset.Fields("cere_tx_tipo") <> "A" Then
      MsgBox "Este item do Centro de Custo não é analítico, logo não pode ser atribuído a uma Nota.", vbCritical + vbOKOnly, "ATENÇÃO"
      CboPcr.BoundText = MskPcr.Text
      CboPcr.SetFocus
      Exit Sub
    End If
    MskPcr.Text = CboPcr.BoundText
End If
  
End Sub

Private Sub CboPlanoContas_Change()
If CboPlanoContas.Text <> "" Then
    DatConta.Recordset.bookmark = CboPlanoContas.SelectedItem
    If DatConta.Recordset.Fields("plco_tx_tipo") <> "A" Then
      MsgBox "Este item do Plano de Contas não é analítico, logo não pode ser atribuído a uma Nota.", vbCritical + vbOKOnly, "ATENÇÃO"
      CboPlanoContas.BoundText = MskConta.Text
      CboPlanoContas.SetFocus
      Exit Sub
    End If
    MskConta.Text = CboPlanoContas.BoundText
  End If
End Sub

Private Sub CmdConfirmar_Click()
    
    Dim XLT_CONTACONTABIL As String
    Dim XLT_PCR As String
    Dim XLT_CONTACORRENTE As String
    Dim XLT_CODEMPREEND As String
    Dim XLT_CODMOEDA As String
    Dim XLT_FORNCLI As String
    Dim XLO_EXPORTA As Object
    Dim XLT_SQL As String
    Dim XLT_TIPOIMOVEL As Integer
    Dim XLO_CODIGOS As ADODB.Recordset
    Dim XLT_PROFTITULAR As Integer
    Dim XLT_PROFCONJUGUE As Integer
    Dim XLT_OBSERVACAO As Integer
    Dim XLT_TIPOPLANO As Integer
    Dim XLT_TEXTO As String
    Dim XLT_CONTRATO As String
    Dim XLT_IMOVEIS As String
    
    Set Formulario = Me
    
    DlgAbrirFigura.CancelError = False
    DlgAbrirFigura.FileName = DlgAbrirFigura.Filter
    DlgAbrirFigura.ShowOpen
    DlgAbrirFigura.DialogTitle = "Importação de Títulos entre Empresas - IMPORTANDO MOEDAS"
    
    If DlgAbrirFigura.FileName = "" Then
        Unload Me
        Exit Sub
    Else
        XGT_LOCALARQ = DlgAbrirFigura.FileName
        XGT_ARQUIVO = DlgAbrirFigura.FileTitle
    End If

    Open XGT_LOCALARQ For Input Access Read As #1
    
    Me.MousePointer = vbHourglass
    
    DatConta.Recordset.bookmark = CboPlanoContas.SelectedItem
    XLT_CONTACONTABIL = DatConta.Recordset.Fields!plco_cd_Conta
    
    DatPcr.Recordset.bookmark = CboPcr.SelectedItem
    XLT_PCR = DatPcr.Recordset.Fields!cere_cd_Pcr
     
    DatContaCorrente.Recordset.bookmark = CboCCorrente.SelectedItem
    XLT_CONTACORRENTE = DatContaCorrente.Recordset.Fields!coco_cd_codigo
    
    DatEmpreendimentos.Recordset.bookmark = CboEmpreendimentos.SelectedItem
    XLT_CODEMPREEND = DatEmpreendimentos.Recordset.Fields!empd_cd_Empreendimento
    
    ' On Error GoTo TrataErro
     
    Conexao.BeginTrans
    
    Input #1, XLT_TEXTO 'Leitura da moeda
        
        'Importando Imóveis
        XLT_SQL = "SELECT moed_cd_CodMoeda FROM Moedas WHERE moed_tx_Descricao = '" & Trim(Mid(XLT_TEXTO, 7, 30)) & "'"
        SubQOpenRecordset XLO_EXPORTA, XLT_SQL, Estatico
                      
        If Not XLO_EXPORTA.EOF Then
                
            XLT_CODMOEDA = XLO_EXPORTA!moed_cd_CodMoeda
             
         Else
            XLO_EXPORTA.Close ' Inserido em 13/10/08 - Patrícia
            XLT_SQL = "INSERT INTO Moedas(moed_tx_simbolo, moed_tx_Descricao," & _
                        " moed_tx_Indexacao,moed_tx_padrao, codvelho)" & _
                        " VALUES ('" & Trim(Mid(XLT_TEXTO, 2, 5)) & "','" & _
                        "" & Trim(Mid(XLT_TEXTO, 7, 30)) & "','" & _
                        "" & Trim(Mid(XLT_TEXTO, 37, 1)) & "','" & _
                        "" & Trim(Mid(XLT_TEXTO, 38, 1)) & "','" & _
                        "" & Trim(Mid(XLT_TEXTO, 39, 5)) & "') "
            Conexao.Execute XLT_SQL
                    
            XLT_SQL = "SELECT max(moed_cd_codmoeda) as moed_cd_CodMoeda FROM Moedas"
            SubQOpenRecordset XLO_EXPORTA, XLT_SQL, Estatico
        
            'Pegando o último código da moeda
            If Not XLO_EXPORTA.EOF Then
                XLT_CODMOEDA = XLO_EXPORTA!moed_cd_CodMoeda
            End If
                    
            XLO_EXPORTA.Close
                 
            
            'Inserindo as cotações
            While Not EOF(1)
                Input #1, XLT_TEXTO
                   Dim XLT_TAXA As String
                        
                        'If Trim(Mid(XLT_TEXTO, 12, 15)) < 0 Then
                        '    XLT_TAXA = Trim(Mid(XLT_TEXTO, 12, 1)) + "0" + Trim(Mid(XLT_TEXTO, 12, 15))
                        'Else
                    If Trim(Mid(XLT_TEXTO, 12, 15)) <> "" And Mid(XLT_TEXTO, 12, 1) = " " Then
                        XLT_TAXA = "0" + Trim(Mid(XLT_TEXTO, 12, 15))
                    ElseIf Trim(Mid(XLT_TEXTO, 12, 15)) <> "" And Mid(XLT_TEXTO, 12, 1) = "-" Then
                        XLT_TAXA = Mid(XLT_TEXTO, 12, 1) + "0" + Trim(Mid(XLT_TEXTO, 13, 15))
                    ElseIf Trim(Mid(XLT_TEXTO, 12, 15)) = 0 Then
                        XLT_TAXA = "Null"
                    Else
                        XLT_TAXA = Trim(Mid(XLT_TEXTO, 12, 15))
                    End If
                    
                  XLT_SQL = "INSERT INTO CotacoesMoedas(cota_dt_DataCotacao, moed_cd_codmoeda," & _
                             " cota_vl_TaxaCotacao, cota_vl_Valor)" & _
                             " VALUES (" & FunNuloData(Mid(XLT_TEXTO, 2, 10), NomeSgbd) & "," & _
                             "" & XLT_CODMOEDA & "," & XLT_TAXA & "," & _
                             "" & Trim(Mid(XLT_TEXTO, 27, 15)) & ") "
                  Conexao.Execute XLT_SQL
                
            Wend
         
         End If
         
       Close #1
    
    DlgAbrirFigura.CancelError = False
    DlgAbrirFigura.FileName = DlgAbrirFigura.Filter
    DlgAbrirFigura.ShowOpen
    DlgAbrirFigura.DialogTitle = "Importação de Títulos entre Empresas - IMPORTANDO TÍTULOS"
        
    If DlgAbrirFigura.FileName = "" Then
        Unload Me
        Exit Sub
    Else
        XGT_LOCALARQ = DlgAbrirFigura.FileName
        XGT_ARQUIVO = DlgAbrirFigura.FileTitle
    End If

   Open XGT_LOCALARQ For Input Access Read As #1
    
    Me.MousePointer = vbHourglass
    
    While Not EOF(1)
    
        Input #1, XLT_TEXTO
        
        Select Case Mid(XLT_TEXTO, 1, 1)
        
            Case "I" 'Leitura dos Imóveis
                
                XLT_SQL = "SELECT imov_cd_Imovel, cont_cd_contrato FROM Contratos WHERE imov_cd_Imovel = '" & Trim(Mid(XLT_TEXTO, 2, 4)) & "'" & _
                           " AND empd_cd_Empreendimento = '" & XLT_CODEMPREEND & "'" & _
                           " AND empr_cd_Empresa=" & PCodEmpresa
                           
                SubQOpenRecordset XLO_EXPORTA, XLT_SQL, Estatico
                
                If Not XLO_EXPORTA.EOF Then
                    
                    XLT_IMOVEIS = XLO_EXPORTA!imov_cd_Imovel
                    
                    XLT_CONTRATO = Format(CInt(XLO_EXPORTA!cont_cd_Contrato) + 1, "00")
                    
                    XLO_EXPORTA.Close
                
                Else
                   XLO_EXPORTA.Close
                   
                   XLT_IMOVEIS = Trim(Mid(XLT_TEXTO, 2, 4))
                   
                   XLT_CONTRATO = ""
                   
                   XLT_SQL = "SELECT * FROM TiposImoveis WHERE tiim_tx_Descricao = '" & Trim(Mid(XLT_TEXTO, 6, 20)) & "'"
                   SubQOpenRecordset XLO_EXPORTA, XLT_SQL, Estatico
            
                   If Not XLO_EXPORTA.EOF Then
                          XLT_TIPOIMOVEL = XLO_EXPORTA!tiim_cd_TipoImovel
                          XLO_EXPORTA.Close
                   Else
                       XLO_EXPORTA.Close
                       XLT_SQL = "SELECT * FROM TiposImoveis WHERE tiim_cd_TipoImovel  = -1 "
                       SubQOpenRecordset XLO_CODIGOS, XLT_SQL, Dinamico
                       
                           XLO_CODIGOS.AddNew
                           XLO_CODIGOS!tiim_tx_Descricao = Trim(Mid(XLT_TEXTO, 6, 20))
                           XLO_CODIGOS.Update
                           
                           XLT_TIPOIMOVEL = XLO_EXPORTA!tiim_cd_TipoImovel
                           XLO_CODIGOS.Close
                   End If
                   
                   XLT_SQL = "INSERT INTO Imoveis(imov_cd_Imovel, empd_cd_Empreendimento, empr_cd_Empresa," & _
                            " tiim_cd_TipoImovel, imov_nr_Area, imov_tx_Garagem," & _
                            " imov_tx_Status, imov_vl_AVista, imov_nr_Andar, imov_dt_EntregaChaves, " & _
                            " imov_cd_UltimoContrato) VALUES ('" & XLT_IMOVEIS & "','" & _
                            "" & XLT_CODEMPREEND & "'," & PCodEmpresa & "," & XLT_TIPOIMOVEL & ",'" & _
                            "" & Trim(Mid(XLT_TEXTO, 26, 8)) & "','" & Trim(Mid(XLT_TEXTO, 34, 50)) & "','" & _
                            "" & Trim(Mid(XLT_TEXTO, 84, 1)) & "','" & (Trim(Mid(XLT_TEXTO, 85, 10))) & "','" & _
                            "" & Trim(Mid(XLT_TEXTO, 95, 2)) & "'," & FunNuloData(Trim(Mid(XLT_TEXTO, 97, 10)), NomeSgbd) & ",'" & Trim(Mid(XLT_TEXTO, 107, 2)) & "') "
                   Conexao.Execute XLT_SQL
                   
                   XLT_SQL = "UPDATE Empreendimentos SET " & _
                            " empd_nr_UnVendidas=empd_nr_UnVendidas +1 " & _
                            " WHERE empr_cd_Empresa=" & PCodEmpresa & _
                             " AND empd_cd_Empreendimento ='" & XLT_CODEMPREEND & "'"
                   Conexao.Execute XLT_SQL
            End If
            
            Case "F"  'Leitura do Forn Clientes
                
                XLT_SQL = "SELECT * FROM FornClientes WHERE focl_tx_RazaoSocial LIKE '" & Trim(Mid(XLT_TEXTO, 24, 50)) & "'"
                SubQOpenRecordset XLO_EXPORTA, XLT_SQL, Estatico
                
                Dim XLB_TESTE As Boolean
                
                If Not XLO_EXPORTA.EOF Then
                    XLT_FORNCLI = XLO_EXPORTA!focl_cd_FornCli
                    XLO_EXPORTA.Close
                    XLB_TESTE = False
                Else
                    XLB_TESTE = True
                    
                    XLT_SQL = "INSERT INTO FornClientes(focl_tx_Classe, focl_tx_Tipo," & _
                                " focl_tx_Fantasia, focl_tx_RazaoSocial, focl_tx_Endereco," & _
                                " focl_tx_Bairro, focl_tx_Cidade, focl_tx_EMail, focl_tx_Estado, " & _
                                " focl_tx_Cep, focl_tx_Fone, focl_tx_Fax, focl_tx_CgcCpf, focl_tx_IEstadual, " & _
                                " focl_tx_ContatoCom, focl_tx_ContatoFin, focl_tx_TipoPag, focl_vl_Icms," & _
                                " focl_tx_Fone1,focl_tx_Celular, focl_nr_Porta, focl_nr_CodAtividade, " & _
                                " focl_cd_CodDos, CodVelho1, CodVelho2, focl_tx_IMunicipal" & _
                                " ) VALUES ('" & Trim(Mid(XLT_TEXTO, 2, 1)) & "','" & _
                                "" & Trim(Mid(XLT_TEXTO, 3, 1)) & "','" & _
                                "" & Trim(Mid(XLT_TEXTO, 4, 20)) & "','" & Trim(Mid(XLT_TEXTO, 24, 50)) & "','" & _
                                "" & Trim(Mid(XLT_TEXTO, 74, 50)) & "','" & Trim(Mid(XLT_TEXTO, 124, 20)) & "','" & _
                                "" & Trim(Mid(XLT_TEXTO, 144, 20)) & "','" & Trim(Mid(XLT_TEXTO, 164, 30)) & "','" & Trim(Mid(XLT_TEXTO, 194, 2)) & "','" & _
                                "" & Trim(Mid(XLT_TEXTO, 196, 10)) & "','" & Trim(Mid(XLT_TEXTO, 206, 15)) & "','" & _
                                "" & Trim(Mid(XLT_TEXTO, 221, 15)) & "','" & Trim(Mid(XLT_TEXTO, 236, 14)) & "','" & _
                                "" & Trim(Mid(XLT_TEXTO, 250, 15)) & "','" & Trim(Mid(XLT_TEXTO, 265, 20)) & "','" & _
                                "" & Trim(Mid(XLT_TEXTO, 285, 20)) & "','" & _
                                "" & Trim(Mid(XLT_TEXTO, 305, 2)) & "','" & Trim(Mid(XLT_TEXTO, 307, 8)) & "','" & _
                                "" & Trim(Mid(XLT_TEXTO, 315, 15)) & "','" & Trim(Mid(XLT_TEXTO, 330, 15)) & "','" & _
                                "" & Trim(Mid(XLT_TEXTO, 345, 5)) & "','" & Trim(Mid(XLT_TEXTO, 350, 5)) & "','" & _
                                "" & Trim(Mid(XLT_TEXTO, 360, 5)) & "','" & Trim(Mid(XLT_TEXTO, 365, 20)) & "','" & _
                                "" & Trim(Mid(XLT_TEXTO, 385, 20)) & "','" & Trim(Mid(XLT_TEXTO, 405, 20)) & "') "
                    Conexao.Execute XLT_SQL
                        
                    XLO_EXPORTA.Close ' Inserido em 13/10/08 - Patrícia
                    XLT_SQL = "SELECT max(focl_cd_FornCli) as focl_cd_FornCli FROM FornClientes"
                    SubQOpenRecordset XLO_EXPORTA, XLT_SQL, Estatico
                     
                    If Not XLO_EXPORTA.EOF Then
                        XLT_FORNCLI = XLO_EXPORTA!focl_cd_FornCli
                    End If
                    XLO_EXPORTA.Close ' Inserido em 13/10/08 - Patrícia
                    
                End If
                
            Case "C" 'Leitura do Clientes
                
                If XLB_TESTE = True Then
                        
                    'Profissão de Titular
                    If Trim(Mid(XLT_TEXTO, 2, 20)) <> "" Then
                       XLT_SQL = "SELECT * FROM Profissoes WHERE prof_tx_Descricao = '" & Trim(Mid(XLT_TEXTO, 2, 20)) & "'"
                       SubQOpenRecordset XLO_EXPORTA, XLT_SQL, Estatico
                
                       If Not XLO_EXPORTA.EOF Then
                              XLT_PROFTITULAR = XLO_EXPORTA!prof_cd_Profissao
                              XLO_EXPORTA.Close
                       Else
                           XLO_EXPORTA.Close
                           XLT_SQL = "SELECT * FROM Profissoes WHERE prof_cd_Profissao  = -1 "
                           SubQOpenRecordset XLO_CODIGOS, XLT_SQL, Dinamico
                           
                               XLO_CODIGOS.AddNew
                               XLO_CODIGOS!prof_tx_Descricao = Trim(Mid(XLT_TEXTO, 2, 20))
                               XLO_CODIGOS.Update
                               
                               XLT_PROFTITULAR = XLO_CODIGOS!prof_cd_Profissao
                               XLO_CODIGOS.Close
                       End If
                     Else
                        XLT_PROFTITULAR = 0
                     End If
                       
                    'Profissão do Conjugue
                    If Trim(Mid(XLT_TEXTO, 218, 20)) <> "" Then
                       XLT_SQL = "SELECT * FROM Profissoes WHERE prof_tx_Descricao = '" & Trim(Mid(XLT_TEXTO, 218, 20)) & "'"
                       SubQOpenRecordset XLO_EXPORTA, XLT_SQL, Estatico
                
                       If Not XLO_EXPORTA.EOF Then
                              XLT_PROFCONJUGUE = XLO_EXPORTA!prof_cd_Profissao
                              XLO_EXPORTA.Close
                       Else
                           XLO_EXPORTA.Close
                           XLT_SQL = "SELECT * FROM Profissoes WHERE prof_cd_Profissao  = -1 "
                           SubQOpenRecordset XLO_CODIGOS, XLT_SQL, Dinamico
                           
                               XLO_CODIGOS.AddNew
                               XLO_CODIGOS!prof_tx_Descricao = Trim(Mid(XLT_TEXTO, 218, 20))
                               XLO_CODIGOS.Update
                               
                               XLT_PROFCONJUGUE = XLO_CODIGOS!prof_cd_Profissao
                               XLO_CODIGOS.Close
                       End If
                    Else
                        XLT_PROFCONJUGUE = 0
                    End If
                    
                    XLT_SQL = "INSERT INTO Clientes(focl_cd_FornCli, prof_cd_ProfTitular, clie_dt_Nascimento," & _
                                " clie_tx_Nacionalidade, clie_nr_Identidade, clie_tx_OrgEmissor," & _
                                " clie_nr_EstCivil, clie_tx_EndResidencial, clie_tx_BairroResidencial, clie_nr_CepResidencial, " & _
                                " clie_tx_MunResidencial, clie_tx_estresidencial, clie_tx_Conjuge, prof_cd_ProfConjuge, clie_dt_NascConjuge, " & _
                                " clie_tx_NacConjuge, clie_nr_IdConjuge, clie_tx_OrgEmissorConjuge, clie_nr_CpfConjuge," & _
                                " clie_vl_RendaFamiliar,clie_tx_EndCorresp, clie_tx_BairroCorresp, clie_nr_CepCorresp, " & _
                                " clie_tx_MunCorresp, clie_tx_EstCorresp) VALUES (" & XLT_FORNCLI & "," & _
                                "" & FunNuloBanco(XLT_PROFTITULAR) & "," & FunNuloData(Trim(Mid(XLT_TEXTO, 22, 10)), NomeSgbd) & ",'" & Trim(Mid(XLT_TEXTO, 32, 10)) & "','" & _
                                "" & Trim(Mid(XLT_TEXTO, 42, 15)) & "','" & Trim(Mid(XLT_TEXTO, 57, 8)) & "','" & _
                                "" & Trim(Mid(XLT_TEXTO, 65, 1)) & "','" & Trim(Mid(XLT_TEXTO, 66, 50)) & "','" & _
                                "" & Trim(Mid(XLT_TEXTO, 116, 20)) & "','" & Trim(Mid(XLT_TEXTO, 136, 10)) & "','" & Trim(Mid(XLT_TEXTO, 146, 20)) & "','" & _
                                "" & Trim(Mid(XLT_TEXTO, 166, 2)) & "','" & Trim(Mid(XLT_TEXTO, 168, 50)) & "'," & _
                                "" & FunNuloBanco(XLT_PROFCONJUGUE) & "," & _
                                "" & FunNuloData(Trim(Mid(XLT_TEXTO, 238, 10)), NomeSgbd) & ",'" & Trim(Mid(XLT_TEXTO, 248, 10)) & "','" & _
                                "" & Trim(Mid(XLT_TEXTO, 258, 15)) & "','" & Trim(Mid(XLT_TEXTO, 273, 8)) & "','" & _
                                "" & (Trim(Mid(XLT_TEXTO, 281, 14))) & "','" & Trim(Mid(XLT_TEXTO, 295, 10)) & "','" & _
                                "" & Trim(Mid(XLT_TEXTO, 305, 50)) & "','" & Trim(Mid(XLT_TEXTO, 355, 20)) & "','" & _
                                "" & Trim(Mid(XLT_TEXTO, 375, 10)) & "','" & Trim(Mid(XLT_TEXTO, 385, 20)) & "','" & Trim(Mid(XLT_TEXTO, 405, 2)) & "') "
                    Conexao.Execute XLT_SQL
                End If
            Case "N" ' 'Leitura dos Contratos
                
                If XLT_CONTRATO = "" Then
                    XLT_CONTRATO = Trim(Mid(XLT_TEXTO, 2, 2))
                End If
                
                XLT_SQL = "INSERT INTO Contratos (cont_cd_Contrato,imov_cd_imovel, empd_cd_empreendimento, " & _
                        "empr_cd_Empresa, cere_cd_Pcr, plco_cd_conta, moed_cd_Codmoeda, " & _
                        "cont_vl_Avista, cont_vl_sinal, cont_vl_Poupanca, cont_vl_Total, " & _
                        "cont_vl_Fgts, cont_vl_Financiamento, cont_vl_Outros, cont_dt_base, " & _
                        "cont_dt_RegistroContrato, cont_dt_EntregaChaves, cont_vl_TaxaSegVida, " & _
                        "cont_dt_Distrato,focl_cd_FornCli, cont_tx_MultaProRata, cont_tx_CorrecProRata, " & _
                        "cont_nr_TipoSeguro, cont_tx_TipoJurosMora, cont_tx_Desagio, cont_vl_PercJurosFin, " & _
                        "cont_vl_PercJurosMulta, cont_vl_PercJurosMora, cont_tx_Status, cont_cd_ContratoOrigem, " & _
                        "cont_dt_Venda, cont_dt_Quitacao, cont_dt_LimiteReserva, coco_cd_Codigo, cont_nr_LimiteDias)" & _
                        " VALUES ('" & XLT_CONTRATO & "','" & XLT_IMOVEIS & "','" & XLT_CODEMPREEND & "','" & PCodEmpresa & "','" & _
                            "" & XLT_PCR & "','" & XLT_CONTACONTABIL & "'," & XLT_CODMOEDA & "," & _
                            "" & Trim(Mid(XLT_TEXTO, 8, 10)) & "," & Trim(Mid(XLT_TEXTO, 18, 10)) & "," & _
                            "" & Trim(Mid(XLT_TEXTO, 28, 10)) & "," & Trim(Mid(XLT_TEXTO, 38, 10)) & "," & _
                            "" & Trim(Mid(XLT_TEXTO, 48, 10)) & "," & Trim(Mid(XLT_TEXTO, 58, 10)) & "," & _
                            "" & Trim(Mid(XLT_TEXTO, 68, 10)) & "," & FunNuloData(Trim(Mid(XLT_TEXTO, 78, 10)), NomeSgbd) & "," & _
                            "" & FunNuloData(Trim(Mid(XLT_TEXTO, 88, 10)), NomeSgbd) & "," & FunNuloData(Trim(Mid(XLT_TEXTO, 98, 10)), NomeSgbd) & "," & _
                            "" & Trim(Mid(XLT_TEXTO, 108, 10)) & "," & FunNuloData(Trim(Mid(XLT_TEXTO, 118, 10)), NomeSgbd) & "," & _
                            "" & XLT_FORNCLI & ",'" & Trim(Mid(XLT_TEXTO, 128, 1)) & "','" & _
                            "" & Trim(Mid(XLT_TEXTO, 129, 1)) & "','" & Trim(Mid(XLT_TEXTO, 130, 1)) & "','" & _
                            "" & Trim(Mid(XLT_TEXTO, 131, 1)) & "','" & Trim(Mid(XLT_TEXTO, 132, 1)) & "'," & _
                            "" & Trim(Mid(XLT_TEXTO, 133, 10)) & "," & Trim(Mid(XLT_TEXTO, 143, 10)) & "," & _
                            "" & Trim(Mid(XLT_TEXTO, 153, 10)) & ",'" & Trim(Mid(XLT_TEXTO, 163, 1)) & "','" & _
                            "" & Trim(Mid(XLT_TEXTO, 164, 2)) & "'," & FunNuloData(Trim(Mid(XLT_TEXTO, 166, 10)), NomeSgbd) & "," & _
                            "" & FunNuloData(Trim(Mid(XLT_TEXTO, 176, 10)), NomeSgbd) & "," & FunNuloData(Trim(Mid(XLT_TEXTO, 186, 10)), NomeSgbd) & ",'" & XLT_CONTACORRENTE & "','" & _
                            "" & Trim(Mid(XLT_TEXTO, 196, 4)) & "') "
                Conexao.Execute XLT_SQL
                
            Case "T" 'Leitura dos Títulos
                     
                'Tipos de Plano
                 If Trim(Mid(XLT_TEXTO, 16, 20)) <> "" Then
                    XLT_SQL = "SELECT * FROM TiposPlanos WHERE tipl_tx_Descricao = '" & Trim(Mid(XLT_TEXTO, 16, 20)) & "'"
                    SubQOpenRecordset XLO_EXPORTA, XLT_SQL, Estatico
            
                    If Not XLO_EXPORTA.EOF Then
                        XLT_TIPOPLANO = XLO_EXPORTA!tipl_cd_TipoPlano
                        XLO_EXPORTA.Close
                    Else
                        XLO_EXPORTA.Close
                       
                        XLT_SQL = "SELECT * FROM TiposPlanos WHERE tipl_cd_TipoPlano  = -1 "
                        SubQOpenRecordset XLO_CODIGOS, XLT_SQL, Dinamico
                       
                        XLO_CODIGOS.AddNew
                        XLO_CODIGOS!tipl_cd_TipoPlano = Trim(Mid(XLT_TEXTO, 16, 20))
                        XLO_CODIGOS.Update
                           
                        XLT_TIPOPLANO = XLO_CODIGOS!tipl_cd_TipoPlano
                        XLO_CODIGOS.Close
                    End If
                  Else
                         XLT_TIPOPLANO = 0
                  End If
                  
                'Tipos de Observação
                 If Trim(Mid(XLT_TEXTO, 36, 50)) <> "" Then
                    XLT_SQL = "SELECT * FROM Observacoes  WHERE obse_tx_observacao = '" & Trim(Mid(XLT_TEXTO, 36, 50)) & "'"
                    SubQOpenRecordset XLO_EXPORTA, XLT_SQL, Estatico
            
                    If Not XLO_EXPORTA.EOF Then
                        XLT_OBSERVACAO = XLO_EXPORTA!obse_cd_Observacao
                        XLO_EXPORTA.Close
                    Else
                        XLO_EXPORTA.Close
                       
                        XLT_SQL = "SELECT * FROM Observacoes WHERE obse_cd_observacao  = -1 "
                        SubQOpenRecordset XLO_CODIGOS, XLT_SQL, Dinamico
                       
                        XLO_CODIGOS.AddNew
                        XLO_CODIGOS!obse_tx_Observacao = Trim(Mid(XLT_TEXTO, 36, 50))
                        XLO_CODIGOS.Update
                           
                        XLT_OBSERVACAO = XLO_CODIGOS!obse_cd_Observacao
                        XLO_CODIGOS.Close
                    End If
                  Else
                     XLT_OBSERVACAO = 0
                  End If
                    XLT_SQL = "INSERT INTO Titulos (titu_cd_Plano,titu_cd_Parcela, titu_cd_Residuo, " & _
                            "empr_cd_Empresa, empd_cd_Empreendimento, imov_cd_Imovel, cont_cd_Contrato, " & _
                            "tipl_cd_TipoPlano, coco_cd_Codigo, obse_cd_Observacao, moed_cd_Moeda1, " & _
                            "moed_cd_Moeda2, titu_vl_Parcela, titu_vl_ParcelaIndex1, titu_vl_corrigido1, " & _
                            "titu_vl_ParcelaIndex2, titu_vl_corrigido2, titu_dt_Vencimento, " & _
                            "titu_tx_IndicePagamento, titu_tx_Financia, titu_vl_PercJurosFin, titu_vl_PercMultaMora, " & _
                            "titu_vl_PercJurosMora, titu_dt_Repasse, titu_dt_Prorrogacao, titu_tx_TipoPag, " & _
                            "titu_dt_Base, cere_cd_Pcr, plco_cd_Conta, titu_tx_NaturezaPlano, titu_tx_CorrigeParcela, titu_tx_PreDatado, titu_dt_BasePagto, " & _
                            "titu_dt_Criacao, nofi_cd_NotaFiscal, titu_tx_ExpBanco, titu_tx_EntregaChave) " & _
                            "VALUES ('" & Trim(Mid(XLT_TEXTO, 2, 2)) & "','" & Trim(Mid(XLT_TEXTO, 4, 3)) & "','" & Trim(Mid(XLT_TEXTO, 7, 3)) & "','" & PCodEmpresa & "','" & _
                                "" & XLT_CODEMPREEND & "','" & XLT_IMOVEIS & "','" & XLT_CONTRATO & "','" & _
                                "" & FunNuloBanco(XLT_TIPOPLANO) & "','" & XLT_CONTACORRENTE & "'," & FunNuloBanco(XLT_OBSERVACAO) & "," & _
                                "" & XLT_CODMOEDA & "," & XLT_CODMOEDA & "," & Trim(Mid(XLT_TEXTO, 86, 20)) & "," & _
                                "" & Trim(Mid(XLT_TEXTO, 106, 10)) & "," & _
                                "" & Trim(Mid(XLT_TEXTO, 116, 10)) & "," & Trim(Mid(XLT_TEXTO, 126, 10)) & "," & _
                                "" & Trim(Mid(XLT_TEXTO, 136, 10)) & "," & FunNuloData(Trim(Mid(XLT_TEXTO, 146, 10)), NomeSgbd) & "," & _
                                "" & FunNuloBanco(Trim(Mid(XLT_TEXTO, 156, 1))) & "," & FunNuloBanco(Trim(Mid(XLT_TEXTO, 157, 1))) & "," & Trim(Mid(XLT_TEXTO, 158, 10)) & "," & _
                                "" & Trim(Mid(XLT_TEXTO, 168, 10)) & "," & Trim(Mid(XLT_TEXTO, 178, 10)) & "," & _
                                "" & FunNuloData(Trim(Mid(XLT_TEXTO, 188, 10)), NomeSgbd) & "," & FunNuloData(Trim(Mid(XLT_TEXTO, 198, 10)), NomeSgbd) & ",'" & _
                                "" & Trim(Mid(XLT_TEXTO, 208, 2)) & "'," & FunNuloData(Trim(Mid(XLT_TEXTO, 210, 10)), NomeSgbd) & "," & XLT_PCR & "," & XLT_CONTACONTABIL & ",'" & _
                                "" & Trim(Mid(XLT_TEXTO, 220, 1)) & "','" & Trim(Mid(XLT_TEXTO, 221, 1)) & "','" & Trim(Mid(XLT_TEXTO, 222, 1)) & "'," & _
                                "" & FunNuloData(Trim(Mid(XLT_TEXTO, 223, 10)), NomeSgbd) & "," & FunNuloData(Trim(Mid(XLT_TEXTO, 233, 10)), NomeSgbd) & ",'" & _
                                "" & Trim(Mid(XLT_TEXTO, 243, 4)) & "','" & Trim(Mid(XLT_TEXTO, 247, 1)) & "','" & Trim(Mid(XLT_TEXTO, 248, 1)) & "') "
                    Conexao.Execute XLT_SQL
        End Select
    Wend
    Conexao.CommitTrans
    Close #1
    MsgBox "Importação realizada com sucesso.", vbInformation + vbOKOnly, "ATENÇÃO"
    Unload Me
End Sub

Private Sub CmdSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    
    subConectarControleDadosNV DatEmpreendimentos, "SELECT * FROM Empreendimentos WHERE empr_cd_Empresa = " & PCodEmpresa & " ORDER BY empd_tx_Nome", Estatico
    subConectarControleDadosNV DatPcr, "SELECT * FROM ConsGENCRcombo WHERE empr_cd_empresa = " & PCodEmpresa & " ORDER BY cere_cd_estruturado", Estatico
    subConectarControleDadosNV DatConta, "SELECT * FROM ConsGENPCcombo WHERE empr_cd_empresa = " & PCodEmpresa & " and (plco_tx_recdesp= 'R' or plco_tx_recdesp= 'G') ORDER BY plco_cd_estruturado", Estatico
    subConectarControleDadosNV DatContaCorrente, "SELECT * FROM ConsGENCCcombo where empr_cd_empresa=" & PCodEmpresa & " ORDER BY coco_tx_Descricao", Estatico

End Sub

Private Sub Mskconta_LostFocus()

  MskConta.PromptInclude = False
  If MskConta.Text <> "" Then
    
    If Mid(XGT_MSKCODREDUZIDO, 1, 1) = "0" Then
        MskConta.Text = Format(MskConta.Text, XGT_MSKCODREDUZIDO)
    End If
    MskConta.PromptInclude = True
    
    CboPlanoContas.BoundText = MskConta.Text
    If CboPlanoContas.Text = "" Then
      MsgBox "Não existe item do Plano de Contas com este código reduzido.", vbCritical + vbOKOnly, "ATENÇÃO"
      MskConta.SetFocus
      If MskConta.Enabled = True Then MskConta.SetFocus
    End If
  End If
End Sub

Private Sub MskPcr_LostFocus()
MskPcr.PromptInclude = False
    If MskPcr.Text <> "" Then
        
        If Len(MskPcr.Text) <> Len(pPCRmascara) And pPCRnivel = 1 Then
          MskPcr.Text = funFormataCCusto(MskPcr.Text)
        End If
        
        
        MskPcr.Mask = pPCRmascara
        MskPcr.PromptInclude = True
        CboPcr.BoundText = MskPcr.Text
        If CboPcr.Text = "" Then
            MsgBox "Não existe item do Centro de Custo com este código.", vbCritical + vbOKOnly, "ATENÇÃO"
            If MskPcr.Visible = True Then MskPcr.SetFocus
        End If
    End If
    MskPcr.PromptInclude = True
End Sub
