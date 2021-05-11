VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRelAnexoFluxoRecebiveisReal 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consulta Anexo de Fluxo Recebido - Real"
   ClientHeight    =   3630
   ClientLeft      =   3330
   ClientTop       =   3615
   ClientWidth     =   6705
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   6705
   StartUpPosition =   2  'CenterScreen
   Begin MSDataListLib.DataCombo CboIndexador2 
      Bindings        =   "FrmRelAnexoFluxoRecebiveisReal.frx":0000
      Height          =   315
      Left            =   4560
      TabIndex        =   18
      Top             =   2600
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      ListField       =   "moed_tx_descricao"
      BoundColumn     =   ""
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo CboIndexador1 
      Bindings        =   "FrmRelAnexoFluxoRecebiveisReal.frx":0017
      Height          =   315
      Left            =   2400
      TabIndex        =   16
      Top             =   2600
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      ListField       =   "moed_tx_descricao"
      BoundColumn     =   ""
      Text            =   ""
   End
   Begin VB.Frame FraEmpreendimento 
      Caption         =   "Empreendimento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6255
      Begin MSMask.MaskEdBox MskEmpreendimentoFinal 
         Height          =   315
         Left            =   960
         TabIndex        =   6
         Top             =   720
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   4
         Mask            =   "AAAA"
         PromptChar      =   " "
      End
      Begin MSDataListLib.DataCombo CboEmpreendimentoFinal 
         Bindings        =   "FrmRelAnexoFluxoRecebiveisReal.frx":002E
         Height          =   315
         Left            =   1680
         TabIndex        =   7
         Top             =   720
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "empd_tx_nome"
         BoundColumn     =   "empd_cd_empreendimento"
         Text            =   "CboEmpreendimento"
      End
      Begin Threed.SSCommand cmdLimparEmpreendimentoFinal 
         Height          =   315
         Left            =   5760
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   720
         Width           =   315
         _Version        =   65536
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   78
         Caption         =   "X"
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
      Begin MSMask.MaskEdBox MskEmpreendimentoInicial 
         Height          =   315
         Left            =   960
         TabIndex        =   2
         Top             =   320
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   4
         Mask            =   "AAAA"
         PromptChar      =   " "
      End
      Begin MSDataListLib.DataCombo CboEmpreendimentoInicial 
         Bindings        =   "FrmRelAnexoFluxoRecebiveisReal.frx":004E
         Height          =   315
         Left            =   1680
         TabIndex        =   3
         Top             =   320
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "empd_tx_nome"
         BoundColumn     =   "empd_cd_empreendimento"
         Text            =   "CboEmpreendimento"
      End
      Begin Threed.SSCommand cmdLimparEmpreendimentoInicial 
         Height          =   315
         Left            =   5760
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   315
         Width           =   315
         _Version        =   65536
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   78
         Caption         =   "X"
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
      Begin VB.Label lblEmpdInicial 
         Caption         =   "Inicial"
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
         TabIndex        =   1
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblEmpdFinal 
         Caption         =   "Final"
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
         TabIndex        =   5
         Top             =   720
         Width           =   615
      End
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "Imprimir..."
      Height          =   345
      Left            =   4440
      TabIndex        =   19
      Top             =   3195
      Width           =   1035
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "Cancelar"
      Height          =   345
      Left            =   5565
      TabIndex        =   20
      Top             =   3195
      Width           =   1035
   End
   Begin MSComDlg.CommonDialog dlgSalvar 
      Left            =   2520
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc DatEmpreendimento 
      Height          =   330
      Left            =   240
      Top             =   3240
      Visible         =   0   'False
      Width           =   1980
      _ExtentX        =   3493
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
   Begin VB.Frame Frame1 
      Height          =   3135
      Left            =   120
      TabIndex        =   21
      Top             =   0
      Width           =   6495
      Begin VB.Frame FraCotacao 
         Caption         =   "Cotação"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   9
         Top             =   1440
         Width           =   6255
         Begin VB.OptionButton optValorHistorico 
            Caption         =   "Valor Historico"
            Height          =   195
            Left            =   360
            TabIndex        =   10
            Top             =   360
            Width           =   1455
         End
         Begin VB.OptionButton optIndexador1 
            Caption         =   "Indexador 1"
            Height          =   195
            Left            =   2400
            TabIndex        =   11
            Top             =   360
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.OptionButton optIndexador2 
            Caption         =   "Indexador 2"
            Height          =   255
            Left            =   4440
            TabIndex        =   12
            Top             =   360
            Width           =   1335
         End
      End
      Begin MSComCtl2.DTPicker DtBase 
         Height          =   315
         Left            =   120
         TabIndex        =   14
         Top             =   2600
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         Format          =   68747265
         CurrentDate     =   39702
      End
      Begin VB.Label lblIndexador2 
         Caption         =   "Indexador 2"
         Height          =   255
         Left            =   4450
         TabIndex        =   17
         Top             =   2325
         Width           =   1095
      End
      Begin VB.Label lblIndexador1 
         Caption         =   "Indexador 1"
         Height          =   255
         Left            =   2290
         TabIndex        =   15
         Top             =   2325
         Width           =   1095
      End
      Begin VB.Label lblDataBase 
         Caption         =   "Data Base"
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
         Left            =   140
         TabIndex        =   13
         Top             =   2320
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc DatMoeda 
      Height          =   330
      Left            =   240
      Top             =   3600
      Visible         =   0   'False
      Width           =   1980
      _ExtentX        =   3493
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
      Caption         =   "DatMoeda"
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
End
Attribute VB_Name = "FrmRelAnexoFluxoRecebiveisReal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public XFO_ExlObj As Excel.Application      ' Create excel object

Private Sub Form_Load()
  subConectarControleDadosNV DatEmpreendimento, "SELECT * FROM Empreendimentos WHERE empr_cd_empresa = " & PCodEmpresa & " ORDER BY empd_cd_empreendimento", Estatico
  
  subConectarControleDadosNV DatMoeda, "SELECT moed_tx_descricao FROM moedas order by moed_tx_descricao", Estatico
  
  'DtBase.Value = Date
  'dtChaves.Value = Date
  'DtPeriodoCarencia.Value = Date
  'DtLimite.Value = Date
End Sub

Private Sub CboEmpreendimentoFinal_Click(Area As Integer)
   If CboEmpreendimentoFinal.Text <> "" Then
      DatEmpreendimento.Recordset.bookmark = CboEmpreendimentoFinal.SelectedItem
      MskEmpreendimentoFinal.Text = CboEmpreendimentoFinal.BoundText
   End If
End Sub

Private Sub CboEmpreendimentoInicial_Click(Area As Integer)
    If CboEmpreendimentoInicial.Text <> "" Then
        DatEmpreendimento.Recordset.bookmark = CboEmpreendimentoInicial.SelectedItem
        MskEmpreendimentoInicial.Text = CboEmpreendimentoInicial.BoundText
    End If
End Sub

Private Sub CmdCancelar_Click()
    Unload Me
End Sub

Private Sub CarregaCabecalhoExcel()

    Dim XLO_DATAUX As ADODB.Recordset
    
    With XFO_ExlObj.ActiveSheet
    
        .Name = "FLUXO DE " & Format(DtBase.Value, "DD.MM.YYYY")
        .Range("F8").ColumnWidth = 2
        
        '-----------
        .Range("A8").ColumnWidth = 9
        .Cells(8, 1).Value = "Permutas"
        .Cells(8, 1).HorizontalAlignment = xlHAlignCenter
        .Cells(8, 1).VerticalAlignment = xlVAlignCenter
        .Cells(8, 1).Font.Bold = True
        
        .Range("B8").ColumnWidth = 9
        .Cells(8, 2).Value = "Vendas Quitadas"
        .Cells(8, 2).HorizontalAlignment = xlHAlignCenter
        .Cells(8, 2).VerticalAlignment = xlVAlignCenter
        .Cells(8, 2).Font.Bold = True
        .Cells(8, 2).WrapText = True
        
        .Range("C8").ColumnWidth = 9
        .Cells(8, 3).Value = "Vendas à Prazo"
        .Cells(8, 3).HorizontalAlignment = xlHAlignCenter
        .Cells(8, 3).VerticalAlignment = xlVAlignCenter
        .Cells(8, 3).Font.Bold = True
        .Cells(8, 3).WrapText = True
        
        .Range("D8").ColumnWidth = 9
        .Cells(8, 4).Value = "Estoque"
        .Cells(8, 4).HorizontalAlignment = xlHAlignCenter
        .Cells(8, 4).VerticalAlignment = xlVAlignCenter
        .Cells(8, 4).Font.Bold = True
        
        .Range("E8").ColumnWidth = 9
        .Cells(8, 5).Value = "Anuência"
        .Cells(8, 5).HorizontalAlignment = xlHAlignCenter
        .Cells(8, 5).VerticalAlignment = xlVAlignCenter
        .Cells(8, 5).Font.Bold = True
        
        '-----------
        .Range("G3").ColumnWidth = 8
        .Cells(3, 7).Value = "CONTRATO"
        .Cells(3, 7).HorizontalAlignment = xlHAlignLeft
        
        .Range("G4").ColumnWidth = 12
        .Cells(4, 7).Value = "NUMERO TOTAL DE UNIDADES"
        .Cells(4, 7).HorizontalAlignment = xlHAlignLeft
        
        SubQOpenRecordset XLO_DATAUX, "SELECT * FROM Empresas WHERE empr_cd_empresa = " & PCodEmpresa, Estatico
        .Range("G5").ColumnWidth = 12
        .Cells(5, 7).Value = XLO_DATAUX!empr_tx_Fantasia
        .Cells(5, 7).HorizontalAlignment = xlHAlignLeft
        
        '-----------
        .Range("J2").ColumnWidth = 12
        .Cells(2, 10).Value = "Data Base"
        .Cells(3, 10).Value = Format(DtBase.Value, "MM/DD/YYYY")
        .Cells(5, 10).Value = Format(DtBase.Value, "MM/DD/YYYY")
        .Cells(2, 10).HorizontalAlignment = xlHAlignCenter
        .Cells(3, 10).HorizontalAlignment = xlHAlignCenter
        .Cells(5, 10).HorizontalAlignment = xlHAlignCenter
        .Cells(2, 10).Font.Bold = True
        .Cells(3, 10).Font.Bold = True
        .Cells(5, 10).Font.Bold = True
        
        .Range("J4").ColumnWidth = 12
        .Cells(4, 10).Value = "Data Base"
        .Cells(4, 10).HorizontalAlignment = xlHAlignCenter
        .Cells(4, 10).Font.Bold = True
        
        .Range("K2").ColumnWidth = 12
        .Cells(2, 11).Value = CboIndexador1.Text
        .Cells(2, 11).HorizontalAlignment = xlHAlignCenter
        .Cells(2, 11).Font.Bold = True
        
        .Cells(4, 11).Value = CboIndexador1.Text
        .Cells(4, 11).HorizontalAlignment = xlHAlignCenter
        .Cells(4, 11).Font.Bold = True
        
        '------------
        'Pega as cotações das moedas para o indexador 1
        SubQOpenRecordset XLO_DATAUX, "select * from moedas where moed_tx_descricao = " & _
        "'" & CboIndexador1.Text & "'", Estatico

            If (XLO_DATAUX!moed_tx_Indexacao) = "M" Then
                SubQOpenRecordset XLO_DATAUX, "SELECT distinct m.moed_tx_Descricao, c.cota_vl_valor, moed_tx_Indexacao, c.cota_dt_datacotacao " & _
                "FROM moedas m inner join cotacoesmoedas c " & _
                "ON m.moed_cd_CodMoeda = c.moed_cd_CodMoeda where c.cota_dt_datacotacao = '" & Format(DtBase.Value, "MM/01/YYYY") & "' " & _
                "AND m.moed_tx_Descricao = '" & CboIndexador1.Text & "'", Estatico
            Else
                SubQOpenRecordset XLO_DATAUX, "SELECT distinct m.moed_tx_Descricao, c.cota_vl_valor, moed_tx_Indexacao, c.cota_dt_datacotacao " & _
                "FROM moedas m inner join cotacoesmoedas c " & _
                "ON m.moed_cd_CodMoeda = c.moed_cd_CodMoeda where c.cota_dt_datacotacao = '" & Format(DtBase.Value, "MM/DD/YYYY") & "' " & _
                "AND m.moed_tx_Descricao = '" & CboIndexador1.Text & "'", Estatico
            End If
        .Cells(3, 11).Value = XLO_DATAUX!cota_vl_valor
        .Cells(3, 11).Font.Bold = True
        .Cells(3, 11).Font.Size = 9
        .Cells(3, 11).HorizontalAlignment = xlHAlignCenter
        
        .Cells(5, 11).Value = XLO_DATAUX!cota_vl_valor
        .Cells(5, 11).Font.Bold = True
        .Cells(5, 11).Font.Size = 9
        .Cells(5, 11).HorizontalAlignment = xlHAlignCenter
        
        '------------
        'Pega as cotações das moedas para o indexador 2
        SubQOpenRecordset XLO_DATAUX, "select * from moedas where moed_tx_descricao = " & _
        "'" & CboIndexador2.Text & "'", Estatico

            If (XLO_DATAUX!moed_tx_Indexacao) = "M" Then
                SubQOpenRecordset XLO_DATAUX, "SELECT distinct m.moed_tx_Descricao, c.cota_vl_valor, moed_tx_Indexacao, c.cota_dt_datacotacao " & _
                "FROM moedas m inner join cotacoesmoedas c " & _
                "ON m.moed_cd_CodMoeda = c.moed_cd_CodMoeda where c.cota_dt_datacotacao = '" & Format(DtBase.Value, "MM/01/YYYY") & "' " & _
                "AND m.moed_tx_Descricao = '" & CboIndexador2.Text & "'", Estatico
            Else
                SubQOpenRecordset XLO_DATAUX, "SELECT distinct m.moed_tx_Descricao, c.cota_vl_valor, moed_tx_Indexacao, c.cota_dt_datacotacao " & _
                "FROM moedas m inner join cotacoesmoedas c " & _
                "ON m.moed_cd_CodMoeda = c.moed_cd_CodMoeda where c.cota_dt_datacotacao = '" & Format(DtBase.Value, "DD/MM/YYYY") & "' " & _
                "AND m.moed_tx_Descricao = '" & CboIndexador2.Text & "'", Estatico
            End If
            
        .Range("L2").ColumnWidth = 12
        .Cells(2, 12).Value = CboIndexador2.Text
        .Cells(2, 12).HorizontalAlignment = xlHAlignCenter
        .Cells(2, 12).Font.Bold = True
        
        .Cells(4, 12).Value = CboIndexador2.Text
        .Cells(4, 12).HorizontalAlignment = xlHAlignCenter
        .Cells(4, 12).Font.Bold = True
        
        .Cells(3, 12).Value = XLO_DATAUX!cota_vl_valor
        .Cells(3, 12).Font.Bold = True
        .Cells(3, 12).Font.Size = 9
        .Cells(3, 12).HorizontalAlignment = xlHAlignCenter
        
        .Cells(5, 12).Value = XLO_DATAUX!cota_vl_valor
        .Cells(5, 12).Font.Bold = True
        .Cells(5, 12).Font.Size = 9
        .Cells(5, 12).HorizontalAlignment = xlHAlignCenter
        
        '-----------
        .Cells(2, 4).HorizontalAlignment = xlHAlignCenter
        .Cells(3, 4).HorizontalAlignment = xlHAlignCenter
        .Cells(4, 4).HorizontalAlignment = xlHAlignCenter
        .Cells(5, 4).HorizontalAlignment = xlHAlignCenter
        .Cells(6, 4).HorizontalAlignment = xlHAlignCenter
        .Cells(7, 4).HorizontalAlignment = xlHAlignCenter
        
        '-----------
        'Preenche os dados do banco
        SubQOpenRecordset XLO_DATAUX, "SELECT * FROM Empreendimentos WHERE empr_cd_empresa = " & PCodEmpresa & " AND empd_tx_Nome = '" & CboEmpreendimentoInicial & "'", Estatico
        .Cells(2, 1).Value = XLO_DATAUX!empd_tx_Nome
        .Cells(2, 1).Font.Bold = True
        .Cells(2, 4).Value = XLO_DATAUX!empd_nr_Unidades
        .Cells(2, 4).Font.Bold = True
        .Cells(2, 5).Value = "APTOS."
        .Cells(2, 5).Font.Bold = True
        .Cells(3, 3).Value = "Vendidos"
        .Cells(3, 4).Value = XLO_DATAUX!empd_nr_unvendidas
        .Cells(4, 4).Value = "=D2-D3"
        
        SubQOpenRecordset XLO_DATAUX, "SELECT * FROM Empreendimentos WHERE empr_cd_empresa = " & PCodEmpresa & " AND empd_tx_Nome = '" & CboEmpreendimentoFinal & "'", Estatico
        .Cells(5, 1).Value = XLO_DATAUX!empd_tx_Nome
        .Cells(5, 1).Font.Bold = True
        .Cells(5, 4).Value = XLO_DATAUX!empd_nr_Unidades
        .Cells(5, 4).Font.Bold = True
        .Cells(5, 5).Value = "APTOS."
        .Cells(5, 5).Font.Bold = True
        .Cells(6, 3).Value = "Vendidos"
        .Cells(6, 4).Value = XLO_DATAUX!empd_nr_unvendidas
        .Cells(7, 4).Value = "=D5-D6"
        
        '----------
        .Cells(10, 3).Value = "=D3+D6"
        .Cells(10, 4).Value = "=D2+D5-C10"
        
        '----------
        .Range("H10").ColumnWidth = 50
        
        '----------
        With XFO_ExlObj.Range("A2:D7").Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlHairline
            .ColorIndex = xlAutomatic
        End With
        With XFO_ExlObj.Range("A2:D7").Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlHairline
            .ColorIndex = xlAutomatic
        End With
        With XFO_ExlObj.Range("A2:D7").Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlHairline
            .ColorIndex = xlAutomatic
        End With
        With XFO_ExlObj.Range("A2:D7").Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .Weight = xlHairline
            .ColorIndex = xlAutomatic
        End With
        With XFO_ExlObj.Range("A2:D7").Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .Weight = xlHairline
            .ColorIndex = xlAutomatic
        End With
        
        'Insere bordas na planilha
        With XFO_ExlObj.Range("A8:E8").Borders(xlEdgeTop)
                .LineStyle = xlDouble
                .Weight = xlThick
                .ColorIndex = xlAutomatic
        End With
        
        With XFO_ExlObj.Range("A8;B8;C8;D8;E8;F8").Borders(xlEdgeLeft)
                .LineStyle = xlDouble
                .Weight = xlThick
                .ColorIndex = xlAutomatic
        End With
        
        With XFO_ExlObj.Range("A8:E8").Borders(xlEdgeBottom)
            .LineStyle = xlDouble
            .Weight = xlThick
            .ColorIndex = xlAutomatic
        End With
        
        '----------
        With XFO_ExlObj.Range("A9:E9").Borders(xlEdgeBottom)
            .LineStyle = xlDouble
            .Weight = xlThick
            .ColorIndex = xlAutomatic
        End With
        
        '----------
        With XFO_ExlObj.Range("J2:L5").Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        
        With XFO_ExlObj.Range("J2:L5").Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        
        With XFO_ExlObj.Range("J2:L5").Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        
        With XFO_ExlObj.Range("J2:L5").Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        
        With XFO_ExlObj.Range("J2:L5").Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        
        With XFO_ExlObj.Range("J2:L5").Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        
    End With
    
    'Formata para não exibir as linhas de grade da planilha
    XFO_ExlObj.ActiveWindow.DisplayGridlines = False
End Sub

Private Sub ExportaDadosExcel(dados As ADODB.Recordset, linha, totalColunas)

    Dim XLO_DATAUX As New ADODB.Recordset
    Dim coluna As Integer 'Contador de colunas para a inserção dos dados
    Dim Empreendimento As String 'Testa se mudou de empreendimento

     With XFO_ExlObj.ActiveSheet
     
        Dim i As Integer
        
        coluna = 7
        
        'Preenchimento dos fields na planilha
            For i = 1 To dados.Fields.Count - 1
                .Range(ConvertToLetter(coluna) & linha & ":" & ConvertToLetter(coluna) & (linha - 1)).Merge
                If coluna <> 8 Then 'Caso a coluna seja a coluna dos nomes
                    .Range(ConvertToLetter(coluna) & linha).ColumnWidth = 12
                End If
                
                If coluna <= 13 Then 'Separar apenas as palavras
                    .Cells((linha - 1), coluna).Value = SepararPalavrasFields(dados.Fields(i).Name)
                Else
                    .Cells((linha - 1), coluna).Value = dados.Fields(i).Name
                    '.Cells((linha - 1), coluna).Formula = UCase(Cells((linha - 1), coluna).Formula)
                    .Range(ConvertToLetter(coluna) & (linha - 1) & ":" & ConvertToLetter(coluna) & linha).NumberFormat = "mmm-yy"
                End If
                    .Cells((linha - 1), coluna).HorizontalAlignment = xlHAlignCenter
                    .Cells((linha - 1), coluna).VerticalAlignment = xlVAlignCenter
                    .Cells((linha - 1), coluna).Font.Bold = True
                    .Cells((linha - 1), coluna).WrapText = True
                
                'Preenchimento das bordas
                With XFO_ExlObj.Range(ConvertToLetter(coluna) & (linha - 1) & ":" & ConvertToLetter(coluna) & (linha - 1)).Borders(xlEdgeTop)
                    .LineStyle = xlDouble
                    .Weight = xlThick
                    .ColorIndex = xlAutomatic
                End With
                With XFO_ExlObj.Range(ConvertToLetter(coluna) & linha & ":" & ConvertToLetter(coluna) & linha & ";" & _
                                      ConvertToLetter(coluna) & (linha + 1) & ":" & ConvertToLetter(coluna) & (linha + 1)).Borders(xlEdgeBottom)
                    .LineStyle = xlDouble
                    .Weight = xlThick
                    .ColorIndex = xlAutomatic
                End With
                With XFO_ExlObj.Range(ConvertToLetter(coluna) & linha & ":" & ConvertToLetter(coluna) & linha & ";" & _
                                      ConvertToLetter(coluna) & (linha - 1) & ":" & ConvertToLetter(coluna) & (linha - 1)).Borders(xlEdgeLeft)
                    .LineStyle = xlDouble
                    .Weight = xlThick
                    .ColorIndex = xlAutomatic
                End With
                With XFO_ExlObj.Range(ConvertToLetter(coluna) & linha & ":" & ConvertToLetter(coluna) & linha & ";" & _
                                      ConvertToLetter(coluna) & (linha - 1) & ":" & ConvertToLetter(coluna) & (linha - 1)).Borders(xlEdgeRight)
                    .LineStyle = xlDouble
                    .Weight = xlThick
                    .ColorIndex = xlAutomatic
                End With
                    coluna = coluna + 1
            Next
        
        coluna = 7
        linha = linha + 1
        'dados.MoveFirst
        
        Empreendimento = dados.Fields(0)
        
        .Cells(linha, 8).HorizontalAlignment = xlHAlignCenter
        .Range(ConvertToLetter(coluna) & linha).RowHeight = 27
        
        SubQOpenRecordset XLO_DATAUX, "SELECT * FROM Empreendimentos WHERE empr_cd_empresa = " & PCodEmpresa & " AND empd_cd_Empreendimento = '" & Empreendimento & "'", Estatico
        If Not XLO_DATAUX.EOF And Not XLO_DATAUX.BOF Then
            .Cells(linha, 8).Value = XLO_DATAUX!empd_tx_Nome
        End If
        
        .Cells(linha, 8).Font.Bold = True
        .Cells(linha, 8).Font.Size = 11
        
        linha = linha + 1
        'Preenchimento dos dados na planilha
        
        While Not dados.EOF
            .Range("G" & linha).NumberFormat = "General"
            .Range("I" & linha).NumberFormat = "m/d/yyyy"
            .Cells(linha, coluna - 1).Value = "=IF(" & ConvertToLetter(coluna - 2) & linha & ">0,1,0"
            .Range(ConvertToLetter(coluna) & linha).RowHeight = 15.75
            For i = 1 To dados.Fields.Count - 1
                If coluna >= 10 Then
                    .Range(ConvertToLetter(coluna) & linha & ":" & ConvertToLetter(coluna) & linha).NumberFormat = "#,##0.00"
                End If
                If coluna >= 12 Then
                    .Cells(linha, coluna).Font.Size = 9
                    .Cells(linha, coluna).HorizontalAlignment = xlHAlignCenter
                End If
                
                .Cells(linha, coluna).Value = dados.Fields(i)
                
                If coluna < 8 Then 'Caso a coluna seja a coluna dos nomes
                    .Cells(linha, coluna).HorizontalAlignment = xlHAlignCenter
                End If
                'If i <> 0 Then 'Teste para ver se é o primeiro dado inserido, se for o primeiro, insere a borda de cima
                    With XFO_ExlObj.Range("A" & linha & ":E" & linha & ";" & _
                                          ConvertToLetter(coluna) & linha & ":" & ConvertToLetter(coluna) & linha).Borders(xlEdgeTop)
                        .LineStyle = xlContinuous
                        .Weight = xlHairline
                        .ColorIndex = xlAutomatic
                    End With
                'End If
                With XFO_ExlObj.Range("A" & linha & ":E" & linha & ";" & _
                                      ConvertToLetter(coluna) & linha & ":" & ConvertToLetter(coluna) & linha).Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .Weight = xlHairline
                    .ColorIndex = xlAutomatic
                End With
                With XFO_ExlObj.Range("A" & linha & ":E" & linha & ";" & _
                                      ConvertToLetter(coluna) & linha & ":" & ConvertToLetter(coluna) & linha).Borders(xlEdgeLeft)
                    .LineStyle = xlContinuous
                    .Weight = xlHairline
                    .ColorIndex = xlAutomatic
                End With
                With XFO_ExlObj.Range("A" & linha & ":E" & linha & ";" & _
                                      ConvertToLetter(coluna) & linha & ":" & ConvertToLetter(coluna) & linha).Borders(xlEdgeRight)
                    .LineStyle = xlContinuous
                    .Weight = xlHairline
                    .ColorIndex = xlAutomatic
                End With
                With XFO_ExlObj.Range("A" & linha & ":E" & linha & ";" & _
                                      ConvertToLetter(coluna) & linha & ":" & ConvertToLetter(coluna) & linha).Borders(xlInsideVertical)
                    .LineStyle = xlContinuous
                    .Weight = xlHairline
                    .ColorIndex = xlAutomatic
                End With
                    coluna = coluna + 1
                    
                    If linha = 10 Then
                        totalColunas = totalColunas + 1
                    End If
            Next
                
                'Preenchimento dos totais dos dados
                .Cells(linha, ConvertToLetter(coluna + 1)).Value = "=SUM(K" & linha & ":" & ConvertToLetter(coluna) & linha & ")"
                .Cells(linha, ConvertToLetter(coluna + 1)).Font.ColorIndex = 3
                .Range(ConvertToLetter(coluna + 1) & linha).ColumnWidth = 12

                linha = linha + 1
                coluna = 7
                dados.MoveNext
                
                If Not dados.EOF Then
                    If dados.Fields(0) <> Empreendimento Then
                        Empreendimento = dados.Fields(0)
                        .Cells(linha, 8).HorizontalAlignment = xlHAlignCenter
                        .Range(ConvertToLetter(coluna) & linha).RowHeight = 27
                        
                        SubQOpenRecordset XLO_DATAUX, "SELECT * FROM Empreendimentos WHERE empr_cd_empresa = " & PCodEmpresa & " AND empd_cd_Empreendimento = '" & Empreendimento & "'", Estatico
                        .Cells(linha, 8).Value = XLO_DATAUX!empd_tx_Nome
                        .Cells(linha, 8).Font.Bold = True
                        .Cells(linha, 8).Font.Size = 11
                        linha = linha + 1
                    Else
                        Empreendimento = dados.Fields(0)
                    End If
                End If
        Wend
    End With

End Sub

Private Sub CarregaTotais(linha, totalColunas)
    Dim coluna As Integer
    
    With XFO_ExlObj.ActiveSheet
    totalColunas = totalColunas + 6
        For coluna = 1 To 12 'dados.Fields.Count - 1
            If coluna <= 5 Then
                
                'Preenchimento das bordas
                With XFO_ExlObj.Range(ConvertToLetter(coluna) & linha).Borders(xlEdgeTop)
                    .LineStyle = xlDouble
                    .Weight = xlThick
                    .ColorIndex = xlAutomatic
                End With
                With XFO_ExlObj.Range(ConvertToLetter(coluna) & linha).Borders(xlEdgeBottom)
                    .LineStyle = xlDouble
                    .Weight = xlThick
                    .ColorIndex = xlAutomatic
                End With
                With XFO_ExlObj.Range(ConvertToLetter(coluna) & linha).Borders(xlEdgeLeft)
                    .LineStyle = xlDouble
                    .Weight = xlThick
                    .ColorIndex = xlAutomatic
                End With
                With XFO_ExlObj.Range(ConvertToLetter(coluna) & linha).Borders(xlEdgeRight)
                    .LineStyle = xlDouble
                    .Weight = xlThick
                    .ColorIndex = xlAutomatic
                End With
                
                .Range(ConvertToLetter(coluna) & linha).Interior.ColorIndex = 19
                .Cells(linha, coluna).Value = "=SUM(" & ConvertToLetter(coluna) & "10:" & ConvertToLetter(coluna) & (linha - 1) & ")"
                .Cells(linha, coluna).HorizontalAlignment = xlHAlignCenter
                
                If coluna = 5 Then
                    .Cells(linha, coluna).Value = "=COUNT(" & ConvertToLetter(coluna + 1) & "10:" & ConvertToLetter(coluna + 1) & (linha - 1) & ")"
                    .Cells(linha, coluna).HorizontalAlignment = xlHAlignCenter
                End If
                
            ElseIf coluna >= 7 Then
                With XFO_ExlObj.Range(ConvertToLetter(coluna) & linha).Borders(xlEdgeTop)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .ColorIndex = xlAutomatic
                End With
            End If
                If coluna = 7 Then
                    .Cells(linha, coluna).Value = "=ROWS(" & ConvertToLetter(coluna) & "10:" & ConvertToLetter(coluna) & (linha - 1) & ")"
                ElseIf coluna >= 10 Then
                    .Range(ConvertToLetter(coluna) & linha).ColumnWidth = 15
                    .Cells(linha, coluna).Value = "=SUM(" & ConvertToLetter(coluna) & "10:" & ConvertToLetter(coluna) & (linha - 1) & ")"
                End If
                
                If coluna >= 8 And coluna <= 10 Then
                    With XFO_ExlObj.Range(ConvertToLetter(coluna) & linha + 2).Borders(xlEdgeTop)
                        .LineStyle = xlDouble
                        .Weight = xlThick
                        .ColorIndex = xlAutomatic
                    End With
                    With XFO_ExlObj.Range(ConvertToLetter(coluna) & linha + 2).Borders(xlEdgeBottom)
                        .LineStyle = xlDouble
                        .Weight = xlThick
                        .ColorIndex = xlAutomatic
                    End With
                    With XFO_ExlObj.Range(ConvertToLetter(coluna) & linha + 2).Borders(xlEdgeLeft)
                        .LineStyle = xlDouble
                        .Weight = xlThick
                        .ColorIndex = xlAutomatic
                    End With
                    With XFO_ExlObj.Range(ConvertToLetter(coluna) & linha + 2).Borders(xlEdgeRight)
                        .LineStyle = xlDouble
                        .Weight = xlThick
                        .ColorIndex = xlAutomatic
                    End With
                        .Range(ConvertToLetter(coluna) & (linha + 2)).Interior.ColorIndex = 6
                        .Cells(linha + 2, 8).HorizontalAlignment = xlHAlignCenter
                        .Cells(linha + 2, 9).HorizontalAlignment = xlHAlignCenter
                        .Cells(linha + 2, 8).Value = "Valor Estoque"
                        .Cells(linha + 2, 8).Font.Bold = True
                        .Cells(linha + 2, 9).Value = "R$"
                End If
        Next
        
        coluna = 13
        
        With XFO_ExlObj.Range(ConvertToLetter(coluna) & linha & ":" & ConvertToLetter(totalColunas) & (linha + 2)).Borders(xlEdgeBottom)
            .LineStyle = xlDouble
            .Weight = xlThick
            .ColorIndex = xlAutomatic
        End With
        With XFO_ExlObj.Range(ConvertToLetter(coluna) & linha & ":" & ConvertToLetter(totalColunas) & (linha + 2)).Borders(xlEdgeLeft)
            .LineStyle = xlDouble
            .Weight = xlThick
            .ColorIndex = xlAutomatic
        End With
        With XFO_ExlObj.Range(ConvertToLetter(coluna) & linha & ":" & ConvertToLetter(totalColunas) & (linha + 2)).Borders(xlEdgeRight)
            .LineStyle = xlDouble
            .Weight = xlThick
            .ColorIndex = xlAutomatic
        End With
        With XFO_ExlObj.Range(ConvertToLetter(coluna) & linha & ":" & ConvertToLetter(totalColunas) & (linha + 2)).Borders(xlEdgeTop)
            .LineStyle = xlDouble
            .Weight = xlThick
            .ColorIndex = xlAutomatic
        End With
        With XFO_ExlObj.Range(ConvertToLetter(coluna) & linha & ":" & ConvertToLetter(totalColunas) & (linha + 2)).Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .Weight = xlHairline
            .ColorIndex = xlAutomatic
        End With
        With XFO_ExlObj.Range(ConvertToLetter(coluna) & linha & ":" & ConvertToLetter(totalColunas) & (linha + 2)).Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .Weight = xlHairline
            .ColorIndex = xlAutomatic
        End With
        
        .Cells(linha, coluna).Value = "Total Mês"
        .Cells(linha, coluna).Font.Bold = True
        .Cells((linha + 1), coluna).Value = "Acumulado"
        .Cells((linha + 1), coluna).Font.Bold = True
        .Cells((linha + 2), coluna).Value = "Lançamentos"
        .Cells((linha + 2), coluna).Font.Bold = True
        coluna = coluna + 1
        
        'Descobrir como pegar a ultima coluna
        For coluna = coluna To totalColunas
            .Range(ConvertToLetter(coluna) & linha).ColumnWidth = 14
            .Cells(linha, coluna).Value = "=SUM(" & ConvertToLetter(coluna) & "10:" & ConvertToLetter(coluna) & (linha - 1) & ")"
            .Cells(linha, coluna).HorizontalAlignment = xlHAlignCenter
            .Cells((linha + 1), coluna).Value = "=SUM(" & ConvertToLetter(coluna) & linha & ":" & ConvertToLetter(totalColunas) & linha & ")"
            .Cells((linha + 1), coluna).HorizontalAlignment = xlHAlignCenter
            .Cells((linha + 2), coluna).Value = "=COUNT(" & ConvertToLetter(coluna) & "10:" & ConvertToLetter(coluna) & (linha - 1) & ")"
            .Cells((linha + 2), coluna).HorizontalAlignment = xlHAlignCenter
        Next
    End With
End Sub

'Cria o relatório em uma planilha
Private Sub CmdImprimir_Click()
    
    Dim XLO_COMANDO As New ADODB.Command
    Dim linha As Integer
    Dim totalColunas As Integer
    
    On Error GoTo TrataErro

    If (MskEmpreendimentoInicial.Text = "" Or MskEmpreendimentoFinal.Text = "") Then
        MsgBox "É preciso selecionar o empreendimento inicial e final!", vbInformation, "Atenção"
        Exit Sub
    End If
    
    If IsNull(CboIndexador1.SelectedItem) Or IsNull(CboIndexador2.SelectedItem) Then
        MsgBox "É preciso selecionar os dois indexadores!", vbInformation, "Atenção"
        Exit Sub
    End If
        

    'Seta o Titulo da Tela de Dialogo salvar
    dlgSalvar.DialogTitle = "Salvar Arquivo"

    'Seta o filtro das extensões de arquivos que será usada
    dlgSalvar.Filter = "Documento do Excel(*.xls)|*.xls"

    'Indica qual sera o filtro padrão
    dlgSalvar.FilterIndex = 1

    dlgSalvar.Flags = cdlOFNHideReadOnly + cdlOFNOverwritePrompt + cdlOFNPathMustExist

    'Exibe a tela do dialogo para o usuário
    dlgSalvar.ShowSave

    'Verifica se o usuário selecionou o caminho do arquivo
    If (dlgSalvar.FileName <> "") Then

        'Indexador
        If optValorHistorico.Value = True Then
          XLI_INDEXADOR = 0
        ElseIf optIndexador1.Value = True Then
          XLI_INDEXADOR = 1
        ElseIf optIndexador2.Value = True Then
          XLI_INDEXADOR = 2
        End If
        
        'Altera o cursor do mouse
        FrmRelAnexoFluxoRecebiveisReal.MousePointer = vbHourglass

        'Inicializa o objeto excel
        Set XFO_ExlObj = CreateObject("excel.application")

        'Adiciona um WorkBook
        XFO_ExlObj.Workbooks.Add
        
        'Configura a planilha e preenche o cabeçalho
        CarregaCabecalhoExcel
        
        'Impede que o usuário veja a planilha enquanto ela está sendo preenchida
        XFO_ExlObj.Visible = False
        DatEmpreendimento.Recordset.MoveFirst
        
        linha = 8

        With XLO_COMANDO
            .CommandTimeout = 300
            .CommandText = "spCAPI_sel_calculaContratoREAL"
            .CommandType = adCmdStoredProc
            Set XLO_COMANDO.ActiveConnection = Conexao
            'Cria automaticamente todos os parâmetros dentro do objeto Comando
            .Parameters.Refresh
        End With
                
        XLO_COMANDO.Parameters(1).Value = PCodEmpresa 'Empresa
        XLO_COMANDO.Parameters(2).Value = Format(DtBase.Value, "MM/DD/YYYY") 'Data base para o cálculo
        XLO_COMANDO.Parameters(3).Value = MskEmpreendimentoInicial.Text 'Código Empreendimento Inicial
        XLO_COMANDO.Parameters(4).Value = MskEmpreendimentoFinal.Text 'Código Empreendimento Final
        XLO_COMANDO.Parameters(5).Value = XLI_INDEXADOR 'Indexador

        'Passa para o recordset o resultado da consulta
        Set XGO_RSRELATORIO = XLO_COMANDO.Execute

        'Cria a planilha com os valores previstos dos imóveis
        ExportaDadosExcel XGO_RSRELATORIO, linha, totalColunas
        
        XGO_RSRELATORIO.Close
                
        CarregaTotais linha, totalColunas
        
        'Salva a planilha no caminho selecionado pelo usuário
        XFO_ExlObj.Workbooks(1).SaveAs dlgSalvar.FileName
        XFO_ExlObj.Workbooks(1).Close
        XFO_ExlObj.Application.Quit
        Set XFO_ExlObj = Nothing

        'Limpa o caminho do arquivo
        dlgSalvar.FileName = ""
        FrmRelAnexoFluxoRecebiveisReal.MousePointer = vbDefault
        MsgBox "Relatório Exportado com sucesso!", vbInformation, "Informação"
        
    End If

Exit Sub

TrataErro:
    If Err.Number = cdlCancel Then
        Exit Sub
    End If

    'Fecha a aplicação - Lucas Santiago - 02.02.2009
    XFO_ExlObj.Workbooks(1).SaveAs dlgSalvar.FileName
    XFO_ExlObj.Workbooks(1).Close
    XFO_ExlObj.Application.Quit
    Set XFO_ExlObj = Nothing
    
    MsgBox "Ocorreu um erro na exportação do relatório. Contate o suporte. Detalhe do Erro: " & Err.Description, vbCritical, "Erro"
    DeleteFile (dlgSalvar.FileName)

    Conexao.RollbackTrans
    FrmRelAnexoFluxoRecebiveisReal.MousePointer = vbDefault
End Sub

Private Sub CmdLimpar_Click()
    CboCliente.Text = ""
End Sub

Private Sub cmdLimparEmpreendimentoFinal_Click()
    CboEmpreendimentoFinal.BoundText = ""
    MskEmpreendimentoFinal.Text = ""
End Sub

Private Sub cmdLimparEmpreendimentoInicial_Click()
    CboEmpreendimentoIncial.BoundText = ""
    MskEmpreendimentoInicial.Text = ""
End Sub

Private Sub Form_Activate()
  Set Formulario = FrmRelAnexoFluxoRecebiveisReal
End Sub

Function ConvertToLetter(ByVal valor As Integer) As String
    Dim PrimeiraLetra As Integer
    Dim SegundaLetra As Integer
    
    If ((valor / 26) = (valor \ 26)) Then
        PrimeiraLetra = valor \ 26 - 1
    Else
        PrimeiraLetra = valor \ 26
    End If
    
    SegundaLetra = valor - (PrimeiraLetra * 26)
   
    If PrimeiraLetra > 0 Then
        ConvertToLetter = Chr(PrimeiraLetra + 64)
    End If
    
    If SegundaLetra > 0 Then
        ConvertToLetter = ConvertToLetter & Chr(SegundaLetra + 64)
    End If
End Function

Function SepararPalavrasFields(palavra As String) As String
    Dim posicao As Integer
    
    For i = 2 To Len(palavra)
         If Mid(palavra, i, 1) = UCase(Mid(palavra, i, 1)) Then
            palavraAux = Mid(palavra, 1, (i - 1)) & " "
            posicao = i
         End If
    Next
        If palavraAux <> Null Or palavraAux <> "" Then ' Tratamento de variável nula ou vazia - Lucas Santiago - 04.02.2009
            palavra = palavraAux & Mid(palavra, posicao, Len(palavra))
        End If
        SepararPalavrasFields = palavra
End Function

Private Sub CarregaMeses()
    Dim Ano As Integer
    Dim Meses As Variant
    Dim coluna As Integer
    
    coluna = 14
    Meses = "JAN FEV MAR ABR MAI JUN JUL AGO SET OUT NOV DEZ"
    Meses = Split(Meses)

    Ano = 2008
End Sub

'Seleciona conteudo do Maskedit
Sub subSelecionaMSK(XObjeto As MaskEdBox)
    XObjeto.SelStart = 0
    XObjeto.SelLength = XObjeto.MaxLength
End Sub

Private Sub MskEmpreendimentoInicial_GotFocus()
    Call subSelecionaMSK(MskEmpreendimentoInicial)
End Sub

Private Sub MskEmpreendimentoFinal_GotFocus()
    Call subSelecionaMSK(MskEmpreendimentoFinal)
End Sub

Private Sub MskEmpreendimentoInicial_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 97 And KeyAscii <= 122 Then
        KeyAscii = KeyAscii - 32
    End If
End Sub

Private Sub MskEmpreendimentoFinal_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 97 And KeyAscii <= 122 Then
        KeyAscii = KeyAscii - 32
    End If
End Sub

Private Sub MskEmpreendimentoInicial_LostFocus()
    If MskEmpreendimentoInicial.Text <> "" Then
        SubPreencheComZero MskEmpreendimentoInicial
        CboEmpreendimentoInicial.BoundText = MskEmpreendimentoInicial.Text
        
        If CboEmpreendimentoInicial.Text = "" Then
            MsgBox "Não existe Empreendimento com este código.", vbCritical + vbOKOnly, "ATENÇÃO"
            MskEmpreendimentoInicial.SetFocus
            Exit Sub
        End If
    Else
        CboEmpreendimentoInicial.BoundText = ""
    End If
End Sub

Private Sub MskEmpreendimentoFinal_LostFocus()
    If MskEmpreendimentoFinal.Text <> "" Then
        SubPreencheComZero MskEmpreendimentoFinal
        CboEmpreendimentoFinal.BoundText = MskEmpreendimentoFinal.Text
        
        If CboEmpreendimentoFinal.Text = "" Then
            MsgBox "Não existe Empreendimento com este código.", vbCritical + vbOKOnly, "ATENÇÃO"
            MskEmpreendimentoFinal.SetFocus
            Exit Sub
        End If
    Else
        CboEmpreendimentoFinal.BoundText = ""
    End If
End Sub

