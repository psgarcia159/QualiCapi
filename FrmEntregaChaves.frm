VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmEntregaChaves 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Alteração de Indexador na entrega das chaves"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   4560
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraTitulo 
      Caption         =   "Título"
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
      Left            =   60
      TabIndex        =   5
      Top             =   1140
      Width           =   4395
      Begin MSMask.MaskEdBox MskCodigoInicial 
         Height          =   315
         Left            =   960
         TabIndex        =   7
         Top             =   300
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   9
         Mask            =   "##.###.##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox MskCodigoFinal 
         Height          =   315
         Left            =   2880
         TabIndex        =   9
         Top             =   300
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   9
         Mask            =   "##.###.##"
         PromptChar      =   " "
      End
      Begin VB.Label LblCodigoInicial 
         Caption         =   "Inicial:"
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
         Left            =   360
         TabIndex        =   6
         Top             =   360
         Width           =   555
      End
      Begin VB.Label LlbCodigoFinal 
         Caption         =   "Final:"
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
         Left            =   2340
         TabIndex        =   8
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Frame FraCorrecao 
      Caption         =   "Correção"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   4395
      Begin VB.OptionButton OptIndexador2 
         Caption         =   "Indexador 2"
         Height          =   195
         Left            =   2460
         TabIndex        =   4
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton OptIndexador1 
         Caption         =   "Indexador 1"
         Height          =   195
         Left            =   720
         TabIndex        =   3
         Top             =   720
         Value           =   -1  'True
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker DtpEntregaChaves 
         Height          =   315
         Left            =   3120
         TabIndex        =   2
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/MM/yy"
         Format          =   75825155
         CurrentDate     =   37180
      End
      Begin VB.Label LblEntrega 
         Alignment       =   1  'Right Justify
         Caption         =   "Data Base (Entrega das Chaves):"
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
         Top             =   300
         Width           =   2955
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Novos Indexadores"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   60
      TabIndex        =   10
      Top             =   1920
      Width           =   4395
      Begin MSDataListLib.DataCombo CboMoeda2 
         Bindings        =   "FrmEntregaChaves.frx":0000
         Height          =   315
         Left            =   1680
         TabIndex        =   15
         Top             =   780
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "moed_tx_descricao"
         BoundColumn     =   "moed_cd_codmoeda"
         Text            =   "CboMoeda2"
      End
      Begin MSDataListLib.DataCombo CboMoeda1 
         Bindings        =   "FrmEntregaChaves.frx":0018
         Height          =   315
         Left            =   1680
         TabIndex        =   12
         Top             =   360
         Width           =   1740
         _ExtentX        =   3069
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "moed_tx_descricao"
         BoundColumn     =   "moed_cd_codmoeda"
         Text            =   "CboMoeda1"
      End
      Begin Threed.SSCommand CmdLimparIndex1 
         Height          =   315
         Left            =   3480
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   360
         Width           =   360
         _Version        =   65536
         _ExtentX        =   635
         _ExtentY        =   556
         _StockProps     =   78
         MouseIcon       =   "FrmEntregaChaves.frx":0030
         Picture         =   "FrmEntregaChaves.frx":0482
      End
      Begin Threed.SSCommand CmdLimparIndex2 
         Height          =   315
         Left            =   3480
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   780
         Width           =   360
         _Version        =   65536
         _ExtentX        =   635
         _ExtentY        =   556
         _StockProps     =   78
         MouseIcon       =   "FrmEntregaChaves.frx":0594
         Picture         =   "FrmEntregaChaves.frx":09E6
      End
      Begin VB.Label LblMoeda1 
         Alignment       =   1  'Right Justify
         Caption         =   "Indexador 1:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   600
         TabIndex        =   11
         Top             =   420
         Width           =   1095
      End
      Begin VB.Label LblMoeda2 
         Alignment       =   1  'Right Justify
         Caption         =   "Indexador 2:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   600
         TabIndex        =   14
         Top             =   840
         Width           =   1095
      End
   End
   Begin VB.CommandButton CmdConfirmarAlteracaoIndexador 
      Caption         =   "&Confirmar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2160
      TabIndex        =   21
      Top             =   4080
      Width           =   975
   End
   Begin VB.CommandButton CmdRetornarAlteracaoIndexador 
      Caption         =   "&Retornar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3420
      TabIndex        =   22
      Top             =   4080
      Width           =   975
   End
   Begin VB.Frame FraJuros 
      Caption         =   "Juros"
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
      Left            =   60
      TabIndex        =   17
      Top             =   3240
      Width           =   4395
      Begin VB.TextBox TxtPercJuros 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   3540
         TabIndex        =   20
         Top             =   240
         Width           =   660
      End
      Begin VB.CheckBox ChkAplicarJuros 
         Caption         =   "Aplicar juros sobre as parcelas?"
         Height          =   195
         Left            =   180
         TabIndex        =   18
         Top             =   300
         Width           =   2655
      End
      Begin VB.Label LblPercJuros 
         Caption         =   "Percent.:"
         Enabled         =   0   'False
         Height          =   195
         Left            =   2820
         TabIndex        =   19
         Top             =   300
         Width           =   675
      End
   End
   Begin MSAdodcLib.Adodc DatMoeda1 
      Height          =   330
      Left            =   -60
      Top             =   4020
      Visible         =   0   'False
      Width           =   1515
      _ExtentX        =   2672
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
      Caption         =   "DatMoeda1"
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
   Begin MSAdodcLib.Adodc DatMoeda2 
      Height          =   330
      Left            =   -240
      Top             =   4080
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
      Caption         =   "DatMoeda2"
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
Attribute VB_Name = "FrmEntregaChaves"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim XLF_COTACAO1 As Double
Dim XLF_COTACAO1MESANTERIOR As Double
Dim XLF_COTACAO2 As Double
Dim XLF_COTACAO2MESANTERIOR As Double

Private Sub CboMoeda1_Change()
  
  If CboMoeda1.BoundText <> "" Then
    DatMoeda1.Recordset.bookmark = CboMoeda1.SelectedItem
  End If

End Sub

Private Sub CboMoeda2_Change()

  If CboMoeda2.BoundText <> "" Then
    DatMoeda2.Recordset.bookmark = CboMoeda2.SelectedItem
  End If

End Sub

Private Sub ChkAplicarJuros_Click()
   
   If ChkAplicarJuros.Value = 0 Then
    TxtPercJuros.Enabled = False
    LblPercJuros.Enabled = False
  Else
    TxtPercJuros.Enabled = True
    LblPercJuros.Enabled = True
  End If
   
End Sub

Private Sub CmdConfirmarAlteracaoIndexador_Click()
    
  Dim XLT_SQL As String
  Dim XLO_TITULOS As New ADODB.Recordset
  Dim XLF_VALORCORRIGIDO As Double
  Dim XLF_JUROS As Double
  Dim XLF_PERCJUROS As Double
  Dim XLF_VALORINDEXADO1 As Double
  Dim XLF_VALORINDEXADO2 As Double
  Dim XLI_MOEDA1 As Integer
  Dim XLI_MOEDA2 As Integer
  Dim XLT_DESCRICAO As String
  
  Dim XLI_MAIORPLANO As Integer
  Dim XLI_CONT As Integer
  Dim XLF_SOMATORIO As Double
  Dim XLT_TITULO As String
  Dim XLI_PARCELAS As Integer
  Dim XLD_DATABASE As Date
  
  If ChkAplicarJuros.Value = 1 Then
    If TxtPercJuros.Text = "0,0000" Or TxtPercJuros.Text = "" Then
      MsgBox "Os juros devem ser definidos!", vbCritical, "ATENÇÃO"
      Exit Sub
    End If
  End If
  
  If CboMoeda1.BoundText = "" And CboMoeda2.BoundText = "" Then
    MsgBox "A moeda deve ser escolhida pelo menos para um dos indexadores!", vbCritical, "ATENÇÃO"
    Exit Sub
  End If
  
  If MsgBox("Os títulos em aberto deste contrato terão a data base, o valor" & vbCrLf & _
              "histórico e os valores indexados alterados! Deseja continuar?", vbCritical + vbYesNo, "ATENÇÃO") = vbYes Then
  
    Conexao.BeginTrans
    Me.MousePointer = vbHourglass
    
    If ChkAplicarJuros.Value = False Then
      
      XLT_SQL = FunCriaConsultaBase(DtpEntregaChaves, NomeSgbd, "ConsCAPBasePagamento", 1)
      XLT_SQL = XLT_SQL & " empr_cd_empresa=" & PCodEmpresa & _
        " AND titu_dt_Pagamento is null" & _
        " AND left(Titulo,12) ='" & Chave & "'"

      If MskCodigoInicial.Text <> "  .   .  " And MskCodigoFinal.Text <> "  .   .  " Then
        XLT_SQL = XLT_SQL & _
          " AND Titulo >= '" & Chave & "." & MskCodigoInicial.Text & "'" & _
          " AND Titulo <='" & Chave & "." & MskCodigoFinal.Text & "'"
      End If
      
      SubQOpenRecordset XLO_TITULOS, XLT_SQL, Estatico
      
      If Not (XLO_TITULOS.EOF) Then
      
        subVerificaCotacoes XLO_TITULOS
        
        While Not XLO_TITULOS.EOF
        
          If OptIndexador1.Value = True Then
            'Se for corrigir pelo indexador1
            If Not IsNull(XLO_TITULOS!moeda1) Then
              'Correção Monetária
              XLF_VALORCORRIGIDO = Format(funCalculaCorrecaoMonetaria(XLO_TITULOS, 1, DtpEntregaChaves.Value), "standard")
            Else
              XLF_VALORCORRIGIDO = Format(XLO_TITULOS!titu_vl_Parcela, "standard")
            End If
          Else
            'Se for corrigir pelo indexador2
            If Not IsNull(XLO_TITULOS!moeda2) Then
              'Correção Monetária
              XLF_VALORCORRIGIDO = Format(funCalculaCorrecaoMonetaria(XLO_TITULOS, 2, DtpEntregaChaves.Value), "standard")
            Else
              XLF_VALORCORRIGIDO = Format(XLO_TITULOS!titu_vl_Parcela, "standard")
            End If
          End If
          
          If CboMoeda1.BoundText <> "" Then
            'Se escolher novo indexador1
            XLF_VALORINDEXADO1 = Format(FunValorIndexado(XLF_VALORCORRIGIDO, XLF_COTACAO1, XLF_COTACAO1MESANTERIOR, DatMoeda1.Recordset!moed_tx_Indexacao, XLO_TITULOS!cont_tx_CorrecProRata, DtpEntregaChaves), "##0.0000")
            XLI_MOEDA1 = CboMoeda1.BoundText
          Else
            'Se não escolher novo indexador1
            If Not IsNull(XLO_TITULOS!moed_cd_Moeda1) Then
              'Se já possuir indexador1
              XLF_VALORINDEXADO1 = Format(FunValorIndexado(XLF_VALORCORRIGIDO, XLO_TITULOS!Cotacao1, XLO_TITULOS!Cotacao1_MesAnterior, XLO_TITULOS!Indexacao1, XLO_TITULOS!cont_tx_CorrecProRata, DtpEntregaChaves), "##0.0000")
              XLI_MOEDA1 = XLO_TITULOS!moed_cd_Moeda1
            Else
              'Se não possuir indexador1
              XLF_VALORINDEXADO1 = XLF_VALORCORRIGIDO
              XLI_MOEDA1 = 0
            End If
          End If
          
          If CboMoeda2.BoundText <> "" Then
            'Se escolher novo indexador2
            XLF_VALORINDEXADO2 = Format(FunValorIndexado(XLF_VALORCORRIGIDO, XLF_COTACAO2, XLF_COTACAO2MESANTERIOR, DatMoeda2.Recordset!moed_tx_Indexacao, XLO_TITULOS!cont_tx_CorrecProRata, DtpEntregaChaves), "##0.0000")
            XLI_MOEDA2 = CboMoeda2.BoundText
          Else
            'Se não escolher novo indexador2
            If Not IsNull(XLO_TITULOS!moed_cd_Moeda2) Then
              'Se já possuir indexador2
              XLF_VALORINDEXADO2 = Format(FunValorIndexado(XLF_VALORCORRIGIDO, XLO_TITULOS!Cotacao2, XLO_TITULOS!Cotacao2_MesAnterior, XLO_TITULOS!Indexacao2, XLO_TITULOS!cont_tx_CorrecProRata, DtpEntregaChaves), "##0.0000")
              XLI_MOEDA2 = XLO_TITULOS!moed_cd_Moeda2
            Else
              'Se não possuir indexador2
              XLF_VALORINDEXADO2 = XLF_VALORCORRIGIDO
              XLI_MOEDA2 = 0
            End If
          End If
              
          XLT_SQL = "UPDATE ConsCAPTitulos SET " & _
             " titu_vl_ParcelaIndex1=" & FunTrataFloat(XLF_VALORINDEXADO1) & _
             ", titu_vl_ParcelaIndex2=" & FunTrataFloat(XLF_VALORINDEXADO2) & _
             ", moed_cd_Moeda1=" & FunNuloBancoZero(XLI_MOEDA1) & _
             ", moed_cd_Moeda2=" & FunNuloBancoZero(XLI_MOEDA2) & _
             ", titu_dt_Base = " & FunNuloData(DtpEntregaChaves.Value, NomeSgbd) & _
             ", titu_tx_EntregaChave='1'" & _
             " WHERE Titulo='" & XLO_TITULOS!Titulo & "' AND empr_cd_empresa = " & PCodEmpresa & ""
              
           Conexao.Execute XLT_SQL, , adCmdText
              
           'Descrição do Log
           'Empr = Código da empresa
           'Tit = Codigo do título
           'VHO = Valor Histórico Original
           'I1O = Indexador 1 Original VI1O = Valor Indexado 1 Original
           'I2O = Indexador 2 Original VI2O = Valor Indexado 2 Original
           'DBO = Data Base Original
           'VHA = Valor Histórico Alterado
           'I1A = Indexador 1 Original VI1O = Valor Indexado 1 Alterada
           'I2A = Indexador 2 Original VI2O = Valor Indexado 2 Alterada
           'DBA = Data Base Alterada
           
           XLT_DESCRICAO = "Empr:" & PCodEmpresa & _
                           " Tit:" & XLO_TITULOS!Titulo & _
                           " VHO:" & XLO_TITULOS!titu_vl_Parcela & _
                           " I1O:" & FunNuloBancoVal(XLO_TITULOS!moed_cd_Moeda1) & _
                           " VI1O:" & XLO_TITULOS!titu_vl_ParcelaIndex1 & _
                           " I2O:" & FunNuloBancoVal(XLO_TITULOS!moed_cd_Moeda2) & _
                           " VI2O:" & XLO_TITULOS!titu_vl_ParcelaIndex2 & _
                           " DBO:" & XLO_TITULOS!titu_dt_Base & _
                           " VHA:" & XLF_VALORCORRIGIDO & _
                           " I1A:" & FunNuloBancoZero(XLI_MOEDA1) & _
                           " VI1A:" & XLF_VALORINDEXADO1 & _
                           " I2A:" & FunNuloBancoZero(XLI_MOEDA2) & _
                           " VI2A:" & XLF_VALORINDEXADO2 & _
                           " DBA:" & Format(DtpEntregaChaves, "dd/mm/yy")
           
          Call subGravaLog(func_cd_operador, "TelaManContratos", "9", Now(), XLT_DESCRICAO)
          
          XLO_TITULOS.MoveNext
        Wend
        
        XLO_TITULOS.Close
        Set XLO_TITULOS = Nothing
        
        MsgBox "Alteração dos indexadores realizada com sucesso!", vbInformation
      Else
        MsgBox "Este contrato não possui nenhum título em aberto!", vbCritical, "ATENÇÃO"
      End If
  
    Else
      
      'Seleciona  o maior plano
      XLT_SQL = "SELECT max(titu_cd_Plano) as MaiorPlano FROM ConsCapTitulos WHERE left(Titulo,12) ='" & Chave & "'"
      SubQOpenRecordset XLO_TITULOS, XLT_SQL, Estatico
          
      XLI_MAIORPLANO = XLO_TITULOS!MaiorPlano
      XLO_TITULOS.Close
      Set XLO_TITULOS = Nothing
      
      For XLI_CONT = 0 To XLI_MAIORPLANO
      
        XLD_DATABASE = "01/01/1900"
        XLF_SOMATORIO = 0
        XLT_TITULO = ""
        
        'Cria a consulta base para efetuar os cálculos
        XLT_SQL = FunCriaConsultaBase(DtpEntregaChaves, NomeSgbd, "ConsCAPBasePagamento", 4)
        XLT_SQL = XLT_SQL & " empr_cd_empresa=" & PCodEmpresa & _
          " AND titu_dt_Pagamento is null" & _
          " AND titu_dt_Vencimento>= " & FunNuloData(DtpEntregaChaves.Value, NomeSgbd) & _
          " AND left(Titulo,15) ='" & Chave & "." & Format(XLI_CONT, "00") & "'"
          
        'Alterado para permitir a seleção de títulos com as condições >=, <= e between - 25/03/10
        If MskCodigoInicial.Text <> "  .   .  " Then
           XLT_SQL = XLT_SQL & _
             " AND Titulo >= '" & Chave & "." & MskCodigoInicial.Text & "'"
        End If
        If MskCodigoFinal.Text <> "  .   .  " Then
           XLT_SQL = XLT_SQL & _
             " AND Titulo <='" & Chave & "." & MskCodigoFinal.Text & "'"
        End If
          
        XLT_SQL = XLT_SQL & " ORDER BY Titulo"
          
        SubQOpenRecordset XLO_TITULOS, XLT_SQL, Estatico
          
        If Not XLO_TITULOS.EOF Then
        
          subVerificaCotacoes XLO_TITULOS
          
          While Not XLO_TITULOS.EOF
          
            'Se a periodicidade não for mensal e a diferença do vencimento em relação
            'à data base for menor que a periodicidade altera logo o valor do título
            'Este titulo terá o valor diferente dos demais
            
            'If (CInt(XLO_TITULOS!tipl_nr_Periodicidade) > 1 And XLO_TITULOS!DiferencaMes < CInt(XLO_TITULOS!tipl_nr_Periodicidade)) _
               Or XLO_TITULOS.RecordCount = 1 Then
            If XLO_TITULOS!DiferencaMes < CInt(XLO_TITULOS!tipl_nr_Periodicidade) _
               Or XLO_TITULOS.RecordCount = 1 Then
            
              If OptIndexador1.Value = True Then
                'Se for corrigir pelo indexador1
                If Not IsNull(XLO_TITULOS!moeda1) Then
                  'Correção Monetária
                  XLF_VALORCORRIGIDO = Format(funCalculaCorrecaoMonetaria(XLO_TITULOS, 1, DtpEntregaChaves.Value), "standard")
                Else
                  XLF_VALORCORRIGIDO = Format(XLO_TITULOS!titu_vl_Parcela, "standard")
                End If
              Else
                'Se for corrigir pelo indexador2
                If Not IsNull(XLO_TITULOS!moeda2) Then
                  'Correção Monetária
                  XLF_VALORCORRIGIDO = Format(funCalculaCorrecaoMonetaria(XLO_TITULOS, 2, DtpEntregaChaves.Value), "standard")
                Else
                  XLF_VALORCORRIGIDO = Format(XLO_TITULOS!titu_vl_Parcela, "standard")
                End If
              End If
              
              XLF_JUROS = FunJuros(XLF_VALORCORRIGIDO, TxtPercJuros.Text, "M", XLO_TITULOS!titu_dt_Vencimento, DtpEntregaChaves.Value, 0)
              XLF_VALORCORRIGIDO = XLF_VALORCORRIGIDO + XLF_JUROS
              
              If CboMoeda1.BoundText <> "" Then
                'Se escolher novo indexador1
                XLF_VALORINDEXADO1 = Format(FunValorIndexado(XLF_VALORCORRIGIDO, XLF_COTACAO1, XLF_COTACAO1MESANTERIOR, DatMoeda1.Recordset!moed_tx_Indexacao, XLO_TITULOS!cont_tx_CorrecProRata, DtpEntregaChaves), "##0.0000")
                XLI_MOEDA1 = CboMoeda1.BoundText
              Else
                'Se não escolher novo indexador1
                If Not IsNull(XLO_TITULOS!moed_cd_Moeda1) Then
                  'Se já possuir indexador1
                  XLF_VALORINDEXADO1 = Format(FunValorIndexado(XLF_VALORCORRIGIDO, XLO_TITULOS!Cotacao1, XLO_TITULOS!Cotacao1_MesAnterior, XLO_TITULOS!Indexacao1, XLO_TITULOS!cont_tx_CorrecProRata, DtpEntregaChaves), "##0.0000")
                  XLI_MOEDA1 = XLO_TITULOS!moed_cd_Moeda1
                Else
                  'Se não possuir indexador1
                  XLF_VALORINDEXADO1 = XLF_VALORCORRIGIDO
                  XLI_MOEDA1 = 0
                End If
              End If
              
              If CboMoeda2.BoundText <> "" Then
                'Se escolher novo indexador2
                XLF_VALORINDEXADO2 = Format(FunValorIndexado(XLF_VALORCORRIGIDO, XLF_COTACAO2, XLF_COTACAO2MESANTERIOR, DatMoeda2.Recordset!moed_tx_Indexacao, XLO_TITULOS!cont_tx_CorrecProRata, DtpEntregaChaves), "##0.0000")
                XLI_MOEDA2 = CboMoeda2.BoundText
              Else
                'Se não escolher novo indexador2
                If Not IsNull(XLO_TITULOS!moed_cd_Moeda2) Then
                  'Se já possuir indexador2
                  XLF_VALORINDEXADO2 = Format(FunValorIndexado(XLF_VALORCORRIGIDO, XLO_TITULOS!Cotacao2, XLO_TITULOS!Cotacao2_MesAnterior, XLO_TITULOS!Indexacao2, XLO_TITULOS!cont_tx_CorrecProRata, DtpEntregaChaves), "##0.0000")
                  XLI_MOEDA2 = XLO_TITULOS!moed_cd_Moeda2
                Else
                  'Se não possuir indexador2
                  XLF_VALORINDEXADO2 = XLF_VALORCORRIGIDO
                  XLI_MOEDA2 = 0
                End If
              End If
                
              XLT_SQL = "UPDATE ConsCAPTitulos SET " & _
                " titu_vl_ParcelaIndex1=" & FunTrataFloat(XLF_VALORINDEXADO1) & _
                ", titu_vl_ParcelaIndex2=" & FunTrataFloat(XLF_VALORINDEXADO2) & _
                ", moed_cd_Moeda1=" & FunNuloBancoZero(XLI_MOEDA1) & _
                ", moed_cd_Moeda2=" & FunNuloBancoZero(XLI_MOEDA2) & _
                ", titu_vl_Parcela=" & FunTrataFloat(XLF_VALORCORRIGIDO) & _
                ", titu_dt_Base = " & FunNuloData(DtpEntregaChaves.Value, NomeSgbd) & _
                ", titu_tx_EntregaChave='2'" & _
                ", titu_vl_PercJurosFin = " & FunTrataFloat(TxtPercJuros.Text) & _
                " WHERE Titulo ='" & XLO_TITULOS!Titulo & "' AND empr_cd_empresa = " & PCodEmpresa & ""
                    
              Conexao.Execute XLT_SQL, , adCmdText
              
              XLT_DESCRICAO = "Empr:" & PCodEmpresa & _
                           " Tit:" & XLO_TITULOS!Titulo & _
                           " VHO:" & XLO_TITULOS!titu_vl_Parcela & _
                           " I1O:" & FunNuloBancoVal(XLO_TITULOS!moed_cd_Moeda1) & _
                           " VI1O:" & XLO_TITULOS!titu_vl_ParcelaIndex1 & _
                           " I2O:" & FunNuloBancoVal(XLO_TITULOS!moed_cd_Moeda2) & _
                           " VI2O:" & XLO_TITULOS!titu_vl_ParcelaIndex2 & _
                           " DBO:" & XLO_TITULOS!titu_dt_Base & _
                           " VHA:" & XLF_VALORCORRIGIDO & _
                           " I1A:" & FunNuloBancoZero(XLI_MOEDA1) & _
                           " VI1A:" & XLF_VALORINDEXADO1 & _
                           " I2A:" & FunNuloBancoZero(XLI_MOEDA2) & _
                           " VI2A:" & XLF_VALORINDEXADO2 & _
                           " DBA:" & Format(DtpEntregaChaves, "dd/mm/yy")
           
              Call subGravaLog(func_cd_operador, "TelaManContratos", "9", Now(), XLT_DESCRICAO)
              
              XLT_TITULO = XLO_TITULOS!Titulo
              
            'Títulos que vão ter o mesmo valor
            Else
              
              'Pega a data do vencimento do primeiro título da série
              If XLD_DATABASE = "01/01/1900" Then
                XLD_DATABASE = DateAdd("m", -CInt(XLO_TITULOS!tipl_nr_Periodicidade), XLO_TITULOS!titu_dt_Vencimento)
              End If
              
              'if XLD_DATABASE < DtpEntregaChaves.Value
               
              If OptIndexador1.Value = True Then
                'Se for corrigir pelo indexador1
                If Not IsNull(XLO_TITULOS!moeda1) Then
                  'Correção Monetária
                  XLF_VALORCORRIGIDO = Format(funCalculaCorrecaoMonetaria(XLO_TITULOS, 1, DtpEntregaChaves), "standard")
                Else
                  XLF_VALORCORRIGIDO = Format(XLO_TITULOS!titu_vl_Parcela, "standard")
                End If
              Else
                'Se for corrigir pelo indexador2
                If Not IsNull(XLO_TITULOS!moeda2) Then
                  'Correção Monetária
                  XLF_VALORCORRIGIDO = Format(funCalculaCorrecaoMonetaria(XLO_TITULOS, 2, DtpEntregaChaves), "standard")
                Else
                  XLF_VALORCORRIGIDO = Format(XLO_TITULOS!titu_vl_Parcela, "standard")
                End If
              End If
              
              XLF_JUROS = FunJuros(XLF_VALORCORRIGIDO, TxtPercJuros.Text, "M", XLD_DATABASE, DtpEntregaChaves.Value, 0)
              XLF_VALORCORRIGIDO = XLF_VALORCORRIGIDO + XLF_JUROS
              XLF_SOMATORIO = XLF_SOMATORIO + XLF_VALORCORRIGIDO
            End If
        
            XLO_TITULOS.MoveNext
          
          Wend
          
          If XLO_TITULOS.RecordCount > 1 Then
          
            XLO_TITULOS.MoveFirst
            
            'Calcula a percentagem de juros
            If XLO_TITULOS!tipl_nr_Periodicidade = 1 Then
              'Se for mensal
              XLF_PERCJUROS = TxtPercJuros.Text / 100
            Else
              'Altera para a taxa mensal
              XLF_PERCJUROS = ((1 + TxtPercJuros.Text / 100) ^ CInt(XLO_TITULOS!tipl_nr_Periodicidade)) - 1
            End If
            
            'Pega o número de parcelas
            If XLT_TITULO <> "" Then
              XLI_PARCELAS = XLO_TITULOS.RecordCount - 1
            Else
              XLI_PARCELAS = XLO_TITULOS.RecordCount
            End If
            
            XLF_VALORCORRIGIDO = XLF_SOMATORIO / ((1 - (1 + XLF_PERCJUROS) ^ (-XLI_PARCELAS)) / XLF_PERCJUROS)
            
            If CboMoeda1.BoundText <> "" Then
              'Se escolher novo indexador1
              XLF_VALORINDEXADO1 = Format(FunValorIndexado(XLF_VALORCORRIGIDO, XLF_COTACAO1, XLF_COTACAO1MESANTERIOR, DatMoeda1.Recordset!moed_tx_Indexacao, XLO_TITULOS!cont_tx_CorrecProRata, DtpEntregaChaves), "##0.0000")
              XLI_MOEDA1 = CboMoeda1.BoundText
            Else
              'Se não escolher novo indexador1
              If Not IsNull(XLO_TITULOS!moed_cd_Moeda1) Then
                'Se já possuir indexador1
                XLF_VALORINDEXADO1 = Format(FunValorIndexado(XLF_VALORCORRIGIDO, XLO_TITULOS!Cotacao1, XLO_TITULOS!Cotacao1_MesAnterior, XLO_TITULOS!Indexacao1, XLO_TITULOS!cont_tx_CorrecProRata, DtpEntregaChaves), "##0.0000")
                XLI_MOEDA1 = XLO_TITULOS!moed_cd_Moeda1
              Else
                'Se não possuir indexador1
                XLF_VALORINDEXADO1 = XLF_VALORCORRIGIDO
                XLI_MOEDA1 = 0
              End If
            End If
            
            If CboMoeda2.BoundText <> "" Then
              'Se escolher novo indexador2
              XLF_VALORINDEXADO2 = Format(FunValorIndexado(XLF_VALORCORRIGIDO, XLF_COTACAO2, XLF_COTACAO2MESANTERIOR, DatMoeda2.Recordset!moed_tx_Indexacao, XLO_TITULOS!cont_tx_CorrecProRata, DtpEntregaChaves), "##0.0000")
              XLI_MOEDA2 = CboMoeda2.BoundText
            Else
              'Se não escolher novo indexador2
              If Not IsNull(XLO_TITULOS!moed_cd_Moeda2) Then
                'Se já possuir indexador2
                XLF_VALORINDEXADO2 = Format(FunValorIndexado(XLF_VALORCORRIGIDO, XLO_TITULOS!Cotacao2, XLO_TITULOS!Cotacao2_MesAnterior, XLO_TITULOS!Indexacao2, XLO_TITULOS!cont_tx_CorrecProRata, DtpEntregaChaves), "##0.0000")
                XLI_MOEDA2 = XLO_TITULOS!moed_cd_Moeda2
              Else
                'Se não possuir indexador2
                XLF_VALORINDEXADO2 = XLF_VALORCORRIGIDO
                XLI_MOEDA2 = 0
              End If
            End If
            
            
            'Alterado em 25/03/10: a consulta original não respeitava o intervalo de títulos informado
            XLT_SQL = "UPDATE ConsCAPTitulos SET " & _
                " titu_vl_ParcelaIndex1=" & FunTrataFloat(XLF_VALORINDEXADO1) & _
                ", titu_vl_ParcelaIndex2=" & FunTrataFloat(XLF_VALORINDEXADO2) & _
                ", moed_cd_Moeda1=" & FunNuloBancoZero(XLI_MOEDA1) & _
                ", moed_cd_Moeda2=" & FunNuloBancoZero(XLI_MOEDA2) & _
                ", titu_vl_Parcela=" & FunTrataFloat(XLF_VALORCORRIGIDO) & _
                ", titu_dt_Base = " & FunNuloData(DtpEntregaChaves.Value, NomeSgbd) & _
                ", titu_tx_EntregaChave='2'" & _
                ", titu_vl_PercJurosFin = " & FunTrataFloat(TxtPercJuros.Text) & _
                " WHERE left(Titulo,15) ='" & Left(XLO_TITULOS!Titulo, 15) & "'" & _
                " AND Titulo <> '" & XLT_TITULO & "'"
                
            If MskCodigoInicial.Text <> "  .   .  " Then
                XLT_SQL = XLT_SQL & _
                " AND Titulo >= '" & Chave & "." & MskCodigoInicial.Text & "'"
            End If
            If MskCodigoFinal.Text <> "  .   .  " Then
               XLT_SQL = XLT_SQL & _
              " AND Titulo <='" & Chave & "." & MskCodigoFinal.Text & "'"
            End If
            
            XLT_SQL = XLT_SQL & _
                " AND titu_dt_Pagamento is null" & _
                " AND titu_dt_Vencimento>= " & FunNuloData(DtpEntregaChaves.Value, NomeSgbd) & " " & _
                " AND empr_cd_empresa = " & PCodEmpresa & ""
                    
            Conexao.Execute XLT_SQL, , adCmdText
            
            XLT_DESCRICAO = "Empr:" & PCodEmpresa & _
                             " Tit:" & XLO_TITULOS!Titulo & _
                             " VHO:" & XLO_TITULOS!titu_vl_Parcela & _
                             " I1O:" & FunNuloBancoVal(XLO_TITULOS!moed_cd_Moeda1) & _
                             " VI1O:" & XLO_TITULOS!titu_vl_ParcelaIndex1 & _
                             " I2O:" & FunNuloBancoVal(XLO_TITULOS!moed_cd_Moeda2) & _
                             " VI2O:" & XLO_TITULOS!titu_vl_ParcelaIndex2 & _
                             " DBO:" & XLO_TITULOS!titu_dt_Base & _
                             " VHA:" & XLF_VALORCORRIGIDO & _
                             " I1A:" & FunNuloBancoZero(XLI_MOEDA1) & _
                             " VI1A:" & XLF_VALORINDEXADO1 & _
                             " I2A:" & FunNuloBancoZero(XLI_MOEDA2) & _
                             " VI2A:" & XLF_VALORINDEXADO2 & _
                             " DBA:" & Format(DtpEntregaChaves, "dd/mm/yy")
             
            Call subGravaLog(func_cd_operador, "TelaManContratos", "9", Now(), XLT_DESCRICAO)
          End If
            
          XLO_TITULOS.Close
        Set XLO_TITULOS = Nothing
        End If
      Next
    End If
    
    Conexao.CommitTrans
    MsgBox "Operação efetuada com sucesso!", vbInformation, "ATENÇÃO"
    Me.MousePointer = vbDefault
      
  End If

End Sub

Private Sub subVerificaCotacoes(XLO_TITULOS As ADODB.Recordset)

  Dim XLT_SQL As String
  Dim XLO_COTACAO As New ADODB.Recordset 'Cotação do novo indexador 1

  'Verifica as cotações para os indexadores dos títulos
        While Not XLO_TITULOS.EOF
          If Not funVerirficaCotacao(XLO_TITULOS, 1, DtpEntregaChaves.Value, True, TelaCotacoes) Then
            XLO_TITULOS.Requery adCmdText
            If Not funVerirficaCotacao(XLO_TITULOS, 1, DtpEntregaChaves.Value, False, TelaCotacoes) Then
              Me.MousePointer = vbDefault
              Exit Sub
            End If
          End If
          
          If Not funVerirficaCotacao(XLO_TITULOS, 2, DtpEntregaChaves, True, TelaCotacoes) Then
            XLO_TITULOS.Requery adCmdText
            If Not funVerirficaCotacao(XLO_TITULOS, 2, DtpEntregaChaves.Value, False, TelaCotacoes) Then
              Me.MousePointer = vbDefault
              Exit Sub
            End If
          End If
          XLO_TITULOS.MoveNext
        Wend
        
        XLO_TITULOS.MoveFirst
        
        'Indexador 1
        If CboMoeda1.BoundText <> "" Then
          'Se for diário
          If DatMoeda1.Recordset!moed_tx_Indexacao = "D" Then
            XLT_SQL = " SELECT cotacoesmoedas.cota_vl_Valor as Cotacao1 FROM cotacoesmoedas where cotacoesmoedas.moed_cd_CodMoeda=" & DatMoeda1.Recordset!moed_cd_CodMoeda & "  and cota_dt_datacotacao=" & FunNuloData(DtpEntregaChaves.Value, NomeSgbd)
            SubQOpenRecordset XLO_COTACAO, XLT_SQL, Estatico
            If Not XLO_COTACAO.EOF Then
              XLF_COTACAO1 = XLO_COTACAO!Cotacao1
              XLO_COTACAO.Close
              Set XLO_COTACAO = Nothing
            Else
              MsgBox "Não existe cotação para a moeda " & CboMoeda1.Text & " para o dia " & DtpEntregaChaves.Value, vbCritical, "ATENÇÃO"
              Me.MousePointer = vbDefault
              XLO_COTACAO.Close
              Set XLO_COTACAO = Nothing
              Exit Sub
            End If
            
          'Se for mensal
          Else
            XLT_SQL = " SELECT cotacoesmoedas.cota_vl_Valor as Cotacao1 FROM cotacoesmoedas where cotacoesmoedas.moed_cd_CodMoeda=" & DatMoeda1.Recordset!moed_cd_CodMoeda & " and cota_dt_datacotacao=" & FunNuloData("01/" & Month(DtpEntregaChaves.Value) & "/" & Year(DtpEntregaChaves.Value), NomeSgbd)
            SubQOpenRecordset XLO_COTACAO, XLT_SQL, Estatico
            
            If Not XLO_COTACAO.EOF Then
              XLF_COTACAO1 = XLO_COTACAO!Cotacao1
              XLO_COTACAO.Close
              Set XLO_COTACAO = Nothing
            Else
              MsgBox "Não existe cotação para a moeda " & CboMoeda1.Text & " para o dia " & Format(DateAdd("m", 1, DtpEntregaChaves.Value), "\01/mm/yy"), vbCritical, "ATENÇÃO"
              Me.MousePointer = vbDefault
              XLO_COTACAO.Close
              Set XLO_COTACAO = Nothing
              Exit Sub
            End If
              
            If XLO_TITULOS!cont_tx_CorrecProRata = "A" Then
              XLT_SQL = " SELECT cotacoesmoedas.cota_vl_Valor as Cotacao1_MesAnterior FROM cotacoesmoedas where cotacoesmoedas.moed_cd_CodMoeda=" & DatMoeda1.Recordset!moed_cd_CodMoeda & " and cota_dt_datacotacao=dateadd(mm,-1," & FunNuloData("01/" & Month(DtpEntregaChaves.Value) & "/" & Year(DtpEntregaChaves.Value), NomeSgbd) & ")"
              SubQOpenRecordset XLO_COTACAO, XLT_SQL, Estatico
              
              If Not XLO_COTACAO.EOF Then
                XLF_COTACAO1MESANTERIOR = XLO_COTACAO!Cotacao1_MesAnterior
                XLO_COTACAO.Close
              Set XLO_COTACAO = Nothing
              Else
                MsgBox "Não existe cotação para a moeda " & CboMoeda1.Text & " para o dia " & Format(DateAdd("m", -1, DtpEntregaChaves.Value), "\01/mm/yy"), vbCritical, "ATENÇÃO"
                Me.MousePointer = vbDefault
                XLO_COTACAO.Close
                Set XLO_COTACAO = Nothing
                Exit Sub
              End If
              
            ElseIf XLO_TITULOS!cont_tx_CorrecProRata = "P" Then
              XLT_SQL = " SELECT cotacoesmoedas.cota_vl_Valor as Cotacao1_MesAnterior FROM cotacoesmoedas where cotacoesmoedas.moed_cd_CodMoeda=" & DatMoeda1.Recordset!moed_cd_CodMoeda & " and cota_dt_datacotacao=dateadd(mm,1," & FunNuloData("01/" & Month(DtpEntregaChaves.Value) & "/" & Year(DtpEntregaChaves.Value), NomeSgbd) & ")"
              SubQOpenRecordset XLO_COTACAO, XLT_SQL, Estatico
              
              If Not XLO_COTACAO.EOF Then
                XLF_COTACAO1MESANTERIOR = XLO_COTACAO!Cotacao1_MesAnterior
                XLO_COTACAO.Close
                Set XLO_COTACAO = Nothing
              Else
                MsgBox "Não existe cotação para a moeda " & CboMoeda1.Text & " para o dia " & Format(DateAdd("m", 1, DtpEntregaChaves.Value), "\01/mm/yy"), vbCritical, "ATENÇÃO"
                Me.MousePointer = vbDefault
                XLO_COTACAO.Close
                Set XLO_COTACAO = Nothing
                Exit Sub
              End If
            End If
          End If
        End If
        
        'Indexador 2
        If CboMoeda2.BoundText <> "" Then
          'Se for diário
          If DatMoeda2.Recordset!moed_tx_Indexacao = "D" Then
            XLT_SQL = " SELECT cotacoesmoedas.cota_vl_Valor as Cotacao2 FROM cotacoesmoedas where cotacoesmoedas.moed_cd_CodMoeda=" & DatMoeda2.Recordset!moed_cd_CodMoeda & "  and cota_dt_datacotacao=" & FunNuloData(DtpEntregaChaves.Value, NomeSgbd)
            SubQOpenRecordset XLO_COTACAO, XLT_SQL, Estatico
            
            If Not XLO_COTACAO.EOF Then
              XLF_COTACAO2 = XLO_COTACAO!Cotacao2
              XLO_COTACAO.Close
              Set XLO_COTACAO = Nothing
            Else
              MsgBox "Não existe cotação para a moeda " & CboMoeda2.Text & " para o dia " & DtpEntregaChaves.Value, vbCritical, "ATENÇÃO"
              Me.MousePointer = vbDefault
              XLO_COTACAO.Close
              Set XLO_COTACAO = Nothing
              Exit Sub
            End If
            
          'Se for mensal
          Else
            XLT_SQL = " SELECT cotacoesmoedas.cota_vl_Valor as Cotacao2 FROM cotacoesmoedas where cotacoesmoedas.moed_cd_CodMoeda=" & DatMoeda2.Recordset!moed_cd_CodMoeda & " and cota_dt_datacotacao=" & FunNuloData("01/" & Month(DtpEntregaChaves.Value) & "/" & Year(DtpEntregaChaves.Value), NomeSgbd)
            SubQOpenRecordset XLO_COTACAO, XLT_SQL, Estatico
            
            If Not XLO_COTACAO.EOF Then
              XLF_COTACAO2 = XLO_COTACAO!Cotacao2
              XLO_COTACAO.Close
              Set XLO_COTACAO = Nothing
            Else
              MsgBox "Não existe cotação para a moeda " & CboMoeda2.Text & " para o dia " & Format(DateAdd("m", 1, DtpEntregaChaves.Value), "\01/mm/yy"), vbCritical, "ATENÇÃO"
              Me.MousePointer = vbDefault
              XLO_COTACAO.Close
              Set XLO_COTACAO = Nothing
              Exit Sub
            End If
              
            If XLO_TITULOS!cont_tx_CorrecProRata = "A" Then
              XLT_SQL = " SELECT cotacoesmoedas.cota_vl_Valor as Cotacao2_MesAnterior FROM cotacoesmoedas where cotacoesmoedas.moed_cd_CodMoeda=" & DatMoeda2.Recordset!moed_cd_CodMoeda & " and cota_dt_datacotacao=dateadd(mm,-1," & FunNuloData("01/" & Month(DtpEntregaChaves.Value) & "/" & Year(DtpEntregaChaves.Value), NomeSgbd) & ")"
              SubQOpenRecordset XLO_COTACAO, XLT_SQL, Estatico
              
              If Not XLO_COTACAO.EOF Then
                XLF_COTACAO2MESANTERIOR = XLO_COTACAO!Cotacao2_MesAnterior
                XLO_COTACAO.Close
                Set XLO_COTACAO = Nothing
              Else
                MsgBox "Não existe cotação para a moeda " & CboMoeda2.Text & " para o dia " & Format(DateAdd("m", -1, DtpEntregaChaves.Value), "\01/mm/yy"), vbCritical, "ATENÇÃO"
                Me.MousePointer = vbDefault
                XLO_COTACAO.Close
                Set XLO_COTACAO = Nothing
                Exit Sub
              End If
              
            ElseIf XLO_TITULOS!cont_tx_CorrecProRata = "P" Then
              XLT_SQL = " SELECT cotacoesmoedas.cota_vl_Valor as Cotacao2_MesAnterior FROM cotacoesmoedas where cotacoesmoedas.moed_cd_CodMoeda=" & DatMoeda2.Recordset!moed_cd_CodMoeda & " and cota_dt_datacotacao=dateadd(mm,1," & FunNuloData("01/" & Month(DtpEntregaChaves.Value) & "/" & Year(DtpEntregaChaves.Value), NomeSgbd) & ")"
              SubQOpenRecordset XLO_COTACAO, XLT_SQL, Estatico
              
              If Not XLO_COTACAO.EOF Then
                XLF_COTACAO2MESANTERIOR = XLO_COTACAO!Cotacao2_MesAnterior
                XLO_COTACAO.Close
                Set XLO_COTACAO = Nothing
              Else
                MsgBox "Não existe cotação para a moeda " & CboMoeda2.Text & " para o dia " & Format(DateAdd("m", 1, DtpEntregaChaves.Value), "\01/mm/yy"), vbCritical, "ATENÇÃO"
                Me.MousePointer = vbDefault
                XLO_COTACAO.Close
                Set XLO_COTACAO = Nothing
                Exit Sub
              End If
            End If
          End If
        End If

End Sub

Private Sub CmdLimparIndex1_Click()

  CboMoeda1.BoundText = ""

End Sub

Private Sub CmdLimparIndex2_Click()

  CboMoeda2.BoundText = ""

End Sub

Private Sub CmdRetornarAlteracaoIndexador_Click()

  Unload Me
  
End Sub

Private Sub Form_Load()
  
  Dim XLO_ENTREGACHAVES As New ADODB.Recordset
  
  subConectarControleDadosNV DatMoeda1, "SELECT * FROM Moedas ORDER BY moed_tx_Descricao", Estatico
  subConectarControleDadosNV DatMoeda2, "SELECT * FROM Moedas ORDER BY moed_tx_Descricao", Estatico

  SubQOpenRecordset XLO_ENTREGACHAVES, "SELECT imov_dt_EntregaChaves FROM Imoveis WHERE imov_cd_Imovel='" & Mid(Chave, 6, 4) & "'", Estatico
  
  If IsNull(XLO_ENTREGACHAVES!imov_dt_EntregaChaves) Then
    DtpEntregaChaves.Value = Date
  Else
    DtpEntregaChaves.Value = XLO_ENTREGACHAVES!imov_dt_EntregaChaves
  End If
  
  XLO_ENTREGACHAVES.Close
  Set XLO_ENTREGACHAVES = Nothing
  
End Sub

Private Sub MskCodigoFinal_LostFocus()
  
  Dim XLO_TITULO As New ADODB.Recordset
  
  If MskCodigoFinal.Text <> "  .   .  " Then
    MskCodigoFinal = UCase(MskCodigoFinal)
    SubQOpenRecordset XLO_TITULO, "SELECT * FROM ConsCAPTitulos WHERE Titulo = '" & Chave & "." & MskCodigoFinal.Text & "'", Estatico
      If XLO_TITULO.EOF Then
        MsgBox "O título Final não existe!", vbCritical, "ATENÇÃO !"
        MskCodigoFinal.SetFocus
      End If
    XLO_TITULO.Close
    Set XLO_TITULO = Nothing
  End If
   
End Sub

Private Sub MskCodigoInicial_LostFocus()
  
  Dim XLO_TITULO As New ADODB.Recordset
  
  If MskCodigoInicial.Text <> "  .   .  " Then
    MskCodigoInicial = UCase(MskCodigoInicial.Text)
    SubQOpenRecordset XLO_TITULO, "SELECT * FROM ConsCAPTitulos WHERE Titulo = '" & Chave & "." & MskCodigoInicial.Text & "'", Estatico
      If XLO_TITULO.EOF Then
        MsgBox "O título Inicial não existe!", vbCritical, "ATENÇÃO !"
        MskCodigoInicial.SetFocus
      Else
        If MskCodigoFinal.Text = "  .   .  " Then
          MskCodigoFinal.Text = MskCodigoInicial.Text
        End If
      End If
    XLO_TITULO.Close
    Set XLO_TITULO = Nothing
  End If
  
End Sub

Private Sub TxtPercJuros_KeyPress(KeyAscii As Integer)
  
  If KeyAscii = 46 Then KeyAscii = 44
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 44 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
  
End Sub

Private Sub TxtPercJuros_LostFocus()
   
   If IsNumeric(TxtPercJuros.Text) Then
     TxtPercJuros.Text = Format$(TxtPercJuros.Text, "##,##0.0000")
   Else
     TxtPercJuros.Text = ""
   End If
   
End Sub
