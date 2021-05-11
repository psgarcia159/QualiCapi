VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form TelaMoedas 
   Caption         =   "Tabela de Indexadores"
   ClientHeight    =   4830
   ClientLeft      =   195
   ClientTop       =   615
   ClientWidth     =   9435
   Icon            =   "TelaMoedas.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4830
   ScaleWidth      =   9435
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8880
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Threed.SSPanel PanMoeda 
      Height          =   1995
      Left            =   1560
      TabIndex        =   12
      Top             =   1200
      Visible         =   0   'False
      Width           =   5775
      _Version        =   65536
      _ExtentX        =   10186
      _ExtentY        =   3519
      _StockProps     =   15
      Caption         =   " Cadastro de Indexadores"
      ForeColor       =   16711680
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      BevelInner      =   1
      Alignment       =   0
      Begin VB.OptionButton OptMensal 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Mensal"
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
         Left            =   2520
         TabIndex        =   19
         Top             =   1350
         Width           =   975
      End
      Begin VB.OptionButton OptDiario 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Diário"
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
         Left            =   1560
         TabIndex        =   18
         Top             =   1350
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.CommandButton CmdCancelar 
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
         Height          =   375
         Left            =   4635
         TabIndex        =   21
         Top             =   1440
         Width           =   975
      End
      Begin VB.CommandButton CmdConfirmar 
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
         Height          =   375
         Left            =   3600
         TabIndex        =   20
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox Txtsimbolo 
         Height          =   285
         Left            =   1560
         MaxLength       =   5
         TabIndex        =   16
         Top             =   945
         Width           =   915
      End
      Begin VB.TextBox TxtDescricao 
         Height          =   285
         Left            =   1560
         MaxLength       =   30
         TabIndex        =   14
         Top             =   480
         Width           =   4035
      End
      Begin VB.Label LblMovimento 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Movimento:"
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
         Left            =   540
         TabIndex        =   17
         Top             =   1320
         Width           =   990
      End
      Begin VB.Label LblSimbolo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Símbolo:"
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
         Left            =   540
         TabIndex        =   15
         Top             =   960
         Width           =   990
      End
      Begin VB.Label LblDescricao 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Descrição:"
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
         Left            =   540
         TabIndex        =   13
         Top             =   495
         Width           =   990
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   6645
      Top             =   3300
      Visible         =   0   'False
      Width           =   1950
      _ExtentX        =   3440
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
      Caption         =   "Adodc1"
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
   Begin Threed.SSPanel PanBotoes 
      Height          =   360
      Left            =   60
      TabIndex        =   22
      Top             =   4140
      Width           =   9360
      _Version        =   65536
      _ExtentX        =   16510
      _ExtentY        =   635
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   0
      Begin VB.CommandButton cmdExport 
         Height          =   330
         Left            =   8880
         Picture         =   "TelaMoedas.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   0
         Width           =   375
      End
      Begin VB.CommandButton CmdPadrao 
         Caption         =   "Pa&drão"
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
         Left            =   5868
         TabIndex        =   9
         Top             =   30
         Width           =   930
      End
      Begin VB.CommandButton CmdAlterar 
         BackColor       =   &H00000000&
         Caption         =   "&Alterar"
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
         Left            =   978
         TabIndex        =   4
         Top             =   30
         Width           =   930
      End
      Begin VB.CommandButton CmdInserir 
         BackColor       =   &H00000000&
         Caption         =   "&Inserir"
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
         Left            =   0
         TabIndex        =   3
         Top             =   30
         Width           =   930
      End
      Begin VB.CommandButton CmdExcluir 
         Caption         =   "&Excluir"
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
         Left            =   1956
         TabIndex        =   5
         Top             =   30
         Width           =   930
      End
      Begin VB.CommandButton CmdRemoverFiltro 
         Caption         =   "Se&leções"
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
         Left            =   2934
         TabIndex        =   6
         Top             =   30
         Width           =   930
      End
      Begin VB.CommandButton CmdOrdem 
         Caption         =   "&Ordem"
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
         Left            =   3912
         TabIndex        =   7
         Top             =   30
         Width           =   930
      End
      Begin VB.CommandButton CmdImprimir 
         Caption         =   "Im&primir"
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
         Left            =   6846
         TabIndex        =   10
         Top             =   30
         Width           =   930
      End
      Begin VB.CommandButton CmdSair 
         Caption         =   "&Sair"
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
         Left            =   7800
         TabIndex        =   11
         Top             =   30
         Width           =   930
      End
      Begin VB.CommandButton CmdCotacoes 
         Caption         =   "Co&tações"
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
         Left            =   4890
         TabIndex        =   8
         Top             =   30
         Width           =   930
      End
   End
   Begin MSDataListLib.DataCombo cboLocalizar 
      Bindings        =   "TelaMoedas.frx":04B9
      Height          =   315
      Left            =   2760
      TabIndex        =   2
      Top             =   3780
      Width           =   5160
      _ExtentX        =   9102
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      ListField       =   "moed_tx_descricao"
      BoundColumn     =   "moed_cd_codmoeda"
      Text            =   ""
   End
   Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
      Bindings        =   "TelaMoedas.frx":04CE
      Height          =   3690
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   8805
      _ExtentX        =   15531
      _ExtentY        =   6509
      _LayoutType     =   4
      _RowHeight      =   23
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Código"
      Columns(0).DataField=   "moed_cd_codmoeda"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Descrição"
      Columns(1).DataField=   "moed_tx_descricao"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Símbolo"
      Columns(2).DataField=   "moed_tx_simbolo"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Indexação"
      Columns(3).DataField=   "moed_tx_indexacao"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   80
      Columns(4)._MaxComboItems=   5
      Columns(4).ValueItems(0)._DefaultItem=   0
      Columns(4).ValueItems(0).Value=   "S"
      Columns(4).ValueItems(0).Value.vt=   8
      Columns(4).ValueItems(0).DisplayValue.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
      Columns(4).ValueItems(0).DisplayValue(0)=   "bHQAADYFAABCTTYFAAAAAAAANgAAACgAAAAVAAAAFAAAAAEAGAAAAAAAAAUAAAAAAAAAAAAAAAAA"
      Columns(4).ValueItems(0).DisplayValue(1)=   "AAAAAAC/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/"
      Columns(4).ValueItems(0).DisplayValue(2)=   "v7+/v7+/v7+/v78Av7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/"
      Columns(4).ValueItems(0).DisplayValue(3)=   "v7+/v7+/v7+/v7+/v7+/v7+/AL+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/"
      Columns(4).ValueItems(0).DisplayValue(4)=   "v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/vwC/v7+/v7+/v7+/v7+/v79/f39/f39/f39/f39/f39/"
      Columns(4).ValueItems(0).DisplayValue(5)=   "f39/f39/f39/f39/f39/f39/f3+/v7+/v7+/v7+/v78Av7+/v7+/v7+/v7+/////v7+/v7+/v7+/"
      Columns(4).ValueItems(0).DisplayValue(6)=   "v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/f39/v7+/v7+/v7+/v7+/AL+/v7+/v7+/v7+/v////7+/"
      Columns(4).ValueItems(0).DisplayValue(7)=   "vwAAAAAAAL+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v39/f7+/v7+/v7+/v7+/vwC/v7+/v7+/v7+/"
      Columns(4).ValueItems(0).DisplayValue(8)=   "v7////+/v78AAAAAAAAAAAC/v7+/v7+/v7+/v7+/v7+/v7+/v79/f3+/v7+/v7+/v7+/v78Av7+/"
      Columns(4).ValueItems(0).DisplayValue(9)=   "v7+/v7+/v7+/////v7+/AAAAAAAAAAAAAAAAv7+/v7+/v7+/v7+/v7+/v7+/f39/v7+/v7+/v7+/"
      Columns(4).ValueItems(0).DisplayValue(10)=   "v7+/AL+/v7+/v7+/v7+/v////7+/vwAAAAAAAL+/vwAAAAAAAL+/v7+/v7+/v7+/v7+/v39/f7+/"
      Columns(4).ValueItems(0).DisplayValue(11)=   "v7+/v7+/v7+/vwC/v7+/v7+/v7+/v7////+/v78AAAAAAAC/v7+/v78AAAAAAAC/v7+/v7+/v7+/"
      Columns(4).ValueItems(0).DisplayValue(12)=   "v79/f3+/v7+/v7+/v7+/v78Av7+/v7+/v7+/v7+/////v7+/v7+/AAAAv7+/v7+/v7+/AAAAAAAA"
      Columns(4).ValueItems(0).DisplayValue(13)=   "v7+/v7+/v7+/f39/v7+/v7+/v7+/v7+/AL+/v7+/v7+/v7+/v////7+/v7+/v7+/v7+/v7+/v7+/"
      Columns(4).ValueItems(0).DisplayValue(14)=   "v7+/vwAAAAAAAL+/v7+/v39/f7+/v7+/v7+/v7+/vwC/v7+/v7+/v7+/v7////+/v7+/v7+/v7+/"
      Columns(4).ValueItems(0).DisplayValue(15)=   "v7+/v7+/v7+/v7+/v78AAAAAAAC/v79/f3+/v7+/v7+/v7+/v78Av7+/v7+/v7+/v7+/////v7+/"
      Columns(4).ValueItems(0).DisplayValue(16)=   "v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/f39/v7+/v7+/v7+/v7+/AL+/v7+/v7+/v7+/"
      Columns(4).ValueItems(0).DisplayValue(17)=   "v////7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v39/f7+/v7+/v7+/v7+/vwC/v7+/"
      Columns(4).ValueItems(0).DisplayValue(18)=   "v7+/v7+/v7////////////////////////////////////////////////+/v7+/v7+/v7+/v7+/"
      Columns(4).ValueItems(0).DisplayValue(19)=   "v78Av7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/"
      Columns(4).ValueItems(0).DisplayValue(20)=   "v7+/v7+/v7+/AL+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/"
      Columns(4).ValueItems(0).DisplayValue(21)=   "v7+/v7+/v7+/v7+/v7+/vwC/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/"
      Columns(4).ValueItems(0).DisplayValue(22)=   "v7+/v7+/v7+/v7+/v7+/v7+/v7+/v78Av7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/"
      Columns(4).ValueItems(0).DisplayValue(23)=   "v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/AA=="
      Columns(4).ValueItems(0).DisplayValue.vt=   9
      Columns(4).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
      Columns(4).ValueItems(1)._DefaultItem=   0
      Columns(4).ValueItems(1).Value=   "N"
      Columns(4).ValueItems(1).Value.vt=   8
      Columns(4).ValueItems(1).DisplayValue.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
      Columns(4).ValueItems(1).DisplayValue(0)=   "bHQAADYFAABCTTYFAAAAAAAANgAAACgAAAAVAAAAFAAAAAEAGAAAAAAAAAUAAAAAAAAAAAAAAAAA"
      Columns(4).ValueItems(1).DisplayValue(1)=   "AAAAAAC/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/"
      Columns(4).ValueItems(1).DisplayValue(2)=   "v7+/v7+/v7+/v78Av7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/"
      Columns(4).ValueItems(1).DisplayValue(3)=   "v7+/v7+/v7+/v7+/v7+/v7+/AL+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/"
      Columns(4).ValueItems(1).DisplayValue(4)=   "v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/vwC/v7+/v7+/v7+/v7+/v79/f39/f39/f39/f39/f39/"
      Columns(4).ValueItems(1).DisplayValue(5)=   "f39/f39/f39/f39/f39/f39/f3+/v7+/v7+/v7+/v78Av7+/v7+/v7+/v7+/////v7+/v7+/v7+/"
      Columns(4).ValueItems(1).DisplayValue(6)=   "v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/f39/v7+/v7+/v7+/v7+/AL+/v7+/v7+/v7+/v////7+/"
      Columns(4).ValueItems(1).DisplayValue(7)=   "v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v39/f7+/v7+/v7+/v7+/vwC/v7+/v7+/v7+/"
      Columns(4).ValueItems(1).DisplayValue(8)=   "v7////+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v79/f3+/v7+/v7+/v7+/v78Av7+/"
      Columns(4).ValueItems(1).DisplayValue(9)=   "v7+/v7+/v7+/////v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/f39/v7+/v7+/v7+/"
      Columns(4).ValueItems(1).DisplayValue(10)=   "v7+/AL+/v7+/v7+/v7+/v////7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v39/f7+/"
      Columns(4).ValueItems(1).DisplayValue(11)=   "v7+/v7+/v7+/vwC/v7+/v7+/v7+/v7////+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/"
      Columns(4).ValueItems(1).DisplayValue(12)=   "v79/f3+/v7+/v7+/v7+/v78Av7+/v7+/v7+/v7+/////v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/"
      Columns(4).ValueItems(1).DisplayValue(13)=   "v7+/v7+/v7+/f39/v7+/v7+/v7+/v7+/AL+/v7+/v7+/v7+/v////7+/v7+/v7+/v7+/v7+/v7+/"
      Columns(4).ValueItems(1).DisplayValue(14)=   "v7+/v7+/v7+/v7+/v7+/v39/f7+/v7+/v7+/v7+/vwC/v7+/v7+/v7+/v7////+/v7+/v7+/v7+/"
      Columns(4).ValueItems(1).DisplayValue(15)=   "v7+/v7+/v7+/v7+/v7+/v7+/v7+/v79/f3+/v7+/v7+/v7+/v78Av7+/v7+/v7+/v7+/////v7+/"
      Columns(4).ValueItems(1).DisplayValue(16)=   "v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/f39/v7+/v7+/v7+/v7+/AL+/v7+/v7+/v7+/"
      Columns(4).ValueItems(1).DisplayValue(17)=   "v////7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v39/f7+/v7+/v7+/v7+/vwC/v7+/"
      Columns(4).ValueItems(1).DisplayValue(18)=   "v7+/v7+/v7////////////////////////////////////////////////+/v7+/v7+/v7+/v7+/"
      Columns(4).ValueItems(1).DisplayValue(19)=   "v78Av7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/"
      Columns(4).ValueItems(1).DisplayValue(20)=   "v7+/v7+/v7+/AL+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/"
      Columns(4).ValueItems(1).DisplayValue(21)=   "v7+/v7+/v7+/v7+/v7+/vwC/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/"
      Columns(4).ValueItems(1).DisplayValue(22)=   "v7+/v7+/v7+/v7+/v7+/v7+/v7+/v78Av7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/"
      Columns(4).ValueItems(1).DisplayValue(23)=   "v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/v7+/AA=="
      Columns(4).ValueItems(1).DisplayValue.vt=   9
      Columns(4).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
      Columns(4).ValueItems.Count=   2
      Columns(4).Caption=   "Padrão"
      Columns(4).DataField=   "moed_tx_padrao"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   5
      Splits(0)._UserFlags=   0
      Splits(0).ExtendRightColumn=   -1  'True
      Splits(0).MarqueeStyle=   3
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).AllowColMove=   -1  'True
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=5"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=532"
      Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
      Splits(0)._ColumnProps(7)=   "Column(0).WrapText=1"
      Splits(0)._ColumnProps(8)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(9)=   "Column(0)._MinWidth=49"
      Splits(0)._ColumnProps(10)=   "Column(1).Width=7382"
      Splits(0)._ColumnProps(11)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(12)=   "Column(1)._WidthInPix=7303"
      Splits(0)._ColumnProps(13)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(14)=   "Column(1)._ColStyle=528"
      Splits(0)._ColumnProps(15)=   "Column(1).WrapText=1"
      Splits(0)._ColumnProps(16)=   "Column(1).AllowFocus=0"
      Splits(0)._ColumnProps(17)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(18)=   "Column(1)._MinWidth=55313876"
      Splits(0)._ColumnProps(19)=   "Column(2).Width=1588"
      Splits(0)._ColumnProps(20)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(21)=   "Column(2)._WidthInPix=1508"
      Splits(0)._ColumnProps(22)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(23)=   "Column(2)._ColStyle=528"
      Splits(0)._ColumnProps(24)=   "Column(2).WrapText=1"
      Splits(0)._ColumnProps(25)=   "Column(2).AllowFocus=0"
      Splits(0)._ColumnProps(26)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(27)=   "Column(2)._MinWidth=55313268"
      Splits(0)._ColumnProps(28)=   "Column(3).Width=1905"
      Splits(0)._ColumnProps(29)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(30)=   "Column(3)._WidthInPix=1826"
      Splits(0)._ColumnProps(31)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(32)=   "Column(3)._ColStyle=529"
      Splits(0)._ColumnProps(33)=   "Column(3).WrapText=1"
      Splits(0)._ColumnProps(34)=   "Column(3).AllowFocus=0"
      Splits(0)._ColumnProps(35)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(36)=   "Column(4).Width=1402"
      Splits(0)._ColumnProps(37)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(38)=   "Column(4)._WidthInPix=1323"
      Splits(0)._ColumnProps(39)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(40)=   "Column(4)._ColStyle=529"
      Splits(0)._ColumnProps(41)=   "Column(4).WrapText=1"
      Splits(0)._ColumnProps(42)=   "Column(4).Order=5"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowUpdate     =   0   'False
      DefColWidth     =   0
      HeadLines       =   2
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      DeadAreaBackColor=   12632256
      RowDividerColor =   12632256
      RowSubDividerColor=   12632256
      DirectionAfterEnter=   1
      MaxRows         =   250000
      ViewColumnCaptionWidth=   0
      ViewColumnWidth =   0
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=-1,.fontsize=825,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
      _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1,.valignment=2,.wraptext=-1"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.alignment=2"
      _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.alignment=0"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14,.alignment=2"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=46,.parent=13,.alignment=0"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14,.alignment=2"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=50,.parent=13,.alignment=2"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14,.alignment=2"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=54,.parent=13,.alignment=2"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
      _StyleDefs(56)  =   "Named:id=33:Normal"
      _StyleDefs(57)  =   ":id=33,.parent=0"
      _StyleDefs(58)  =   "Named:id=34:Heading"
      _StyleDefs(59)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(60)  =   ":id=34,.wraptext=-1"
      _StyleDefs(61)  =   "Named:id=35:Footing"
      _StyleDefs(62)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(63)  =   "Named:id=36:Selected"
      _StyleDefs(64)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(65)  =   "Named:id=37:Caption"
      _StyleDefs(66)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(67)  =   "Named:id=38:HighlightRow"
      _StyleDefs(68)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(69)  =   "Named:id=39:EvenRow"
      _StyleDefs(70)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(71)  =   "Named:id=40:OddRow"
      _StyleDefs(72)  =   ":id=40,.parent=33"
      _StyleDefs(73)  =   "Named:id=41:RecordSelector"
      _StyleDefs(74)  =   ":id=41,.parent=34"
      _StyleDefs(75)  =   "Named:id=42:FilterBar"
      _StyleDefs(76)  =   ":id=42,.parent=33"
   End
   Begin VB.Label lblLocalizar 
      Alignment       =   1  'Right Justify
      Caption         =   "Localizar Descrição:"
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   15
      TabIndex        =   1
      Top             =   3795
      Width           =   2670
   End
End
Attribute VB_Name = "TelaMoedas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim errLoop As ADODB.Error
Dim StrError As String

'*****************************************************************************************
'Variaveis que guardarão valores específicos de cada formulário
Dim PrimeiraVez As Boolean

Dim FRM_ContFil As Integer          ' Contador que armazena o número de filtros ativos
Dim FRM_Ordem As String             ' Nome do Campo que tem a chave da Tabela ativa
Dim FRM_Campos_Obrig As String      ' Lista dos campos que são obrigatórios na tabela
Dim FRM_TipoVar As String           ' Tipo do campo a ser filtrado
Dim FRM_FiltroOrig As String        ' Indica os filtros determinados pelo próprio sistema
Dim FRM_FiltroAtual As String       ' Indica o filtro atual da tabela corrente
Dim FRM_filtra_especial As Integer  ' Indica se a tabela utilizada possui filtros pré-definidos pelo sistema em FiltroOrig
Dim FRM_Tem_Filtro As Boolean       ' Informa se já foi relizado um filtro na tabela pelo usuário
Dim FRM_ConTelaOrdem As Integer     ' Indica o número de chaves utilizados para ordenação
Dim FRM_Quantid As Long             ' Quantidade de registros da tabela
Dim FRM_Filtrou As Boolean          ' Informa se Algum campo foi filtrado
Dim FRM_Ordens() As Integer         ' Vetor que mantem a ordem dos campos
Dim FRM_Filtros() As String         ' Array para armazenar os Filtros da Tabela

Dim FRM_BarraFerramentasPressionada As Boolean
Dim FRM_VariasVezes As Boolean      'Indica se vai executar a insercao na TelaPadrao varias ou uma vez
Dim FRM_XInserir As String          'Flag para identificar se entrou na rotina para inserir ou alterar ('I'Inserir, 'A'Alterar, 'C'Consultar)
Dim FRM_XOrigem As String           'Identifica a origem (que formulario) chamou o formulário ativo
Dim FRM_SQLAtivo As String          'Atualizar o SQL Ativo após modificações de filtro e ordem
Sub subTelaValoresGlobais(GuardaOUPegaValores As String)
    Dim Count As Integer, Count2 As Integer
    If GuardaOUPegaValores = "G" Then 'Guardando variáveis de definição da tela
        FRM_ContFil = ContFil
        FRM_Ordem = Ordem
        FRM_Campos_Obrig = Campos_Obrig
        FRM_TipoVar = TipoVar
        FRM_FiltroOrig = FiltroOrig
        FRM_FiltroAtual = FiltroAtual
        FRM_filtra_especial = filtra_especial
        FRM_Tem_Filtro = Tem_Filtro
        FRM_ConTelaOrdem = ConTelaOrdem
        FRM_Quantid = Quantid
        FRM_Filtrou = Filtrou
        FRM_BarraFerramentasPressionada = BarraFerramentasPressionada
        FRM_VariasVezes = VariasVezes
        FRM_XInserir = XInserir
        FRM_XOrigem = XOrigem
        FRM_SQLAtivo = SQLAtivo
        xLocalDescricao = xLocalDescricao
        xLocalCampo = Ordem
        
        ReDim FRM_Filtros(UBound(Filtros, 1), UBound(Filtros, 2))
        ReDim FRM_Ordens(UBound(Ordens, 1)) As Integer
        
        For Count = LBound(Filtros, 1) To UBound(Filtros, 1)
            For Count2 = LBound(Filtros, 2) To UBound(Filtros, 2)
                FRM_Filtros(Count, Count2) = Filtros(Count, Count2)
            Next Count2
        Next Count
        For Count = LBound(Ordens) To UBound(Ordens)
            FRM_Ordens(Count) = Ordens(Count)
        Next Count
    ElseIf GuardaOUPegaValores = "P" Then  'Regravando os valores relativos a tela atual variáveis de definição da tela
        ContFil = FRM_ContFil
        Ordem = FRM_Ordem
        Campos_Obrig = FRM_Campos_Obrig
        TipoVar = FRM_TipoVar
        FiltroOrig = FRM_FiltroOrig
        FiltroAtual = FRM_FiltroAtual
        filtra_especial = FRM_filtra_especial
        Tem_Filtro = FRM_Tem_Filtro
        ConTelaOrdem = FRM_ConTelaOrdem
        Quantid = FRM_Quantid
        Filtrou = FRM_Filtrou
        BarraFerramentasPressionada = FRM_BarraFerramentasPressionada
        VariasVezes = FRM_VariasVezes
        XInserir = FRM_XInserir
        XOrigem = FRM_XOrigem
        SQLAtivo = FRM_SQLAtivo
        
        ReDim Filtros(UBound(FRM_Filtros, 1), UBound(FRM_Filtros, 2))
        ReDim Ordens(UBound(FRM_Ordens, 1))
        
        For Count = LBound(FRM_Filtros, 1) To UBound(FRM_Filtros, 1)
            For Count2 = LBound(FRM_Filtros, 2) To UBound(FRM_Filtros, 2)
                Filtros(Count, Count2) = FRM_Filtros(Count, Count2)
            Next Count2
        Next Count
        For Count = LBound(FRM_Ordens) To UBound(FRM_Ordens)
            Ordens(Count) = FRM_Ordens(Count)
        Next Count
    End If
End Sub
Private Sub subDesabilitaBotoes()
    'Objetivo: Desabilita os Botões da Tela
    CmdInserir.Enabled = False
    CmdAlterar.Enabled = False
    CmdCotacoes.Enabled = False
    CmdExcluir.Enabled = False
    CmdRemoverFiltro.Enabled = False
    CmdImprimir.Enabled = False
    CmdOrdem.Enabled = False
'    CmdPadrao.Enabled = False
    CmdSair.Enabled = False
End Sub
Private Sub cboLocalizar_Click(Area As Integer)
    If Area = 2 Then
        Adodc1.Recordset.Bookmark = cboLocalizar.SelectedItem
    End If
End Sub
Private Sub CmdCancelar_Click()
    If Not FunTabelaVazia(Adodc1) Then Adodc1.Recordset.CancelUpdate
    PanBotoes.Enabled = True
    PanMoeda.Visible = False
    PanMoeda.Enabled = True
    'CmdConfirmar.Caption = "Confirmar"
    subHabilitaBotoes
'    DataGrid1.Height = 4260
    TDBGrid1.Enabled = True
    TDBGrid1.SetFocus
End Sub
Private Sub CmdConfirmar_Click()
    Dim Sql As String
    Dim nreg As Variant
    Dim XStatus As String
    Dim CodEquip As Integer
    Dim IntLoop As Integer
    Dim x As String
    Dim XLO_MOEDAS As ADODB.Recordset
    Dim XLT_MOVIMENTO As String
    
    On Error GoTo MostraErro
    
    'Função para verificar preenchimento dos campos obrigatórios
    
    If FunObrigatorioTXT(TxtDescricao, "A Descricao é Obrigatória") Then Exit Sub
    'If FunObrigatorioCBOX(cmbTipoIndexador, "O Tipo de Indexação é Obrigatório") Then Exit Sub
    'If funObrigatorioTXT(Txtindexacao, "O Tipo de Indexação é Obrigatório") Then Exit Sub
    'Txtindexacao.Text = UCase(Txtindexacao.Text)
    'If Txtindexacao.Text <> "M" And Txtindexacao.Text <> "D" Then
    '    MsgBox "O tipo de Indexação deve ser (D)Diário ou (M)Mensal!", vbCritical, "ATENÇÃO"
    '    Txtindexacao.SetFocus
    '    Exit Sub
    'End If
    If OptDiario.Value Then
        XLT_MOVIMENTO = "D"
    Else
        XLT_MOVIMENTO = "M"
    End If
    Conexao.BeginTrans
    If PanMoeda.Tag = "I" Then 'Inclusão
        XGT_SQL = "SELECT moed_cd_CodMoeda FROM Moedas"
        SubQOpenRecordset XLO_MOEDAS, XGT_SQL, Estatico
        If XLO_MOEDAS.RecordCount > 0 Then
            XGT_SQL = "INSERT INTO Moedas (moed_tx_simbolo,moed_tx_descricao,moed_tx_indexacao, moed_tx_padrao) " & _
                      "VALUES ('" & FunNulo(Txtsimbolo.Text) & "','" & FunNulo(TxtDescricao.Text) & _
                      "','" & XLT_MOVIMENTO & "','N')"
        Else
            XGT_SQL = "INSERT INTO Moedas (moed_tx_simbolo,moed_tx_descricao,moed_tx_indexacao, moed_tx_padrao) " & _
                      "VALUES ('" & FunNulo(Txtsimbolo.Text) & "','" & FunNulo(TxtDescricao.Text) & _
                      "','" & XLT_MOVIMENTO & "','S')"
        End If
        Conexao.Execute XGT_SQL
        XLO_MOEDAS.Close
        
        MdiPrincipal.BarraStatus.Panels(3).Text = CStr((Adodc1.Recordset.RecordCount + 1)) + "/" + CStr((Adodc1.Recordset.RecordCount + 1))
    Else    'Alteração
        XGT_SQL = "UPDATE Moedas " & _
              "SET moed_tx_simbolo = '" & FunNulo(Txtsimbolo.Text) & "', " & _
              "    moed_tx_descricao = '" & FunNulo(TxtDescricao.Text) & "'," & _
              "    moed_tx_indexacao = '" & XLT_MOVIMENTO & "' " & _
              "WHERE moed_cd_codmoeda = " & CInt(PanMoeda.Tag)
        Conexao.Execute XGT_SQL
    End If
    Conexao.CommitTrans
    
    If PanMoeda.Tag <> "I" Then
        FRM_Quantid = funOperacaoQuantid("+")
    End If
    subrecarregadadosNV Adodc1, "", FRM_Ordem, "", FRM_FiltroAtual
    'Adodc1.Refresh TestesubrecarregadadosNV
    If PanMoeda.Tag <> "I" Then
        'subLocalizaData Adodc1, "moed_cd_codmoeda", CLng(PanMoeda.Tag)
        cboLocalizar.BoundText = PanMoeda.Tag
        If Not IsNull(cboLocalizar.SelectedItem) Then
            Adodc1.Recordset.Bookmark = cboLocalizar.SelectedItem
        End If
    End If
    
    PanMoeda.Visible = False
    PanMoeda.Enabled = False
    PanBotoes.Enabled = True
    subHabilitaBotoes
    
    TDBGrid1.Enabled = True
    
    Exit Sub
    
MostraErro:

    Conexao.RollbackTrans

    ' Enumera a coleção de erros e apresenta
    ' as propriedades de cada erro ocorrido
    For Each errLoop In Conexao.Errors
        StrError = "Error #" & errLoop.Number & vbCr & _
            "   " & errLoop.Description & vbCr & _
            "   (Source: " & errLoop.Source & ")" & vbCr & _
            "   (SQL State: " & errLoop.SQLState & ")" & vbCr & _
            "   (NativeError: " & errLoop.NativeError & ")" & vbCr

    MsgBox StrError, vbCritical + vbOKOnly, "ATENÇÃO !!!"
    Next
End Sub
Private Sub CmdCotacoes_Click()
    subDesabilitaBotoes
    XFormulario = "TelaMoedas"
    TelaCotacoes.LblChave.Caption = Adodc1.Recordset.Fields("moed_cd_codmoeda").Value
    'TelaCotacoes.XPT_NomeMoeda = Adodc1.Recordset.Fields("moed_tx_descricao").Value
    'TelaCotacoes.XPT_IndexMoeda = Adodc1.Recordset.Fields("moed_tx_Indexacao").Value
    TelaCotacoes.Show 1
    subHabilitaBotoes
End Sub

Private Sub cmdExport_Click()
    Dim sLine As String
    Dim tmpFile As String
    Dim outFile As String
    Dim i As Integer
    
    ' Sets the Dialog Title to Save File
    CommonDialog1.DialogTitle = "Salvar arquivo"

    ' Sets the File List box to Text File and All Files
    CommonDialog1.Filter = "Arquivo CSV (*.csv)|*.csv"

    ' Set the default files type to Text File
    CommonDialog1.FilterIndex = 1

    ' Sets the flags - Hide Read only, prompt to overwrite, and path must exist
    CommonDialog1.Flags = cdlOFNHideReadOnly + cdlOFNOverwritePrompt + cdlOFNPathMustExist

    ' Set dialog box so an error occurs if the dialogbox is cancelled
    CommonDialog1.CancelError = True

    ' Enables error handling to catch cancel error
    On Error Resume Next
    ' display the dialog box
    CommonDialog1.ShowSave
    If Err Then
        ' This code runs if the dialog was cancelled
        Exit Sub
    End If
    
    'Call your export function
    outFile = CommonDialog1.FileName
    tmpFile = outFile & ".tmp"
    
    TDBGrid1.ExportToDelimitedFile tmpFile, , ";"
 
    Close tmpFile
    
    'Now use this to add the Headers (to a new file)
    Open tmpFile For Input As #1
    Open outFile For Output As #2
 
    'Write Headers
    For i = 0 To TDBGrid1.Columns.Count - 1
         If i = 0 And TDBGrid1.Columns(i).Visible Then
             sLine = Chr(34) & TDBGrid1.Columns(i).Caption & Chr(34)
         ElseIf TDBGrid1.Columns(i).Visible Then
             sLine = sLine & ";" & Chr(34) & TDBGrid1.Columns(i).Caption & Chr(34)
         End If
    Next
 
    Print #2, sLine

     'Write contents of original file
    Do Until EOF(1)
         Line Input #1, sLine
         Print #2, sLine
    Loop
 
    Close #1
    Close #2
 
    Kill tmpFile
          
    Screen.MousePointer = vbDefault
    
    MsgBox "Total exportado: " & Adodc1.Recordset.RecordCount & " registros", vbInformation, "Informação"
End Sub

Private Sub CmdPadrao_Click()
On Error GoTo TrataErro

    Dim VAR_L_VAR_CodMoeda As Variant
    VAR_L_VAR_CodMoeda = Adodc1.Recordset("moed_cd_codmoeda").Value
    Conexao.BeginTrans
    'altera todos os indexadores para 'N'
    Sql = "UPDATE Moedas " & _
          "SET moed_tx_padrao = 'N'"
    Conexao.Execute Sql
    Sql = "UPDATE Moedas " & _
          "SET   moed_tx_padrao = 'S' " & _
          "WHERE moed_cd_codmoeda = " & VAR_L_VAR_CodMoeda
    Conexao.Execute Sql
    Conexao.CommitTrans
    Adodc1.Refresh
    TDBGrid1.ReBind
    
    Exit Sub
TrataErro:
    Conexao.RollbackTrans
    funTrataErros (ComMensagem)
End Sub
Private Sub CmdRemoverFiltro_Click()
    Dim Fil As Integer
    subDesabilitaBotoes
    CmdRemoverFiltro.Enabled = False
    
    Set Formulario = TelaMoedas
    subTelaValoresGlobais "P"
    TelaFiltro.Show 1 'Mostra a tela de filtros no modo modal
    subTelaValoresGlobais "G"

    If Filtrou = True Then
        Fil = ContFil
        subrecarregadadosNV Adodc1, "", FRM_Ordem, "", FRM_FiltroAtual
        ContFil = Fil
    End If
    
    subHabilitaBotoes
    TDBGrid1.SetFocus
End Sub
Private Sub CmdImprimir_Click()
    subDesabilitaBotoes
    Call subImprimeListagemGRID(1, TDBGrid1.PrintInfo, "Listagem de Indexadores")
    subHabilitaBotoes
End Sub
Private Sub CmdAlterar_Click()
    If Not FunVerifica_Permissao("TelaMoedas", "3") Then
        Exit Sub
    End If
        
    TelaMoedas.MousePointer = vbHourglass
    PanMoeda.Tag = Adodc1.Recordset("moed_cd_codmoeda")
    PanBotoes.Enabled = False
    PanMoeda.Visible = True
    PanMoeda.Enabled = True
    PanMoeda.Left = (TDBGrid1.Width - PanMoeda.Width) / 2
    PanMoeda.Top = (TDBGrid1.Height - PanMoeda.Height) / 2
    PanMoeda.Caption = " Alteração de Indexadores"
    subDesabilitaBotoes
    TxtDescricao.Text = FunNulo(Adodc1.Recordset.Fields("moed_tx_Descricao"))
    Txtsimbolo.Text = FunNulo(Adodc1.Recordset.Fields("moed_tx_simbolo"))
    'Txtindexacao.Text = funNulo(Adodc1.Recordset.Fields("moed_tx_indexacao"))
    'If FunNulo(Adodc1.Recordset.Fields("moed_tx_indexacao")) = Null Then
    '    cmbTipoIndexador.ListIndex = -1
    'Else
        If Adodc1.Recordset.Fields("moed_tx_indexacao") = "D" Then
            'cmbTipoIndexador.ListIndex = 0
            OptDiario.Value = True
        Else
            'cmbTipoIndexador.ListIndex = 1
            OptMensal.Value = True
        End If
    'End If
    TxtDescricao.SetFocus
    TelaMoedas.MousePointer = vbDefault
    
End Sub
Private Sub CmdInserir_Click()
    If Not FunVerifica_Permissao("TelaMoedas", "2") Then
        Exit Sub
    End If
    
    TelaMoedas.MousePointer = vbHourglass
    PanMoeda.Tag = "I"
    PanBotoes.Enabled = False
    PanMoeda.Visible = True
    PanMoeda.Enabled = True
    PanMoeda.Left = (TDBGrid1.Width - PanMoeda.Width) / 2
    PanMoeda.Top = (TDBGrid1.Height - PanMoeda.Height) / 2
    PanMoeda.Caption = " Inserção de Indexadores"
    'DataGrid1.Height = 2775
    
    TxtDescricao.Text = ""
    Txtsimbolo.Text = ""
    OptDiario.Value = True
    'cmbTipoIndexador.ListIndex = -1
    
    subDesabilitaBotoes
    TxtDescricao.SetFocus
    TelaMoedas.MousePointer = vbDefault
  
End Sub
Private Sub CmdOrdem_Click()
    subDesabilitaBotoes
    Set Formulario = TelaMoedas
    subTelaValoresGlobais "P"
    TelaOrdem.Show 1
    subTelaValoresGlobais "G"
    subHabilitaBotoes
    TDBGrid1.SetFocus
    
    lblLocalizar.Caption = "Localizar " & xLocalDescricao & ":"
    If InStr(1, xLocalCampo, ",") > 0 Then
        cboLocalizar.ListField = Left(xLocalCampo, InStr(1, xLocalCampo, ",") - 1)
    Else
        cboLocalizar.ListField = xLocalCampo
    End If
    cboLocalizar.Text = ""

End Sub

Private Sub CmdExcluir_Click()
    If Not FunVerifica_Permissao("TelaMoedas", "4") Then
        Exit Sub
    End If
    
    Dim Posicao As Variant
    
    On Error GoTo RotuloErro
    
    Posicao = Adodc1.Recordset.Bookmark
    subDesabilitaBotoes
       
    If MsgBox("Confirma Remoção do Registro ?", vbYesNo + vbCritical + vbDefaultButton2, "CUIDADO") = vbYes Then 'Confirma exclusão

      Conexao.BeginTrans
      Conexao.Execute "DELETE FROM moedas WHERE moed_cd_codmoeda =" & Adodc1.Recordset.Fields("moed_cd_codmoeda")
      Conexao.CommitTrans
      FRM_Quantid = funOperacaoQuantid("-")
      subrecarregadadosNV Adodc1, "", FRM_Ordem, "", FRM_FiltroAtual
    
    End If
       
    subHabilitaBotoes
              
    TDBGrid1.Enabled = True
    
    TDBGrid1.SetFocus
    TDBGrid1.Col = 0
    
    'Posicionando o cursor na linha anterior - Lucas Santiago 22.01.09
    If Posicao > 1 Then
        Adodc1.Recordset.Bookmark = Posicao - 1
    End If
    
    Exit Sub
    
RotuloErro:
    Conexao.RollbackTrans
  '  If Err.Number = -2147467259 Then
      MsgBox "Esse indexador não pode ser excluído pois já possui algum registro relacionado!", vbCritical, "ATENÇÃO"
   ' End If
    subHabilitaBotoes

End Sub
Private Sub CmdSair_Click()
    MdiPrincipal.BarraStatus.Panels(3).Text = ""
    MdiPrincipal.BarraStatus.Panels(1).Text = ""
    MdiPrincipal.MousePointer = vbDefault
    Unload Me
End Sub
Private Sub TDBGrid1_DblClick()
    If Adodc1.Recordset.RecordCount <> 0 Then
        CmdAlterar_Click
    End If
End Sub
Private Sub TDBGrid1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    cboLocalizar.Text = ""
    If Button = 2 Then  'Verifica se o botão da direita foi pressionado
        filtra_especial = False
        Set Formulario = TelaMoedas
        subTelaValoresGlobais "P"
        FunExecutaFiltro Adodc1, Formulario, x
        subTelaValoresGlobais "G"
    End If
    
    subHabilitaBotoes
          
    If TDBGrid1.ApproxCount = 0 Then
        CmdExcluir.Enabled = False 'Desabilita o botão Excluir
        CmdImprimir.Enabled = False
        CmdOrdem.Enabled = False
        CmdRemoverFiltro.Enabled = False
        CmdCotacoes.Enabled = False
    End If
    
End Sub
Private Sub Form_Activate()
    subrecarregadadosNV Adodc1, "", FRM_Ordem, "", FRM_FiltroAtual
    Set Formulario = TelaMoedas
    If Not PrimeiraVez Then
        subTelaValoresGlobais "P"
    End If
    PrimeiraVez = False
    
    MdiPrincipal.BarraStatus.Panels(3).Text = CStr(Adodc1.Recordset.RecordCount) + "/" + CStr(Adodc1.Recordset.RecordCount)
End Sub
'Habilita os Botões da Tela
Private Sub subHabilitaBotoes()
    CmdInserir.Enabled = True
    If Not FunTabelaVazia(Adodc1) Then
        CmdAlterar.Enabled = True
        CmdExcluir.Enabled = True
        CmdRemoverFiltro.Enabled = True
        CmdCotacoes.Enabled = True
        CmdOrdem.Enabled = True
        CmdImprimir.Enabled = True
'        CmdPadrao.Enabled = True
        cboLocalizar.Enabled = True
    Else
        CmdAlterar.Enabled = False
        CmdExcluir.Enabled = False
        CmdRemoverFiltro.Enabled = False
        CmdCotacoes.Enabled = False
        CmdOrdem.Enabled = False
        CmdImprimir.Enabled = False
'        CmdPadrao.Enabled = False
        cboLocalizar.Enabled = False
    End If
    CmdSair.Enabled = True

    'Verifica se tem Filtro Ativo
    If ContFil = 0 Then
       CmdRemoverFiltro.Enabled = False
    Else
       CmdRemoverFiltro.Enabled = True
    End If
End Sub
Private Sub Form_Load()

    Set Formulario = TelaMoedas
    
    Call subAjustaTelaLoc
    
    subManutencaoJanelasAtivas "I", "TelaMoedas"
    
    xLocalDescricao = "Descrição"
    
    subCarregaDadosNV Adodc1, "Moedas", "moed_tx_descricao", "", "SELECT * FROM Moedas"
    
    filtra_especial = False
    
    subHabilitaBotoes
    
    subTelaValoresGlobais "G"
    PrimeiraVez = True
    
    'Checar a existencia de qualquer entrada para inabilitar a possibilidade de
    'mudar a moeda padrão
    Dim XLO_CHECK As Object
    XGT_SQL = "SELECT * FROM EntradaMateriais"
    SubQOpenRecordset XLO_CHECK, XGT_SQL, Estatico
    If Not XLO_CHECK.EOF Then
        CmdPadrao.Enabled = False
    End If
    XLO_CHECK.Close
    
    If TDBGrid1.ApproxCount = 0 Then
        CmdAlterar.Enabled = False
        CmdExcluir.Enabled = False
        CmdImprimir.Enabled = False
        CmdOrdem.Enabled = False
        CmdCotacoes.Enabled = False
    End If
    
    If XGT_NomeAplic = "QUALIFIN" Or XGT_NomeAplic = "QUALICAPI" Then
      TDBGrid1.Columns(4).Visible = False
    End If
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
    subManutencaoJanelasAtivas "R", "TelaMoedas"
End Sub


