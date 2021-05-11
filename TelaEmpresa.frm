VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form TelaEmpresas 
   Caption         =   "Tabela de Empresas"
   ClientHeight    =   4860
   ClientLeft      =   1515
   ClientTop       =   1815
   ClientWidth     =   9360
   Icon            =   "TelaEmpresa.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4860
   ScaleWidth      =   9360
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8880
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
      Bindings        =   "TelaEmpresa.frx":0442
      Height          =   4080
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   7197
      _LayoutType     =   4
      _RowHeight      =   25
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Código"
      Columns(0).DataField=   "empr_cd_Empresa"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Fantasia"
      Columns(1).DataField=   "empr_tx_fantasia"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Razão Social"
      Columns(2).DataField=   "empr_tx_razaosocial"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "CGC"
      Columns(3).DataField=   "empr_tx_cgc"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Telefone"
      Columns(4).DataField=   "empr_tx_telefone"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Fax"
      Columns(5).DataField=   "empr_tx_fax"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "Endereço"
      Columns(6).DataField=   "empr_tx_endereco"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "Nº Porta"
      Columns(7).DataField=   "empr_nr_porta"
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "Bairro"
      Columns(8).DataField=   "empr_tx_bairro"
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "Cidade"
      Columns(9).DataField=   "empr_tx_cidade"
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "Estado"
      Columns(10).DataField=   "empr_tx_estado"
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "País"
      Columns(11).DataField=   "empr_tx_pais"
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).Caption=   "CEP"
      Columns(12).DataField=   "empr_tx_cep"
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(13)._VlistStyle=   0
      Columns(13)._MaxComboItems=   5
      Columns(13).Caption=   "EMail"
      Columns(13).DataField=   "empr_tx_email"
      Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(14)._VlistStyle=   0
      Columns(14)._MaxComboItems=   5
      Columns(14).Caption=   "Home Page"
      Columns(14).DataField=   "empr_tx_homepage"
      Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(15)._VlistStyle=   0
      Columns(15)._MaxComboItems=   5
      Columns(15).Caption=   "Inscrição Estadual"
      Columns(15).DataField=   "empr_tx_InscrEst"
      Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(16)._VlistStyle=   0
      Columns(16)._MaxComboItems=   5
      Columns(16).Caption=   "Inscrição Municipal"
      Columns(16).DataField=   "empr_tx_InscrMun"
      Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(17)._VlistStyle=   0
      Columns(17)._MaxComboItems=   5
      Columns(17).Caption=   "Máscara Centro de Custo"
      Columns(17).DataField=   "empr_tx_pcrmascara"
      Columns(17)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(18)._VlistStyle=   0
      Columns(18)._MaxComboItems=   5
      Columns(18).Caption=   "Nível Centro de Custo"
      Columns(18).DataField=   "empr_nr_pcrnivel"
      Columns(18)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(19)._VlistStyle=   0
      Columns(19)._MaxComboItems=   5
      Columns(19).Caption=   "Máscara Plano de Contas"
      Columns(19).DataField=   "empr_tx_planocmascara"
      Columns(19)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(20)._VlistStyle=   0
      Columns(20)._MaxComboItems=   5
      Columns(20).Caption=   "Nível Plano de Contas"
      Columns(20).DataField=   "empr_nr_planocnivel"
      Columns(20)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(21)._VlistStyle=   0
      Columns(21)._MaxComboItems=   5
      Columns(21).Caption=   "Conta Corrente"
      Columns(21).DataField=   "baempresa"
      Columns(21)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(22)._VlistStyle=   0
      Columns(22)._MaxComboItems=   5
      Columns(22).Caption=   "Exp. Contábil"
      Columns(22).DataField=   "ExpContabil"
      Columns(22)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(23)._VlistStyle=   0
      Columns(23)._MaxComboItems=   5
      Columns(23).Caption=   "Integrado CAPI"
      Columns(23).DataField=   "CAPI"
      Columns(23)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(24)._VlistStyle=   0
      Columns(24)._MaxComboItems=   5
      Columns(24).Caption=   "Empresa Ativa"
      Columns(24).DataField=   "empr_tx_atual"
      Columns(24)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(25)._VlistStyle=   0
      Columns(25)._MaxComboItems=   5
      Columns(25).Caption=   "Dias para Filtro"
      Columns(25).DataField=   "empr_nr_QtdDiasFiltro"
      Columns(25)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(26)._VlistStyle=   0
      Columns(26)._MaxComboItems=   5
      Columns(26).Caption=   "Codigo"
      Columns(26).DataField=   "empr_cd_empresa"
      Columns(26)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(27)._VlistStyle=   0
      Columns(27)._MaxComboItems=   5
      Columns(27).Caption=   "Simples Nacional"
      Columns(27).DataField=   "SimplesNacional"
      Columns(27)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   28
      Splits(0)._UserFlags=   0
      Splits(0).ExtendRightColumn=   -1  'True
      Splits(0).MarqueeStyle=   3
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).AllowColMove=   -1  'True
      Splits(0).FetchRowStyle=   -1  'True
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=28"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=20"
      Splits(0)._ColumnProps(6)=   "Column(0).WrapText=1"
      Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(8)=   "Column(1).Width=3175"
      Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=3096"
      Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=8720"
      Splits(0)._ColumnProps(13)=   "Column(1).WrapText=1"
      Splits(0)._ColumnProps(14)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(15)=   "Column(1)._MinWidth=64"
      Splits(0)._ColumnProps(16)=   "Column(2).Width=4789"
      Splits(0)._ColumnProps(17)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(18)=   "Column(2)._WidthInPix=4710"
      Splits(0)._ColumnProps(19)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(20)=   "Column(2)._ColStyle=8720"
      Splits(0)._ColumnProps(21)=   "Column(2).WrapText=1"
      Splits(0)._ColumnProps(22)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(23)=   "Column(2)._MinWidth=-1"
      Splits(0)._ColumnProps(24)=   "Column(3).Width=2910"
      Splits(0)._ColumnProps(25)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(26)=   "Column(3)._WidthInPix=2831"
      Splits(0)._ColumnProps(27)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(28)=   "Column(3)._ColStyle=8720"
      Splits(0)._ColumnProps(29)=   "Column(3).WrapText=1"
      Splits(0)._ColumnProps(30)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(31)=   "Column(4).Width=2355"
      Splits(0)._ColumnProps(32)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(33)=   "Column(4)._WidthInPix=2275"
      Splits(0)._ColumnProps(34)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(35)=   "Column(4)._ColStyle=8720"
      Splits(0)._ColumnProps(36)=   "Column(4).WrapText=1"
      Splits(0)._ColumnProps(37)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(38)=   "Column(5).Width=2381"
      Splits(0)._ColumnProps(39)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(40)=   "Column(5)._WidthInPix=2302"
      Splits(0)._ColumnProps(41)=   "Column(5)._EditAlways=0"
      Splits(0)._ColumnProps(42)=   "Column(5)._ColStyle=8720"
      Splits(0)._ColumnProps(43)=   "Column(5).WrapText=1"
      Splits(0)._ColumnProps(44)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(45)=   "Column(5)._MinWidth=56550368"
      Splits(0)._ColumnProps(46)=   "Column(6).Width=4551"
      Splits(0)._ColumnProps(47)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(48)=   "Column(6)._WidthInPix=4471"
      Splits(0)._ColumnProps(49)=   "Column(6)._EditAlways=0"
      Splits(0)._ColumnProps(50)=   "Column(6)._ColStyle=8720"
      Splits(0)._ColumnProps(51)=   "Column(6).WrapText=1"
      Splits(0)._ColumnProps(52)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(53)=   "Column(6)._MinWidth=56549904"
      Splits(0)._ColumnProps(54)=   "Column(7).Width=2725"
      Splits(0)._ColumnProps(55)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(56)=   "Column(7)._WidthInPix=2646"
      Splits(0)._ColumnProps(57)=   "Column(7)._EditAlways=0"
      Splits(0)._ColumnProps(58)=   "Column(7)._ColStyle=8212"
      Splits(0)._ColumnProps(59)=   "Column(7).WrapText=1"
      Splits(0)._ColumnProps(60)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(61)=   "Column(8).Width=2593"
      Splits(0)._ColumnProps(62)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(63)=   "Column(8)._WidthInPix=2514"
      Splits(0)._ColumnProps(64)=   "Column(8)._EditAlways=0"
      Splits(0)._ColumnProps(65)=   "Column(8)._ColStyle=8720"
      Splits(0)._ColumnProps(66)=   "Column(8).WrapText=1"
      Splits(0)._ColumnProps(67)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(68)=   "Column(9).Width=2434"
      Splits(0)._ColumnProps(69)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(70)=   "Column(9)._WidthInPix=2355"
      Splits(0)._ColumnProps(71)=   "Column(9)._EditAlways=0"
      Splits(0)._ColumnProps(72)=   "Column(9)._ColStyle=8720"
      Splits(0)._ColumnProps(73)=   "Column(9).WrapText=1"
      Splits(0)._ColumnProps(74)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(75)=   "Column(10).Width=1217"
      Splits(0)._ColumnProps(76)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(77)=   "Column(10)._WidthInPix=1138"
      Splits(0)._ColumnProps(78)=   "Column(10)._EditAlways=0"
      Splits(0)._ColumnProps(79)=   "Column(10)._ColStyle=8721"
      Splits(0)._ColumnProps(80)=   "Column(10).WrapText=1"
      Splits(0)._ColumnProps(81)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(82)=   "Column(11).Width=2725"
      Splits(0)._ColumnProps(83)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(84)=   "Column(11)._WidthInPix=2646"
      Splits(0)._ColumnProps(85)=   "Column(11)._EditAlways=0"
      Splits(0)._ColumnProps(86)=   "Column(11)._ColStyle=8212"
      Splits(0)._ColumnProps(87)=   "Column(11).WrapText=1"
      Splits(0)._ColumnProps(88)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(89)=   "Column(12).Width=1720"
      Splits(0)._ColumnProps(90)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(91)=   "Column(12)._WidthInPix=1640"
      Splits(0)._ColumnProps(92)=   "Column(12)._EditAlways=0"
      Splits(0)._ColumnProps(93)=   "Column(12)._ColStyle=8720"
      Splits(0)._ColumnProps(94)=   "Column(12).WrapText=1"
      Splits(0)._ColumnProps(95)=   "Column(12).Order=13"
      Splits(0)._ColumnProps(96)=   "Column(13).Width=4419"
      Splits(0)._ColumnProps(97)=   "Column(13).DividerColor=0"
      Splits(0)._ColumnProps(98)=   "Column(13)._WidthInPix=4339"
      Splits(0)._ColumnProps(99)=   "Column(13)._EditAlways=0"
      Splits(0)._ColumnProps(100)=   "Column(13)._ColStyle=8720"
      Splits(0)._ColumnProps(101)=   "Column(13).WrapText=1"
      Splits(0)._ColumnProps(102)=   "Column(13).Order=14"
      Splits(0)._ColumnProps(103)=   "Column(14).Width=2725"
      Splits(0)._ColumnProps(104)=   "Column(14).DividerColor=0"
      Splits(0)._ColumnProps(105)=   "Column(14)._WidthInPix=2646"
      Splits(0)._ColumnProps(106)=   "Column(14)._EditAlways=0"
      Splits(0)._ColumnProps(107)=   "Column(14)._ColStyle=8212"
      Splits(0)._ColumnProps(108)=   "Column(14).WrapText=1"
      Splits(0)._ColumnProps(109)=   "Column(14).Order=15"
      Splits(0)._ColumnProps(110)=   "Column(15).Width=2355"
      Splits(0)._ColumnProps(111)=   "Column(15).DividerColor=0"
      Splits(0)._ColumnProps(112)=   "Column(15)._WidthInPix=2275"
      Splits(0)._ColumnProps(113)=   "Column(15)._EditAlways=0"
      Splits(0)._ColumnProps(114)=   "Column(15)._ColStyle=8720"
      Splits(0)._ColumnProps(115)=   "Column(15).WrapText=1"
      Splits(0)._ColumnProps(116)=   "Column(15).Order=16"
      Splits(0)._ColumnProps(117)=   "Column(16).Width=2725"
      Splits(0)._ColumnProps(118)=   "Column(16).DividerColor=0"
      Splits(0)._ColumnProps(119)=   "Column(16)._WidthInPix=2646"
      Splits(0)._ColumnProps(120)=   "Column(16)._EditAlways=0"
      Splits(0)._ColumnProps(121)=   "Column(16)._ColStyle=8212"
      Splits(0)._ColumnProps(122)=   "Column(16).WrapText=1"
      Splits(0)._ColumnProps(123)=   "Column(16).Order=17"
      Splits(0)._ColumnProps(124)=   "Column(17).Width=2540"
      Splits(0)._ColumnProps(125)=   "Column(17).DividerColor=0"
      Splits(0)._ColumnProps(126)=   "Column(17)._WidthInPix=2461"
      Splits(0)._ColumnProps(127)=   "Column(17)._EditAlways=0"
      Splits(0)._ColumnProps(128)=   "Column(17)._ColStyle=8720"
      Splits(0)._ColumnProps(129)=   "Column(17).WrapText=1"
      Splits(0)._ColumnProps(130)=   "Column(17).Order=18"
      Splits(0)._ColumnProps(131)=   "Column(18).Width=2275"
      Splits(0)._ColumnProps(132)=   "Column(18).DividerColor=0"
      Splits(0)._ColumnProps(133)=   "Column(18)._WidthInPix=2196"
      Splits(0)._ColumnProps(134)=   "Column(18)._EditAlways=0"
      Splits(0)._ColumnProps(135)=   "Column(18)._ColStyle=8720"
      Splits(0)._ColumnProps(136)=   "Column(18).WrapText=1"
      Splits(0)._ColumnProps(137)=   "Column(18).Order=19"
      Splits(0)._ColumnProps(138)=   "Column(19).Width=2090"
      Splits(0)._ColumnProps(139)=   "Column(19).DividerColor=0"
      Splits(0)._ColumnProps(140)=   "Column(19)._WidthInPix=2011"
      Splits(0)._ColumnProps(141)=   "Column(19)._EditAlways=0"
      Splits(0)._ColumnProps(142)=   "Column(19)._ColStyle=8720"
      Splits(0)._ColumnProps(143)=   "Column(19).WrapText=1"
      Splits(0)._ColumnProps(144)=   "Column(19).Order=20"
      Splits(0)._ColumnProps(145)=   "Column(20).Width=2090"
      Splits(0)._ColumnProps(146)=   "Column(20).DividerColor=0"
      Splits(0)._ColumnProps(147)=   "Column(20)._WidthInPix=2011"
      Splits(0)._ColumnProps(148)=   "Column(20)._EditAlways=0"
      Splits(0)._ColumnProps(149)=   "Column(20)._ColStyle=8720"
      Splits(0)._ColumnProps(150)=   "Column(20).WrapText=1"
      Splits(0)._ColumnProps(151)=   "Column(20).Order=21"
      Splits(0)._ColumnProps(152)=   "Column(21).Width=2725"
      Splits(0)._ColumnProps(153)=   "Column(21).DividerColor=0"
      Splits(0)._ColumnProps(154)=   "Column(21)._WidthInPix=2646"
      Splits(0)._ColumnProps(155)=   "Column(21)._EditAlways=0"
      Splits(0)._ColumnProps(156)=   "Column(21)._ColStyle=8720"
      Splits(0)._ColumnProps(157)=   "Column(21).Visible=0"
      Splits(0)._ColumnProps(158)=   "Column(21).WrapText=1"
      Splits(0)._ColumnProps(159)=   "Column(21).Order=22"
      Splits(0)._ColumnProps(160)=   "Column(22).Width=2725"
      Splits(0)._ColumnProps(161)=   "Column(22).DividerColor=0"
      Splits(0)._ColumnProps(162)=   "Column(22)._WidthInPix=2646"
      Splits(0)._ColumnProps(163)=   "Column(22)._EditAlways=0"
      Splits(0)._ColumnProps(164)=   "Column(22)._ColStyle=8212"
      Splits(0)._ColumnProps(165)=   "Column(22).WrapText=1"
      Splits(0)._ColumnProps(166)=   "Column(22).Order=23"
      Splits(0)._ColumnProps(167)=   "Column(23).Width=2725"
      Splits(0)._ColumnProps(168)=   "Column(23).DividerColor=0"
      Splits(0)._ColumnProps(169)=   "Column(23)._WidthInPix=2646"
      Splits(0)._ColumnProps(170)=   "Column(23)._EditAlways=0"
      Splits(0)._ColumnProps(171)=   "Column(23)._ColStyle=8212"
      Splits(0)._ColumnProps(172)=   "Column(23).WrapText=1"
      Splits(0)._ColumnProps(173)=   "Column(23).Order=24"
      Splits(0)._ColumnProps(174)=   "Column(24).Width=2725"
      Splits(0)._ColumnProps(175)=   "Column(24).DividerColor=0"
      Splits(0)._ColumnProps(176)=   "Column(24)._WidthInPix=2646"
      Splits(0)._ColumnProps(177)=   "Column(24)._EditAlways=0"
      Splits(0)._ColumnProps(178)=   "Column(24)._ColStyle=8212"
      Splits(0)._ColumnProps(179)=   "Column(24).Visible=0"
      Splits(0)._ColumnProps(180)=   "Column(24).WrapText=1"
      Splits(0)._ColumnProps(181)=   "Column(24).Order=25"
      Splits(0)._ColumnProps(182)=   "Column(25).Width=2725"
      Splits(0)._ColumnProps(183)=   "Column(25).DividerColor=0"
      Splits(0)._ColumnProps(184)=   "Column(25)._WidthInPix=2646"
      Splits(0)._ColumnProps(185)=   "Column(25)._EditAlways=0"
      Splits(0)._ColumnProps(186)=   "Column(25)._ColStyle=20"
      Splits(0)._ColumnProps(187)=   "Column(25).WrapText=1"
      Splits(0)._ColumnProps(188)=   "Column(25).Order=26"
      Splits(0)._ColumnProps(189)=   "Column(26).Width=2725"
      Splits(0)._ColumnProps(190)=   "Column(26).DividerColor=0"
      Splits(0)._ColumnProps(191)=   "Column(26)._WidthInPix=2646"
      Splits(0)._ColumnProps(192)=   "Column(26)._EditAlways=0"
      Splits(0)._ColumnProps(193)=   "Column(26)._ColStyle=8212"
      Splits(0)._ColumnProps(194)=   "Column(26).Visible=0"
      Splits(0)._ColumnProps(195)=   "Column(26).WrapText=1"
      Splits(0)._ColumnProps(196)=   "Column(26).Order=27"
      Splits(0)._ColumnProps(197)=   "Column(27).Width=1773"
      Splits(0)._ColumnProps(198)=   "Column(27).DividerColor=0"
      Splits(0)._ColumnProps(199)=   "Column(27)._WidthInPix=1693"
      Splits(0)._ColumnProps(200)=   "Column(27)._EditAlways=0"
      Splits(0)._ColumnProps(201)=   "Column(27)._ColStyle=16"
      Splits(0)._ColumnProps(202)=   "Column(27).WrapText=1"
      Splits(0)._ColumnProps(203)=   "Column(27).Order=28"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
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
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.fgcolor=&H80000008&,.bold=0,.fontsize=825"
      _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
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
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=138,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=135,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=136,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=137,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=28,.parent=13,.alignment=0,.locked=-1"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14,.alignment=2"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15,.alignment=3"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=32,.parent=13,.alignment=0,.locked=-1"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14,.alignment=2"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=46,.parent=13,.alignment=0,.locked=-1"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14,.alignment=2"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=50,.parent=13,.alignment=0,.locked=-1"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=14,.alignment=2"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=15"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=17"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=54,.parent=13,.alignment=0,.locked=-1"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=51,.parent=14,.alignment=2"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=52,.parent=15"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=53,.parent=17"
      _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=58,.parent=13,.alignment=0,.locked=-1"
      _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=55,.parent=14,.alignment=2"
      _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=56,.parent=15"
      _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=57,.parent=17"
      _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=134,.parent=13,.locked=-1"
      _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=131,.parent=14"
      _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=132,.parent=15"
      _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=133,.parent=17"
      _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=62,.parent=13,.alignment=0,.locked=-1"
      _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=59,.parent=14,.alignment=2"
      _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=60,.parent=15"
      _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=61,.parent=17"
      _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=66,.parent=13,.alignment=0,.locked=-1"
      _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=63,.parent=14,.alignment=2"
      _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=64,.parent=15"
      _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=65,.parent=17"
      _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=70,.parent=13,.alignment=2,.locked=-1"
      _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=67,.parent=14,.alignment=2"
      _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=68,.parent=15"
      _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=69,.parent=17"
      _StyleDefs(80)  =   "Splits(0).Columns(11).Style:id=114,.parent=13,.locked=-1"
      _StyleDefs(81)  =   "Splits(0).Columns(11).HeadingStyle:id=111,.parent=14"
      _StyleDefs(82)  =   "Splits(0).Columns(11).FooterStyle:id=112,.parent=15"
      _StyleDefs(83)  =   "Splits(0).Columns(11).EditorStyle:id=113,.parent=17"
      _StyleDefs(84)  =   "Splits(0).Columns(12).Style:id=74,.parent=13,.alignment=0,.locked=-1"
      _StyleDefs(85)  =   "Splits(0).Columns(12).HeadingStyle:id=71,.parent=14,.alignment=2"
      _StyleDefs(86)  =   "Splits(0).Columns(12).FooterStyle:id=72,.parent=15"
      _StyleDefs(87)  =   "Splits(0).Columns(12).EditorStyle:id=73,.parent=17"
      _StyleDefs(88)  =   "Splits(0).Columns(13).Style:id=78,.parent=13,.alignment=0,.locked=-1"
      _StyleDefs(89)  =   "Splits(0).Columns(13).HeadingStyle:id=75,.parent=14,.alignment=2"
      _StyleDefs(90)  =   "Splits(0).Columns(13).FooterStyle:id=76,.parent=15"
      _StyleDefs(91)  =   "Splits(0).Columns(13).EditorStyle:id=77,.parent=17"
      _StyleDefs(92)  =   "Splits(0).Columns(14).Style:id=122,.parent=13,.locked=-1"
      _StyleDefs(93)  =   "Splits(0).Columns(14).HeadingStyle:id=119,.parent=14"
      _StyleDefs(94)  =   "Splits(0).Columns(14).FooterStyle:id=120,.parent=15"
      _StyleDefs(95)  =   "Splits(0).Columns(14).EditorStyle:id=121,.parent=17"
      _StyleDefs(96)  =   "Splits(0).Columns(15).Style:id=82,.parent=13,.alignment=0,.locked=-1"
      _StyleDefs(97)  =   "Splits(0).Columns(15).HeadingStyle:id=79,.parent=14,.alignment=2"
      _StyleDefs(98)  =   "Splits(0).Columns(15).FooterStyle:id=80,.parent=15"
      _StyleDefs(99)  =   "Splits(0).Columns(15).EditorStyle:id=81,.parent=17"
      _StyleDefs(100) =   "Splits(0).Columns(16).Style:id=126,.parent=13,.locked=-1"
      _StyleDefs(101) =   "Splits(0).Columns(16).HeadingStyle:id=123,.parent=14"
      _StyleDefs(102) =   "Splits(0).Columns(16).FooterStyle:id=124,.parent=15"
      _StyleDefs(103) =   "Splits(0).Columns(16).EditorStyle:id=125,.parent=17"
      _StyleDefs(104) =   "Splits(0).Columns(17).Style:id=94,.parent=13,.alignment=0,.locked=-1"
      _StyleDefs(105) =   "Splits(0).Columns(17).HeadingStyle:id=91,.parent=14,.alignment=2"
      _StyleDefs(106) =   "Splits(0).Columns(17).FooterStyle:id=92,.parent=15"
      _StyleDefs(107) =   "Splits(0).Columns(17).EditorStyle:id=93,.parent=17"
      _StyleDefs(108) =   "Splits(0).Columns(18).Style:id=98,.parent=13,.alignment=0,.locked=-1"
      _StyleDefs(109) =   "Splits(0).Columns(18).HeadingStyle:id=95,.parent=14,.alignment=2"
      _StyleDefs(110) =   "Splits(0).Columns(18).FooterStyle:id=96,.parent=15"
      _StyleDefs(111) =   "Splits(0).Columns(18).EditorStyle:id=97,.parent=17"
      _StyleDefs(112) =   "Splits(0).Columns(19).Style:id=106,.parent=13,.alignment=0,.locked=-1"
      _StyleDefs(113) =   "Splits(0).Columns(19).HeadingStyle:id=103,.parent=14,.alignment=2"
      _StyleDefs(114) =   "Splits(0).Columns(19).FooterStyle:id=104,.parent=15"
      _StyleDefs(115) =   "Splits(0).Columns(19).EditorStyle:id=105,.parent=17"
      _StyleDefs(116) =   "Splits(0).Columns(20).Style:id=110,.parent=13,.alignment=0,.locked=-1"
      _StyleDefs(117) =   "Splits(0).Columns(20).HeadingStyle:id=107,.parent=14,.alignment=2"
      _StyleDefs(118) =   "Splits(0).Columns(20).FooterStyle:id=108,.parent=15"
      _StyleDefs(119) =   "Splits(0).Columns(20).EditorStyle:id=109,.parent=17"
      _StyleDefs(120) =   "Splits(0).Columns(21).Style:id=118,.parent=13,.alignment=0,.locked=-1"
      _StyleDefs(121) =   "Splits(0).Columns(21).HeadingStyle:id=115,.parent=14,.alignment=2"
      _StyleDefs(122) =   "Splits(0).Columns(21).FooterStyle:id=116,.parent=15"
      _StyleDefs(123) =   "Splits(0).Columns(21).EditorStyle:id=117,.parent=17"
      _StyleDefs(124) =   "Splits(0).Columns(22).Style:id=130,.parent=13,.locked=-1"
      _StyleDefs(125) =   "Splits(0).Columns(22).HeadingStyle:id=127,.parent=14"
      _StyleDefs(126) =   "Splits(0).Columns(22).FooterStyle:id=128,.parent=15"
      _StyleDefs(127) =   "Splits(0).Columns(22).EditorStyle:id=129,.parent=17"
      _StyleDefs(128) =   "Splits(0).Columns(23).Style:id=86,.parent=13,.locked=-1"
      _StyleDefs(129) =   "Splits(0).Columns(23).HeadingStyle:id=83,.parent=14"
      _StyleDefs(130) =   "Splits(0).Columns(23).FooterStyle:id=84,.parent=15"
      _StyleDefs(131) =   "Splits(0).Columns(23).EditorStyle:id=85,.parent=17"
      _StyleDefs(132) =   "Splits(0).Columns(24).Style:id=90,.parent=13,.locked=-1"
      _StyleDefs(133) =   "Splits(0).Columns(24).HeadingStyle:id=87,.parent=14"
      _StyleDefs(134) =   "Splits(0).Columns(24).FooterStyle:id=88,.parent=15"
      _StyleDefs(135) =   "Splits(0).Columns(24).EditorStyle:id=89,.parent=17"
      _StyleDefs(136) =   "Splits(0).Columns(25).Style:id=142,.parent=13"
      _StyleDefs(137) =   "Splits(0).Columns(25).HeadingStyle:id=139,.parent=14"
      _StyleDefs(138) =   "Splits(0).Columns(25).FooterStyle:id=140,.parent=15"
      _StyleDefs(139) =   "Splits(0).Columns(25).EditorStyle:id=141,.parent=17"
      _StyleDefs(140) =   "Splits(0).Columns(26).Style:id=102,.parent=13,.locked=-1"
      _StyleDefs(141) =   "Splits(0).Columns(26).HeadingStyle:id=99,.parent=14"
      _StyleDefs(142) =   "Splits(0).Columns(26).FooterStyle:id=100,.parent=15"
      _StyleDefs(143) =   "Splits(0).Columns(26).EditorStyle:id=101,.parent=17"
      _StyleDefs(144) =   "Splits(0).Columns(27).Style:id=146,.parent=13,.alignment=0"
      _StyleDefs(145) =   "Splits(0).Columns(27).HeadingStyle:id=143,.parent=14"
      _StyleDefs(146) =   "Splits(0).Columns(27).FooterStyle:id=144,.parent=15"
      _StyleDefs(147) =   "Splits(0).Columns(27).EditorStyle:id=145,.parent=17"
      _StyleDefs(148) =   "Named:id=33:Normal"
      _StyleDefs(149) =   ":id=33,.parent=0"
      _StyleDefs(150) =   "Named:id=34:Heading"
      _StyleDefs(151) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(152) =   ":id=34,.wraptext=-1"
      _StyleDefs(153) =   "Named:id=35:Footing"
      _StyleDefs(154) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(155) =   "Named:id=36:Selected"
      _StyleDefs(156) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(157) =   "Named:id=37:Caption"
      _StyleDefs(158) =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(159) =   "Named:id=38:HighlightRow"
      _StyleDefs(160) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(161) =   "Named:id=39:EvenRow"
      _StyleDefs(162) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(163) =   "Named:id=40:OddRow"
      _StyleDefs(164) =   ":id=40,.parent=33"
      _StyleDefs(165) =   "Named:id=41:RecordSelector"
      _StyleDefs(166) =   ":id=41,.parent=34"
      _StyleDefs(167) =   "Named:id=42:FilterBar"
      _StyleDefs(168) =   ":id=42,.parent=33"
   End
   Begin Threed.SSPanel PanBotoes 
      Height          =   375
      Left            =   60
      TabIndex        =   0
      Top             =   4455
      Width           =   9240
      _Version        =   65536
      _ExtentX        =   16298
      _ExtentY        =   661
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   0
      Begin VB.CommandButton cmdExport 
         Height          =   330
         Left            =   8760
         Picture         =   "TelaEmpresa.frx":0457
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   0
         Width           =   375
      End
      Begin VB.CommandButton CmdAtivar 
         Caption         =   "A&tivar"
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
         Left            =   3321
         TabIndex        =   9
         Top             =   15
         Width           =   1020
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
         Left            =   7635
         TabIndex        =   7
         Top             =   15
         Width           =   1020
      End
      Begin VB.CommandButton Cmdimprimir 
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
         Left            =   6552
         TabIndex        =   6
         Top             =   15
         Width           =   1020
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
         Left            =   5475
         TabIndex        =   5
         Top             =   15
         Width           =   1020
      End
      Begin VB.CommandButton Botao_Fil 
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
         Left            =   4398
         TabIndex        =   4
         Top             =   15
         Width           =   1020
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
         Left            =   2244
         TabIndex        =   3
         Top             =   15
         Width           =   1020
      End
      Begin VB.CommandButton CmdAlterar 
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
         Left            =   1167
         TabIndex        =   2
         Top             =   15
         Width           =   1020
      End
      Begin VB.CommandButton CmdInserir 
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
         Left            =   80
         TabIndex        =   1
         Top             =   15
         Width           =   1020
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   120
      Top             =   960
      Visible         =   0   'False
      Width           =   2565
      _ExtentX        =   4524
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
   Begin MSDataListLib.DataCombo cboLocalizar 
      Bindings        =   "TelaEmpresa.frx":0606
      Height          =   315
      Left            =   2760
      TabIndex        =   10
      Top             =   4125
      Width           =   5160
      _ExtentX        =   9102
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      ListField       =   "empr_tx_fantasia"
      BoundColumn     =   "empr_cd_empresa"
      Text            =   ""
   End
   Begin VB.Label lblLocalizar 
      Alignment       =   1  'Right Justify
      Caption         =   "Localizar Fantasia:"
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   0
      TabIndex        =   11
      Top             =   4140
      Width           =   2670
   End
End
Attribute VB_Name = "TelaEmpresas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*****************************************************************************************
'Variaveis que guardarão valores específicos de cada formulário
Dim PrimeiraVez As Boolean

Dim FRM_ContFil As Integer ' Contador que armazena o número de filtros ativos
Dim FRM_Ordem As String ' Nome do Campo que tem a chave da Tabela ativa
Dim FRM_Campos_Obrig As String ' Lista dos campos que são obrigatórios na tabela
Dim FRM_TipoVar As String ' Tipo do campo a ser filtrado
Dim FRM_FiltroOrig As String ' Indica os filtros determinados pelo próprio sistema
Dim FRM_FiltroAtual As String ' Indica o filtro atual da tabela corrente
Dim FRM_filtra_especial As Integer ' Indica se a tabela utilizada possui filtros pré-definidos pelo sistema em FiltroOrig
Dim FRM_Tem_Filtro As Boolean ' Informa se já foi relizado um filtro na tabela pelo usuário
Dim FRM_ConTelaOrdem As Integer ' Indica o número de chaves utilizados para ordenação
Dim FRM_Quantid As Long  ' Quantidade de registros da tabela
Dim FRM_Filtrou As Boolean ' Informa se Algum campo foi filtrado
Dim FRM_Ordens() As Integer    ' Vetor que mantem a ordem dos campos
Dim FRM_Filtros() As String ' Array para armazenar os Filtros da Tabela

Dim FRM_BarraFerramentasPressionada As Boolean
Dim FRM_VariasVezes As Boolean    'Indica se vai executar a insercao na TelaPadrao varias ou uma vez
Dim FRM_XInserir As String        'Flag para identificar se entrou na rotina para inserir ou alterar ('I'Inserir, 'A'Alterar, 'C'Consultar)
Dim FRM_XOrigem As String         'Identifica a origem (que formulario) chamou o formulário ativo
Dim FRM_SQLAtivo As String        'Atualizar o SQL Ativo após modificações de filtro e ordem

Dim ExpTexto As Collection   'Vetor com indice Alfanumérico para compor a coluna de marcação

Function FunVerificaHabilitacaoEmpresa() As Boolean
    
    Dim XLO_FUNCIONARIOEMPRESA As ADODB.Recordset
    
    SubQOpenRecordset XLO_FUNCIONARIOEMPRESA, "SELECT * FROM FuncionarioEmpresa WHERE empr_cd_empresa = " & TDBGrid1.Columns("empr_cd_empresa") & " and func_cd_operador = " & func_cd_operador, Estatico
    
    If XLO_FUNCIONARIOEMPRESA.EOF Then
        XLO_FUNCIONARIOEMPRESA.Close
        Set XLO_FUNCIONARIOEMPRESA = Nothing
        FunVerificaHabilitacaoEmpresa = False
    Else
        XLO_FUNCIONARIOEMPRESA.Close
        Set XLO_FUNCIONARIOEMPRESA = Nothing
        FunVerificaHabilitacaoEmpresa = True
    End If
    
    
End Function

Sub TelaValoresGlobais(GuardaOUPegaValores As String)
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
        xLocalCampo = Ordem
        xLocalDescricao = xLocalDescricao
                
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
Private Sub HabilitaBotoes()

    'Objetivo: Habilita os Botões da Tela
    CmdInserir.Enabled = True
    If Not FunTabelaVazia(Adodc1) Then
        CmdAlterar.Enabled = True
        CmdExcluir.Enabled = True
        Cmdimprimir.Enabled = True
        CmdOrdem.Enabled = True
        CmdAtivar.Enabled = True
        cboLocalizar.Enabled = True
    Else
        CmdAlterar.Enabled = False
        CmdExcluir.Enabled = False
        Cmdimprimir.Enabled = False
        CmdOrdem.Enabled = False
        CmdAtivar.Enabled = False
        cboLocalizar.Enabled = False
    End If
    CmdSair.Enabled = True

    'Verifica se tem Filtro Ativo
    If ContFil = 0 Then
       Botao_Fil.Enabled = False
    Else
       Botao_Fil.Enabled = True
    End If
End Sub
Private Sub DesabilitaBotoes()
    'Objetivo: Desabilita os Botões da Tela
    CmdInserir.Enabled = False
    CmdAlterar.Enabled = False
    CmdExcluir.Enabled = False
    'Botao_Fil.Enabled = False
    CmdAtivar.Enabled = False
    CmdSair.Enabled = False
    Cmdimprimir.Enabled = False
    CmdOrdem.Enabled = False
End Sub
Private Sub Botao_Fil_Click()
    Dim Fil As Integer
    DesabilitaBotoes
    Botao_Fil.Enabled = False
    Set Formulario = TelaEmpresas
    TelaValoresGlobais "P"
    TelaFiltro.Show 1 'Mostra a tela de filtros no modo modal
    TelaValoresGlobais "G"

    If Filtrou = True Then
        Fil = ContFil
        subrecarregadadosNV Adodc1, "", FRM_Ordem, "", FRM_FiltroAtual
        ContFil = Fil
    End If
    
    HabilitaBotoes
    
    TDBGrid1.SetFocus
End Sub

Private Sub cboLocalizar_Click(Area As Integer)
    If Area = 2 Then
        Adodc1.Recordset.bookmark = cboLocalizar.SelectedItem
    End If
End Sub

Private Sub CmdAlterar_Click()
  Chave = Adodc1.Recordset.Fields("empr_cd_empresa")
  XInserir = "A"
  xChave = Adodc1.Recordset.bookmark
  DesabilitaBotoes
  FrmPadraoEmpresa.Show vbModal
  If XGB_RESPOSTA Then
    subrecarregadadosNV Adodc1, "", FRM_Ordem, "", FRM_FiltroAtual
  End If
  HabilitaBotoes
End Sub
Private Sub CmdAtivar_Click()

On Error GoTo TrataErro:

    Dim XLI_BOOKMARK As Integer
    
    'Função que verifica se o usuário está associado à empresa que deseja ativar
    If FunVerificaHabilitacaoEmpresa = False Then
        MsgBox "Você não está habilitado a ativar essa empresa!", vbInformation, "Permissão"
        Exit Sub
    End If
    
    If TDBGrid1.Columns("Empresa Ativa").Value = "N" Then
        PCodEmpresa = Adodc1.Recordset.Fields("empr_cd_empresa")
        PFantasia = Adodc1.Recordset.Fields("empr_tx_fantasia")
        PEmpresa = Adodc1.Recordset.Fields("empr_tx_razaosocial")
        PEndereco = FunNulo(Adodc1.Recordset.Fields("empr_tx_endereco"))
        pPCRmascara = Adodc1.Recordset.Fields("empr_tx_pCRMascara")
        pPCRnivel = Adodc1.Recordset.Fields("empr_nr_pcrnivel")
        pPlanoCmascara = Adodc1.Recordset.Fields("empr_tx_planocmascara")
        pPlanoCNivel = Adodc1.Recordset.Fields("empr_nr_planocnivel")
        PLogo = FunNulo(Adodc1.Recordset.Fields("empr_im_caminhoLogo"))

        XGI_PCATEGORIANVFINANCEIRO = FunNuloVal(Adodc1.Recordset.Fields("empr_nr_CategoriaNvContabil"))
        MdiPrincipal.BarraStatus.Panels(2).Text = PFantasia
    
        Conexao.BeginTrans
        Conexao.Execute ("UPDATE Empresas SET empr_tx_Atual = 'N' WHERE empr_tx_Atual = 'S'")
        Conexao.Execute ("UPDATE Empresas SET empr_tx_Atual = 'S' WHERE empr_cd_Empresa =" & PCodEmpresa)

        Conexao.CommitTrans
        
        XLI_BOOKMARK = TDBGrid1.bookmark
        subrecarregadadosNV Adodc1, "", FRM_Ordem, "", FRM_FiltroAtual
        TDBGrid1.bookmark = XLI_BOOKMARK
    Else
        MsgBox "Esta empresa já está ativa.", vbCritical + vbOKOnly, "ATENÇÃO"
    End If
    
    Exit Sub
TrataErro:
    Conexao.RollbackTrans
    funTrataErros (ComMensagem)
    Resume Next
End Sub
Private Sub CmdExcluir_Click()
    
    If CmdExcluir.Caption = "&Excluir" Then
        'If Not funVerifica_Permissao("FrmExemplo", OrdemRotina) Then Exit Sub
    End If

    Dim opcao As Integer
    Dim Houve_Erro As Integer
    Dim ResReg As Object
    Dim Posicao As Variant

    Houve_Erro = False
    
    Posicao = Adodc1.Recordset.bookmark
    DesabilitaBotoes
       
    'opcao =
   
    If MsgBox("Confirma Remoção do Registro ?", vbYesNo + vbCritical + vbDefaultButton2, "CUIDADO") = vbYes Then 'Confirma exclusão
'        On Error GoTo RotuloErro
            
        'Verifica se a empresa possui dados cadastrados
        SubQOpenRecordset ResReg, "SELECT * FROM NotasFiscais WHERE empr_cd_empresa =" & Adodc1.Recordset.Fields("empr_cd_empresa"), Estatico
        If Not ResReg.EOF Then
            MsgBox "A empresa possui movimentos cadastrados. Não é possível excluí-la.", vbCritical + vbOKOnly, "ATENÇÃO"
            ResReg.Close
            Exit Sub
        Else
            ResReg.Close
            Conexao.BeginTrans
                        
            'Deletar tabelas relacionadas
            Conexao.Execute "Delete from Plccustoforncli Where empr_cd_empresa=" & Adodc1.Recordset.Fields("empr_cd_empresa")
            Conexao.Execute "Delete from Plcdescontosimpostos Where empr_cd_empresa=" & Adodc1.Recordset.Fields("empr_cd_empresa")
            Conexao.Execute "Delete from Plcdescontosacrescimos Where empr_cd_empresa=" & Adodc1.Recordset.Fields("empr_cd_empresa")
            Conexao.Execute "Delete from Contascorrente Where empr_cd_empresa=" & Adodc1.Recordset.Fields("empr_cd_empresa")
            Conexao.Execute "Delete from Planosdiversos Where empr_cd_empresa=" & Adodc1.Recordset.Fields("empr_cd_empresa")
            Conexao.Execute "Delete from Centroresultados Where empr_cd_empresa=" & Adodc1.Recordset.Fields("empr_cd_empresa")
            Conexao.Execute "Delete from Planocontas Where empr_cd_empresa=" & Adodc1.Recordset.Fields("empr_cd_empresa")
            Conexao.Execute "Delete from CategoriaPlanoContas Where empr_cd_empresa=" & Adodc1.Recordset.Fields("empr_cd_empresa")
            Conexao.Execute "Delete from Fechamento Where empr_cd_empresa=" & Adodc1.Recordset.Fields("empr_cd_empresa")
            
            'delete tabela principal
            Conexao.Execute "Delete from empresas Where empr_cd_empresa=" & Adodc1.Recordset.Fields("empr_cd_empresa")
            
            Conexao.CommitTrans
            'Adodc1.Refresh TestesubrecarregadadosNV
            If Houve_Erro = False Then
                 FRM_Quantid = funOperacaoQuantid("-")
                 subrecarregadadosNV Adodc1, "", FRM_Ordem, "", FRM_FiltroAtual
            End If
         End If
    End If
    HabilitaBotoes
           
    TDBGrid1.SetFocus
    Adodc1.Refresh
    TDBGrid1.Refresh
    TDBGrid1.Col = 0
    
    'Posicionando o cursor na linha anterior - Lucas Santiago 22.01.09
    If Posicao > 1 Then
        Adodc1.Recordset.bookmark = Posicao - 1
    End If
    
    Exit Sub
    
RotuloErro:
    Conexao.RollbackTrans
    Houve_Erro = True
    funTrataErros (ComMensagem)
    Resume Next
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

Private Sub CmdImprimir_Click()
    DesabilitaBotoes
    Call subImprimeListagemGRID(2, TDBGrid1.PrintInfo, "Listagem de Empresas")
    HabilitaBotoes
End Sub
Private Sub CmdInserir_Click()
  Chave = -1
  XInserir = "I"
  DesabilitaBotoes
  FrmPadraoEmpresa.Show vbModal
  If XGB_RESPOSTA Then
    If Chave <> -1 Then
      FRM_Quantid = funOperacaoQuantid("+")
      subrecarregadadosNV Adodc1, "", FRM_Ordem, "", FRM_FiltroAtual
      cboLocalizar.BoundText = Chave
      If Not IsNull(cboLocalizar.SelectedItem) Then Adodc1.Recordset.bookmark = cboLocalizar.SelectedItem
    End If
  End If
  
  HabilitaBotoes
  
End Sub
Private Sub CmdOrdem_Click()
    
    DesabilitaBotoes
    Set Formulario = TelaEmpresas
    TelaValoresGlobais "P"
    TelaOrdem.Show 1
    TelaValoresGlobais "G"
    HabilitaBotoes
    
    lblLocalizar.Caption = "Localizar " & xLocalDescricao & ":"
    If InStr(1, xLocalCampo, ",") > 0 Then
        cboLocalizar.ListField = Left(xLocalCampo, InStr(1, xLocalCampo, ",") - 1)
    Else
        cboLocalizar.ListField = xLocalCampo
    End If
    cboLocalizar.Text = ""
       
    TDBGrid1.SetFocus
    
End Sub
Private Sub CmdSair_Click()
    Dim Ativa As Boolean
    MdiPrincipal.BarraStatus.Panels(2).Text = PFantasia
    MdiPrincipal.BarraStatus.Panels(1).Text = ""
    Ativa = False
    TDBGrid1.MoveFirst
    While Not TDBGrid1.EOF
        If TDBGrid1.Columns("Empresa Ativa").Value = "S" Then
            Ativa = True
        End If
        TDBGrid1.MoveNext
    Wend
    
    If Not Ativa Then
       MsgBox "É preciso ativar uma empresa.", vbCritical + vbOKOnly, "ATENÇÃO"
       Exit Sub
    End If
    MdiPrincipal.MousePointer = vbDefault
    Unload Me
End Sub
Private Sub TDBGrid1_DblClick()
  If Adodc1.Recordset.RecordCount <> 0 Then
    CmdAlterar_Click
  End If
End Sub

Private Sub TDBGrid1_FetchRowStyle(ByVal Split As Integer, bookmark As Variant, ByVal RowStyle As TrueOleDBGrid70.StyleDisp)
  If TDBGrid1.Columns("Empresa Ativa").CellText(bookmark) = "S" Then
    RowStyle.ForeColor = &HFF0000
  End If
End Sub

Private Sub TDBGrid1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    cboLocalizar.Text = ""
    If Button = 2 And CmdInserir.Caption = "&Inserir" Then  'Verifica se o botão da direita foi pressionado
        Set Formulario = TelaEmpresas
        TelaValoresGlobais "P"
        FunExecutaFiltro Adodc1, Formulario, x
        TelaValoresGlobais "G"
    End If
    
    HabilitaBotoes
    
End Sub

Private Sub Form_Activate()
  Adodc1.Refresh
  Set Formulario = TelaEmpresas
  If Not PrimeiraVez Then
    TelaValoresGlobais "P"
  End If
  PrimeiraVez = False
  MdiPrincipal.BarraStatus.Panels(3).Text = CStr(Adodc1.Recordset.RecordCount) + "/" + CStr(Adodc1.Recordset.RecordCount)
End Sub

Private Sub Form_Load()
  Set Formulario = Me
  Call subAjustaTelaLoc
  subManutencaoJanelasAtivas "I", "TelaEmpresas"
  xLocalDescricao = "Fantasia"
  subCarregaDadosNV Adodc1, "", "empr_tx_fantasia", "", "SELECT * FROM ConsGENEmpresa "
  HabilitaBotoes
  filtra_especial = False
  TelaValoresGlobais "G"
  PrimeiraVez = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    subManutencaoJanelasAtivas "R", "TelaEmpresas"
End Sub
