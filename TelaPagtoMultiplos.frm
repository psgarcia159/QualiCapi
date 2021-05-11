VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0D623638-DBA2-11D1-B5DF-0060976089D0}#7.0#0"; "tdbg7.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form TelaFinPagtoMultiplos 
   Caption         =   "Pagamento Múltiplos de Títulos"
   ClientHeight    =   4500
   ClientLeft      =   1245
   ClientTop       =   2925
   ClientWidth     =   8820
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4500
   ScaleWidth      =   8820
   WindowState     =   2  'Maximized
   Begin TrueDBGrid70.TDBGrid TDBGrid1 
      Height          =   1815
      Left            =   15
      TabIndex        =   11
      Top             =   2010
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   3201
      _LayoutType     =   4
      _RowHeight      =   19
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Vencimento"
      Columns(0).FooterText=   "D"
      Columns(0).DataField=   "titu_dt_prorrogacao"
      Columns(0).NumberFormat=   "dd/mm/yy"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Título"
      Columns(1).FooterText=   "T"
      Columns(1).DataField=   "Titulo"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Tipo Plano"
      Columns(2).FooterText=   "T"
      Columns(2).DataField=   "tipl_tx_descricao"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Natureza"
      Columns(3).FooterText=   "T"
      Columns(3).DataField=   "natureza"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   80
      Columns(4)._MaxComboItems=   5
      Columns(4).ValueItems(0)._DefaultItem=   0
      Columns(4).ValueItems(0).Value=   "N"
      Columns(4).ValueItems(0).Value.vt=   8
      Columns(4).ValueItems(0).DisplayValue.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
      Columns(4).ValueItems(0).DisplayValue(0)=   "bHQAACYEAABCTSYEAAAAAAAANgAAACgAAAASAAAAEgAAAAEAGAAAAAAA8AMAAAAAAAAAAAAAAAAA"
      Columns(4).ValueItems(0).DisplayValue(1)=   "AAAAAAD/////////////////////////////////////////////////////////////////////"
      Columns(4).ValueItems(0).DisplayValue(2)=   "//8AAP//////////////////////////////////////////////////////////////////////"
      Columns(4).ValueItems(0).DisplayValue(3)=   "/wAA////////////////////////////////////////////////////////////////////////"
      Columns(4).ValueItems(0).DisplayValue(4)=   "AAD///////////////////////////////////////////////////////////////////////8A"
      Columns(4).ValueItems(0).DisplayValue(5)=   "AP///////////4CAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgP///////////////wAA"
      Columns(4).ValueItems(0).DisplayValue(6)=   "////////////gICA////////////////////////////////////gICA////////////////AAD/"
      Columns(4).ValueItems(0).DisplayValue(7)=   "//////////+AgID///////////////////////////////////+AgID///////////////8AAP//"
      Columns(4).ValueItems(0).DisplayValue(8)=   "/////////4CAgP///////////////////////////////////4CAgP///////////////wAA////"
      Columns(4).ValueItems(0).DisplayValue(9)=   "////////gICA////////////////////////////////////gICA////////////////AAD/////"
      Columns(4).ValueItems(0).DisplayValue(10)=   "//////+AgID///////////////////////////////////+AgID///////////////8AAP//////"
      Columns(4).ValueItems(0).DisplayValue(11)=   "/////4CAgP///////////////////////////////////4CAgP///////////////wAA////////"
      Columns(4).ValueItems(0).DisplayValue(12)=   "////gICA////////////////////////////////////gICA////////////////AAD/////////"
      Columns(4).ValueItems(0).DisplayValue(13)=   "//+AgID///////////////////////////////////+AgID///////////////8AAP//////////"
      Columns(4).ValueItems(0).DisplayValue(14)=   "/4CAgP///////////////////////////////////4CAgP///////////////wAA////////////"
      Columns(4).ValueItems(0).DisplayValue(15)=   "gICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICA////////////////AAD/////////////"
      Columns(4).ValueItems(0).DisplayValue(16)=   "//////////////////////////////////////////////////////////8AAP//////////////"
      Columns(4).ValueItems(0).DisplayValue(17)=   "/////////////////////////////////////////////////////////wAA////////////////"
      Columns(4).ValueItems(0).DisplayValue(18)=   "////////////////////////////////////////////////////////AAA="
      Columns(4).ValueItems(0).DisplayValue.vt=   9
      Columns(4).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
      Columns(4).ValueItems(1)._DefaultItem=   -1
      Columns(4).ValueItems(1).Value=   "S"
      Columns(4).ValueItems(1).Value.vt=   8
      Columns(4).ValueItems(1).DisplayValue.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
      Columns(4).ValueItems(1).DisplayValue(0)=   "bHQAABIEAABCTRIEAAAAAAAANgAAACgAAAARAAAAEwAAAAEAGAAAAAAA3AMAAAAAAAAAAAAAAAAA"
      Columns(4).ValueItems(1).DisplayValue(1)=   "AAAAAAD///////////////////////////////////////////////////////////////////8A"
      Columns(4).ValueItems(1).DisplayValue(2)=   "////////////////////////////////////////////////////////////////////AP//////"
      Columns(4).ValueItems(1).DisplayValue(3)=   "/////////////////////////////////////////////////////////////wD/////////////"
      Columns(4).ValueItems(1).DisplayValue(4)=   "//////////////////////////////////////////////////////8A////////////gICAgICA"
      Columns(4).ValueItems(1).DisplayValue(5)=   "gICAgICAgICAgICAgICAgICAgICAgICAgICA////////////AP///////////4CAgP//////////"
      Columns(4).ValueItems(1).DisplayValue(6)=   "/////////////////////////4CAgP///////////wD///////////+AgID///////////8AAAD/"
      Columns(4).ValueItems(1).DisplayValue(7)=   "//////////////////+AgID///////////8A////////////gICA////////AAAAAAAAAAAA////"
      Columns(4).ValueItems(1).DisplayValue(8)=   "////////////gICA////////////AP///////////4CAgP///wAAAAAAAAAAAAAAAAAAAP//////"
      Columns(4).ValueItems(1).DisplayValue(9)=   "/////4CAgP///////////wD///////////+AgID///8AAAAAAAD///8AAAAAAAAAAAD///////+A"
      Columns(4).ValueItems(1).DisplayValue(10)=   "gID///////////8A////////////gICA////AAAA////////////AAAAAAAAAAAA////gICA////"
      Columns(4).ValueItems(1).DisplayValue(11)=   "////////AP///////////4CAgP///////////////////////wAAAAAAAP///4CAgP//////////"
      Columns(4).ValueItems(1).DisplayValue(12)=   "/wD///////////+AgID///////////////////////////8AAAD///+AgID///////////8A////"
      Columns(4).ValueItems(1).DisplayValue(13)=   "////////gICA////////////////////////////////////gICA////////////AP//////////"
      Columns(4).ValueItems(1).DisplayValue(14)=   "/4CAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgP///////////wD/////////////////"
      Columns(4).ValueItems(1).DisplayValue(15)=   "//////////////////////////////////////////////////8A////////////////////////"
      Columns(4).ValueItems(1).DisplayValue(16)=   "////////////////////////////////////////////AP//////////////////////////////"
      Columns(4).ValueItems(1).DisplayValue(17)=   "/////////////////////////////////////wD/////////////////////////////////////"
      Columns(4).ValueItems(1).DisplayValue(18)=   "//////////////////////////////8A"
      Columns(4).ValueItems(1).DisplayValue.vt=   9
      Columns(4).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
      Columns(4).ValueItems.Count=   2
      Columns(4).Caption=   "Sel."
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Valor Histórico"
      Columns(5).FooterText=   "N"
      Columns(5).DataField=   "titu_vl_parcela"
      Columns(5).NumberFormat=   "Standard"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "Correção"
      Columns(6).FooterText=   "N"
      Columns(6).DataField=   "titu_vl_ParcelaIndex1"
      Columns(6).NumberFormat=   "Standard"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "Valor Atual"
      Columns(7).FooterText=   "N"
      Columns(7).DataField=   ""
      Columns(7).NumberFormat=   "Standard"
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "Juros"
      Columns(8).FooterText=   "N"
      Columns(8).DataField=   ""
      Columns(8).NumberFormat=   "Standard"
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "Multa"
      Columns(9).FooterText=   "N"
      Columns(9).DataField=   ""
      Columns(9).NumberFormat=   "Standard"
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "Deságio"
      Columns(10).FooterText=   "N"
      Columns(10).DataField=   ""
      Columns(10).NumberFormat=   "Standard"
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "Desconto"
      Columns(11).FooterText=   "N"
      Columns(11).DataField=   ""
      Columns(11).NumberFormat=   "Standard"
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).Caption=   "Outros(+)"
      Columns(12).FooterText=   "N"
      Columns(12).DataField=   ""
      Columns(12).NumberFormat=   "Standard"
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(13)._VlistStyle=   0
      Columns(13)._MaxComboItems=   5
      Columns(13).Caption=   "Valor Total"
      Columns(13).FooterText=   "N"
      Columns(13).DataField=   ""
      Columns(13).NumberFormat=   "Standard"
      Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(14)._VlistStyle=   0
      Columns(14)._MaxComboItems=   5
      Columns(14).Caption=   "Valor a Pagar"
      Columns(14).FooterText=   "N"
      Columns(14).DataField=   ""
      Columns(14).NumberFormat=   "Standard"
      Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(15)._VlistStyle=   0
      Columns(15)._MaxComboItems=   5
      Columns(15).Caption=   "Moeda1"
      Columns(15).DataField=   "moed_cd_Moeda1"
      Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(16)._VlistStyle=   0
      Columns(16)._MaxComboItems=   5
      Columns(16).Caption=   "Moeda2"
      Columns(16).DataField=   "moed_cd_Moeda2"
      Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(17)._VlistStyle=   0
      Columns(17)._MaxComboItems=   5
      Columns(17).Caption=   "CCusto"
      Columns(17).DataField=   ""
      Columns(17)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(18)._VlistStyle=   0
      Columns(18)._MaxComboItems=   5
      Columns(18).Caption=   "CConta"
      Columns(18).DataField=   ""
      Columns(18)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(19)._VlistStyle=   0
      Columns(19)._MaxComboItems=   5
      Columns(19).Caption=   "Prorrogação"
      Columns(19).DataField=   ""
      Columns(19)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(20)._VlistStyle=   0
      Columns(20)._MaxComboItems=   5
      Columns(20).Caption=   "Observação"
      Columns(20).DataField=   ""
      Columns(20)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(21)._VlistStyle=   0
      Columns(21)._MaxComboItems=   5
      Columns(21).Caption=   "titu_tx_naturezaplano"
      Columns(21).DataField=   ""
      Columns(21)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(22)._VlistStyle=   0
      Columns(22)._MaxComboItems=   5
      Columns(22).Caption=   "titu_vl_PercJurosFin"
      Columns(22).DataField=   ""
      Columns(22)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(23)._VlistStyle=   0
      Columns(23)._MaxComboItems=   5
      Columns(23).Caption=   "titu_vl_PercMultaMora"
      Columns(23).DataField=   ""
      Columns(23)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(24)._VlistStyle=   0
      Columns(24)._MaxComboItems=   5
      Columns(24).Caption=   "titu_vl_PercJurosMora"
      Columns(24).DataField=   ""
      Columns(24)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(25)._VlistStyle=   0
      Columns(25)._MaxComboItems=   5
      Columns(25).Caption=   "iden_cd_Projeto"
      Columns(25).DataField=   ""
      Columns(25)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   26
      Splits(0)._UserFlags=   0
      Splits(0).ExtendRightColumn=   -1  'True
      Splits(0).MarqueeStyle=   4
      Splits(0).AllowSizing=   -1  'True
      Splits(0).RecordSelectorWidth=   503
      Splits(0).AllowColMove=   -1  'True
      Splits(0).FetchRowStyle=   -1  'True
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=26"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2037"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1958"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=74257"
      Splits(0)._ColumnProps(6)=   "Column(0).WrapText=1"
      Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(8)=   "Column(1).Width=1773"
      Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=1693"
      Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=74256"
      Splits(0)._ColumnProps(13)=   "Column(1).WrapText=1"
      Splits(0)._ColumnProps(14)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(15)=   "Column(2).Width=2725"
      Splits(0)._ColumnProps(16)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(17)=   "Column(2)._WidthInPix=2646"
      Splits(0)._ColumnProps(18)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(19)=   "Column(2).AllowSizing=0"
      Splits(0)._ColumnProps(20)=   "Column(2)._ColStyle=66064"
      Splits(0)._ColumnProps(21)=   "Column(2).Visible=0"
      Splits(0)._ColumnProps(22)=   "Column(2).WrapText=1"
      Splits(0)._ColumnProps(23)=   "Column(2).AllowFocus=0"
      Splits(0)._ColumnProps(24)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(25)=   "Column(3).Width=2725"
      Splits(0)._ColumnProps(26)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(27)=   "Column(3)._WidthInPix=2646"
      Splits(0)._ColumnProps(28)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(29)=   "Column(3)._ColStyle=66064"
      Splits(0)._ColumnProps(30)=   "Column(3).Visible=0"
      Splits(0)._ColumnProps(31)=   "Column(3).WrapText=1"
      Splits(0)._ColumnProps(32)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(33)=   "Column(4).Width=1773"
      Splits(0)._ColumnProps(34)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(35)=   "Column(4)._WidthInPix=1693"
      Splits(0)._ColumnProps(36)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(37)=   "Column(4)._ColStyle=66065"
      Splits(0)._ColumnProps(38)=   "Column(4).Visible=0"
      Splits(0)._ColumnProps(39)=   "Column(4).WrapText=1"
      Splits(0)._ColumnProps(40)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(41)=   "Column(4)._MinWidth=56843344"
      Splits(0)._ColumnProps(42)=   "Column(5).Width=2381"
      Splits(0)._ColumnProps(43)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(44)=   "Column(5)._WidthInPix=2302"
      Splits(0)._ColumnProps(45)=   "Column(5)._EditAlways=0"
      Splits(0)._ColumnProps(46)=   "Column(5)._ColStyle=66066"
      Splits(0)._ColumnProps(47)=   "Column(5).Visible=0"
      Splits(0)._ColumnProps(48)=   "Column(5).WrapText=1"
      Splits(0)._ColumnProps(49)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(50)=   "Column(5)._MinWidth=56841056"
      Splits(0)._ColumnProps(51)=   "Column(6).Width=1958"
      Splits(0)._ColumnProps(52)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(53)=   "Column(6)._WidthInPix=1879"
      Splits(0)._ColumnProps(54)=   "Column(6)._EditAlways=0"
      Splits(0)._ColumnProps(55)=   "Column(6)._ColStyle=66066"
      Splits(0)._ColumnProps(56)=   "Column(6).Visible=0"
      Splits(0)._ColumnProps(57)=   "Column(6).WrapText=1"
      Splits(0)._ColumnProps(58)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(59)=   "Column(6)._MinWidth=56841200"
      Splits(0)._ColumnProps(60)=   "Column(7).Width=2196"
      Splits(0)._ColumnProps(61)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(62)=   "Column(7)._WidthInPix=2117"
      Splits(0)._ColumnProps(63)=   "Column(7)._EditAlways=0"
      Splits(0)._ColumnProps(64)=   "Column(7)._ColStyle=66066"
      Splits(0)._ColumnProps(65)=   "Column(7).Visible=0"
      Splits(0)._ColumnProps(66)=   "Column(7).WrapText=1"
      Splits(0)._ColumnProps(67)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(68)=   "Column(8).Width=2275"
      Splits(0)._ColumnProps(69)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(70)=   "Column(8)._WidthInPix=2196"
      Splits(0)._ColumnProps(71)=   "Column(8)._EditAlways=0"
      Splits(0)._ColumnProps(72)=   "Column(8)._ColStyle=66066"
      Splits(0)._ColumnProps(73)=   "Column(8).Visible=0"
      Splits(0)._ColumnProps(74)=   "Column(8).WrapText=1"
      Splits(0)._ColumnProps(75)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(76)=   "Column(9).Width=2143"
      Splits(0)._ColumnProps(77)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(78)=   "Column(9)._WidthInPix=2064"
      Splits(0)._ColumnProps(79)=   "Column(9)._EditAlways=0"
      Splits(0)._ColumnProps(80)=   "Column(9)._ColStyle=66066"
      Splits(0)._ColumnProps(81)=   "Column(9).Visible=0"
      Splits(0)._ColumnProps(82)=   "Column(9).WrapText=1"
      Splits(0)._ColumnProps(83)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(84)=   "Column(10).Width=2170"
      Splits(0)._ColumnProps(85)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(86)=   "Column(10)._WidthInPix=2090"
      Splits(0)._ColumnProps(87)=   "Column(10)._EditAlways=0"
      Splits(0)._ColumnProps(88)=   "Column(10)._ColStyle=66066"
      Splits(0)._ColumnProps(89)=   "Column(10).Visible=0"
      Splits(0)._ColumnProps(90)=   "Column(10).WrapText=1"
      Splits(0)._ColumnProps(91)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(92)=   "Column(11).Width=2170"
      Splits(0)._ColumnProps(93)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(94)=   "Column(11)._WidthInPix=2090"
      Splits(0)._ColumnProps(95)=   "Column(11)._EditAlways=0"
      Splits(0)._ColumnProps(96)=   "Column(11)._ColStyle=66066"
      Splits(0)._ColumnProps(97)=   "Column(11).Visible=0"
      Splits(0)._ColumnProps(98)=   "Column(11).WrapText=1"
      Splits(0)._ColumnProps(99)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(100)=   "Column(12).Width=2223"
      Splits(0)._ColumnProps(101)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(102)=   "Column(12)._WidthInPix=2143"
      Splits(0)._ColumnProps(103)=   "Column(12)._EditAlways=0"
      Splits(0)._ColumnProps(104)=   "Column(12)._ColStyle=66066"
      Splits(0)._ColumnProps(105)=   "Column(12).Visible=0"
      Splits(0)._ColumnProps(106)=   "Column(12).WrapText=1"
      Splits(0)._ColumnProps(107)=   "Column(12).Order=13"
      Splits(0)._ColumnProps(108)=   "Column(13).Width=2725"
      Splits(0)._ColumnProps(109)=   "Column(13).DividerColor=0"
      Splits(0)._ColumnProps(110)=   "Column(13)._WidthInPix=2646"
      Splits(0)._ColumnProps(111)=   "Column(13)._EditAlways=0"
      Splits(0)._ColumnProps(112)=   "Column(13).AllowSizing=0"
      Splits(0)._ColumnProps(113)=   "Column(13)._ColStyle=74000"
      Splits(0)._ColumnProps(114)=   "Column(13).Visible=0"
      Splits(0)._ColumnProps(115)=   "Column(13).WrapText=1"
      Splits(0)._ColumnProps(116)=   "Column(13).Order=14"
      Splits(0)._ColumnProps(117)=   "Column(14).Width=2725"
      Splits(0)._ColumnProps(118)=   "Column(14).DividerColor=0"
      Splits(0)._ColumnProps(119)=   "Column(14)._WidthInPix=2646"
      Splits(0)._ColumnProps(120)=   "Column(14)._EditAlways=0"
      Splits(0)._ColumnProps(121)=   "Column(14)._ColStyle=66066"
      Splits(0)._ColumnProps(122)=   "Column(14).Visible=0"
      Splits(0)._ColumnProps(123)=   "Column(14).WrapText=1"
      Splits(0)._ColumnProps(124)=   "Column(14).Order=15"
      Splits(0)._ColumnProps(125)=   "Column(15).Width=2725"
      Splits(0)._ColumnProps(126)=   "Column(15).DividerColor=0"
      Splits(0)._ColumnProps(127)=   "Column(15)._WidthInPix=2646"
      Splits(0)._ColumnProps(128)=   "Column(15)._EditAlways=0"
      Splits(0)._ColumnProps(129)=   "Column(15).AllowSizing=0"
      Splits(0)._ColumnProps(130)=   "Column(15)._ColStyle=65808"
      Splits(0)._ColumnProps(131)=   "Column(15).Visible=0"
      Splits(0)._ColumnProps(132)=   "Column(15).WrapText=1"
      Splits(0)._ColumnProps(133)=   "Column(15).Order=16"
      Splits(0)._ColumnProps(134)=   "Column(15)._MinWidth=54106848"
      Splits(0)._ColumnProps(135)=   "Column(16).Width=2725"
      Splits(0)._ColumnProps(136)=   "Column(16).DividerColor=0"
      Splits(0)._ColumnProps(137)=   "Column(16)._WidthInPix=2646"
      Splits(0)._ColumnProps(138)=   "Column(16)._EditAlways=0"
      Splits(0)._ColumnProps(139)=   "Column(16).AllowSizing=0"
      Splits(0)._ColumnProps(140)=   "Column(16)._ColStyle=65808"
      Splits(0)._ColumnProps(141)=   "Column(16).Visible=0"
      Splits(0)._ColumnProps(142)=   "Column(16).WrapText=1"
      Splits(0)._ColumnProps(143)=   "Column(16).Order=17"
      Splits(0)._ColumnProps(144)=   "Column(16)._MinWidth=54105776"
      Splits(0)._ColumnProps(145)=   "Column(17).Width=2725"
      Splits(0)._ColumnProps(146)=   "Column(17).DividerColor=0"
      Splits(0)._ColumnProps(147)=   "Column(17)._WidthInPix=2646"
      Splits(0)._ColumnProps(148)=   "Column(17)._EditAlways=0"
      Splits(0)._ColumnProps(149)=   "Column(17)._ColStyle=65808"
      Splits(0)._ColumnProps(150)=   "Column(17).Visible=0"
      Splits(0)._ColumnProps(151)=   "Column(17).WrapText=1"
      Splits(0)._ColumnProps(152)=   "Column(17).Order=18"
      Splits(0)._ColumnProps(153)=   "Column(18).Width=2725"
      Splits(0)._ColumnProps(154)=   "Column(18).DividerColor=0"
      Splits(0)._ColumnProps(155)=   "Column(18)._WidthInPix=2646"
      Splits(0)._ColumnProps(156)=   "Column(18)._EditAlways=0"
      Splits(0)._ColumnProps(157)=   "Column(18)._ColStyle=65808"
      Splits(0)._ColumnProps(158)=   "Column(18).Visible=0"
      Splits(0)._ColumnProps(159)=   "Column(18).WrapText=1"
      Splits(0)._ColumnProps(160)=   "Column(18).Order=19"
      Splits(0)._ColumnProps(161)=   "Column(19).Width=2725"
      Splits(0)._ColumnProps(162)=   "Column(19).DividerColor=0"
      Splits(0)._ColumnProps(163)=   "Column(19)._WidthInPix=2646"
      Splits(0)._ColumnProps(164)=   "Column(19)._EditAlways=0"
      Splits(0)._ColumnProps(165)=   "Column(19)._ColStyle=65808"
      Splits(0)._ColumnProps(166)=   "Column(19).Visible=0"
      Splits(0)._ColumnProps(167)=   "Column(19).WrapText=1"
      Splits(0)._ColumnProps(168)=   "Column(19).Order=20"
      Splits(0)._ColumnProps(169)=   "Column(20).Width=2725"
      Splits(0)._ColumnProps(170)=   "Column(20).DividerColor=0"
      Splits(0)._ColumnProps(171)=   "Column(20)._WidthInPix=2646"
      Splits(0)._ColumnProps(172)=   "Column(20)._EditAlways=0"
      Splits(0)._ColumnProps(173)=   "Column(20)._ColStyle=65808"
      Splits(0)._ColumnProps(174)=   "Column(20).Visible=0"
      Splits(0)._ColumnProps(175)=   "Column(20).WrapText=1"
      Splits(0)._ColumnProps(176)=   "Column(20).Order=21"
      Splits(0)._ColumnProps(177)=   "Column(21).Width=2725"
      Splits(0)._ColumnProps(178)=   "Column(21).DividerColor=0"
      Splits(0)._ColumnProps(179)=   "Column(21)._WidthInPix=2646"
      Splits(0)._ColumnProps(180)=   "Column(21)._EditAlways=0"
      Splits(0)._ColumnProps(181)=   "Column(21)._ColStyle=65808"
      Splits(0)._ColumnProps(182)=   "Column(21).Visible=0"
      Splits(0)._ColumnProps(183)=   "Column(21).WrapText=1"
      Splits(0)._ColumnProps(184)=   "Column(21).Order=22"
      Splits(0)._ColumnProps(185)=   "Column(22).Width=2725"
      Splits(0)._ColumnProps(186)=   "Column(22).DividerColor=0"
      Splits(0)._ColumnProps(187)=   "Column(22)._WidthInPix=2646"
      Splits(0)._ColumnProps(188)=   "Column(22)._EditAlways=0"
      Splits(0)._ColumnProps(189)=   "Column(22)._ColStyle=65808"
      Splits(0)._ColumnProps(190)=   "Column(22).Visible=0"
      Splits(0)._ColumnProps(191)=   "Column(22).WrapText=1"
      Splits(0)._ColumnProps(192)=   "Column(22).Order=23"
      Splits(0)._ColumnProps(193)=   "Column(23).Width=2725"
      Splits(0)._ColumnProps(194)=   "Column(23).DividerColor=0"
      Splits(0)._ColumnProps(195)=   "Column(23)._WidthInPix=2646"
      Splits(0)._ColumnProps(196)=   "Column(23)._EditAlways=0"
      Splits(0)._ColumnProps(197)=   "Column(23)._ColStyle=65808"
      Splits(0)._ColumnProps(198)=   "Column(23).Visible=0"
      Splits(0)._ColumnProps(199)=   "Column(23).WrapText=1"
      Splits(0)._ColumnProps(200)=   "Column(23).Order=24"
      Splits(0)._ColumnProps(201)=   "Column(24).Width=2725"
      Splits(0)._ColumnProps(202)=   "Column(24).DividerColor=0"
      Splits(0)._ColumnProps(203)=   "Column(24)._WidthInPix=2646"
      Splits(0)._ColumnProps(204)=   "Column(24)._EditAlways=0"
      Splits(0)._ColumnProps(205)=   "Column(24)._ColStyle=65808"
      Splits(0)._ColumnProps(206)=   "Column(24).Visible=0"
      Splits(0)._ColumnProps(207)=   "Column(24).WrapText=1"
      Splits(0)._ColumnProps(208)=   "Column(24).Order=25"
      Splits(0)._ColumnProps(209)=   "Column(25).Width=2725"
      Splits(0)._ColumnProps(210)=   "Column(25).DividerColor=0"
      Splits(0)._ColumnProps(211)=   "Column(25)._WidthInPix=2646"
      Splits(0)._ColumnProps(212)=   "Column(25)._EditAlways=0"
      Splits(0)._ColumnProps(213)=   "Column(25)._ColStyle=65808"
      Splits(0)._ColumnProps(214)=   "Column(25).Visible=0"
      Splits(0)._ColumnProps(215)=   "Column(25).WrapText=1"
      Splits(0)._ColumnProps(216)=   "Column(25).Order=26"
      Splits(1)._UserFlags=   0
      Splits(1).ExtendRightColumn=   -1  'True
      Splits(1).MarqueeStyle=   4
      Splits(1).RecordSelectors=   0   'False
      Splits(1).RecordSelectorWidth=   503
      Splits(1).AllowColMove=   -1  'True
      Splits(1).FetchRowStyle=   -1  'True
      Splits(1).DividerStyle=   4
      Splits(1).DividerColor=   12632256
      Splits(1).SpringMode=   0   'False
      Splits(1)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(1)._ColumnProps(0)=   "Columns.Count=26"
      Splits(1)._ColumnProps(1)=   "Column(0).Width=1958"
      Splits(1)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(1)._ColumnProps(3)=   "Column(0)._WidthInPix=1879"
      Splits(1)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(1)._ColumnProps(5)=   "Column(0).AllowSizing=0"
      Splits(1)._ColumnProps(6)=   "Column(0)._ColStyle=66065"
      Splits(1)._ColumnProps(7)=   "Column(0).Visible=0"
      Splits(1)._ColumnProps(8)=   "Column(0).WrapText=1"
      Splits(1)._ColumnProps(9)=   "Column(0).Order=1"
      Splits(1)._ColumnProps(10)=   "Column(0)._MinWidth=87714576"
      Splits(1)._ColumnProps(11)=   "Column(1).Width=3387"
      Splits(1)._ColumnProps(12)=   "Column(1).DividerColor=0"
      Splits(1)._ColumnProps(13)=   "Column(1)._WidthInPix=3307"
      Splits(1)._ColumnProps(14)=   "Column(1)._EditAlways=0"
      Splits(1)._ColumnProps(15)=   "Column(1).AllowSizing=0"
      Splits(1)._ColumnProps(16)=   "Column(1)._ColStyle=66064"
      Splits(1)._ColumnProps(17)=   "Column(1).Visible=0"
      Splits(1)._ColumnProps(18)=   "Column(1).WrapText=1"
      Splits(1)._ColumnProps(19)=   "Column(1).AllowFocus=0"
      Splits(1)._ColumnProps(20)=   "Column(1).Order=2"
      Splits(1)._ColumnProps(21)=   "Column(1)._MinWidth=87714576"
      Splits(1)._ColumnProps(22)=   "Column(2).Width=1984"
      Splits(1)._ColumnProps(23)=   "Column(2).DividerColor=0"
      Splits(1)._ColumnProps(24)=   "Column(2)._WidthInPix=1905"
      Splits(1)._ColumnProps(25)=   "Column(2)._EditAlways=0"
      Splits(1)._ColumnProps(26)=   "Column(2)._ColStyle=74256"
      Splits(1)._ColumnProps(27)=   "Column(2).WrapText=1"
      Splits(1)._ColumnProps(28)=   "Column(2).Order=3"
      Splits(1)._ColumnProps(29)=   "Column(3).Width=2487"
      Splits(1)._ColumnProps(30)=   "Column(3).DividerColor=0"
      Splits(1)._ColumnProps(31)=   "Column(3)._WidthInPix=2408"
      Splits(1)._ColumnProps(32)=   "Column(3)._EditAlways=0"
      Splits(1)._ColumnProps(33)=   "Column(3)._ColStyle=74256"
      Splits(1)._ColumnProps(34)=   "Column(3).WrapText=1"
      Splits(1)._ColumnProps(35)=   "Column(3).Order=4"
      Splits(1)._ColumnProps(36)=   "Column(4).Width=741"
      Splits(1)._ColumnProps(37)=   "Column(4).DividerColor=0"
      Splits(1)._ColumnProps(38)=   "Column(4)._WidthInPix=661"
      Splits(1)._ColumnProps(39)=   "Column(4)._EditAlways=0"
      Splits(1)._ColumnProps(40)=   "Column(4)._ColStyle=66065"
      Splits(1)._ColumnProps(41)=   "Column(4).WrapText=1"
      Splits(1)._ColumnProps(42)=   "Column(4).Order=5"
      Splits(1)._ColumnProps(43)=   "Column(5).Width=2381"
      Splits(1)._ColumnProps(44)=   "Column(5).DividerColor=0"
      Splits(1)._ColumnProps(45)=   "Column(5)._WidthInPix=2302"
      Splits(1)._ColumnProps(46)=   "Column(5)._EditAlways=0"
      Splits(1)._ColumnProps(47)=   "Column(5)._ColStyle=74258"
      Splits(1)._ColumnProps(48)=   "Column(5).WrapText=1"
      Splits(1)._ColumnProps(49)=   "Column(5).Order=6"
      Splits(1)._ColumnProps(50)=   "Column(6).Width=1958"
      Splits(1)._ColumnProps(51)=   "Column(6).DividerColor=0"
      Splits(1)._ColumnProps(52)=   "Column(6)._WidthInPix=1879"
      Splits(1)._ColumnProps(53)=   "Column(6)._EditAlways=0"
      Splits(1)._ColumnProps(54)=   "Column(6)._ColStyle=74258"
      Splits(1)._ColumnProps(55)=   "Column(6).WrapText=1"
      Splits(1)._ColumnProps(56)=   "Column(6).Order=7"
      Splits(1)._ColumnProps(57)=   "Column(7).Width=2196"
      Splits(1)._ColumnProps(58)=   "Column(7).DividerColor=0"
      Splits(1)._ColumnProps(59)=   "Column(7)._WidthInPix=2117"
      Splits(1)._ColumnProps(60)=   "Column(7)._EditAlways=0"
      Splits(1)._ColumnProps(61)=   "Column(7)._ColStyle=74258"
      Splits(1)._ColumnProps(62)=   "Column(7).WrapText=1"
      Splits(1)._ColumnProps(63)=   "Column(7).Order=8"
      Splits(1)._ColumnProps(64)=   "Column(8).Width=1640"
      Splits(1)._ColumnProps(65)=   "Column(8).DividerColor=0"
      Splits(1)._ColumnProps(66)=   "Column(8)._WidthInPix=1561"
      Splits(1)._ColumnProps(67)=   "Column(8)._EditAlways=0"
      Splits(1)._ColumnProps(68)=   "Column(8)._ColStyle=66066"
      Splits(1)._ColumnProps(69)=   "Column(8).WrapText=1"
      Splits(1)._ColumnProps(70)=   "Column(8).Order=9"
      Splits(1)._ColumnProps(71)=   "Column(8)._MinWidth=116881424"
      Splits(1)._ColumnProps(72)=   "Column(9).Width=1640"
      Splits(1)._ColumnProps(73)=   "Column(9).DividerColor=0"
      Splits(1)._ColumnProps(74)=   "Column(9)._WidthInPix=1561"
      Splits(1)._ColumnProps(75)=   "Column(9)._EditAlways=0"
      Splits(1)._ColumnProps(76)=   "Column(9)._ColStyle=66066"
      Splits(1)._ColumnProps(77)=   "Column(9).WrapText=1"
      Splits(1)._ColumnProps(78)=   "Column(9).Order=10"
      Splits(1)._ColumnProps(79)=   "Column(9)._MinWidth=116865856"
      Splits(1)._ColumnProps(80)=   "Column(10).Width=1561"
      Splits(1)._ColumnProps(81)=   "Column(10).DividerColor=0"
      Splits(1)._ColumnProps(82)=   "Column(10)._WidthInPix=1482"
      Splits(1)._ColumnProps(83)=   "Column(10)._EditAlways=0"
      Splits(1)._ColumnProps(84)=   "Column(10)._ColStyle=74258"
      Splits(1)._ColumnProps(85)=   "Column(10).WrapText=1"
      Splits(1)._ColumnProps(86)=   "Column(10).Order=11"
      Splits(1)._ColumnProps(87)=   "Column(10)._MinWidth=116879424"
      Splits(1)._ColumnProps(88)=   "Column(11).Width=1773"
      Splits(1)._ColumnProps(89)=   "Column(11).DividerColor=0"
      Splits(1)._ColumnProps(90)=   "Column(11)._WidthInPix=1693"
      Splits(1)._ColumnProps(91)=   "Column(11)._EditAlways=0"
      Splits(1)._ColumnProps(92)=   "Column(11)._ColStyle=66066"
      Splits(1)._ColumnProps(93)=   "Column(11).WrapText=1"
      Splits(1)._ColumnProps(94)=   "Column(11).Order=12"
      Splits(1)._ColumnProps(95)=   "Column(12).Width=1746"
      Splits(1)._ColumnProps(96)=   "Column(12).DividerColor=0"
      Splits(1)._ColumnProps(97)=   "Column(12)._WidthInPix=1667"
      Splits(1)._ColumnProps(98)=   "Column(12)._EditAlways=0"
      Splits(1)._ColumnProps(99)=   "Column(12)._ColStyle=66066"
      Splits(1)._ColumnProps(100)=   "Column(12).WrapText=1"
      Splits(1)._ColumnProps(101)=   "Column(12).Order=13"
      Splits(1)._ColumnProps(102)=   "Column(13).Width=2725"
      Splits(1)._ColumnProps(103)=   "Column(13).DividerColor=0"
      Splits(1)._ColumnProps(104)=   "Column(13)._WidthInPix=2646"
      Splits(1)._ColumnProps(105)=   "Column(13)._EditAlways=0"
      Splits(1)._ColumnProps(106)=   "Column(13)._ColStyle=74258"
      Splits(1)._ColumnProps(107)=   "Column(13).WrapText=1"
      Splits(1)._ColumnProps(108)=   "Column(13).Order=14"
      Splits(1)._ColumnProps(109)=   "Column(14).Width=2725"
      Splits(1)._ColumnProps(110)=   "Column(14).DividerColor=0"
      Splits(1)._ColumnProps(111)=   "Column(14)._WidthInPix=2646"
      Splits(1)._ColumnProps(112)=   "Column(14)._EditAlways=0"
      Splits(1)._ColumnProps(113)=   "Column(14)._ColStyle=66066"
      Splits(1)._ColumnProps(114)=   "Column(14).WrapText=1"
      Splits(1)._ColumnProps(115)=   "Column(14).Order=15"
      Splits(1)._ColumnProps(116)=   "Column(15).Width=2725"
      Splits(1)._ColumnProps(117)=   "Column(15).DividerColor=0"
      Splits(1)._ColumnProps(118)=   "Column(15)._WidthInPix=2646"
      Splits(1)._ColumnProps(119)=   "Column(15)._EditAlways=0"
      Splits(1)._ColumnProps(120)=   "Column(15).AllowSizing=0"
      Splits(1)._ColumnProps(121)=   "Column(15)._ColStyle=65808"
      Splits(1)._ColumnProps(122)=   "Column(15).Visible=0"
      Splits(1)._ColumnProps(123)=   "Column(15).WrapText=1"
      Splits(1)._ColumnProps(124)=   "Column(15).Order=16"
      Splits(1)._ColumnProps(125)=   "Column(16).Width=2725"
      Splits(1)._ColumnProps(126)=   "Column(16).DividerColor=0"
      Splits(1)._ColumnProps(127)=   "Column(16)._WidthInPix=2646"
      Splits(1)._ColumnProps(128)=   "Column(16)._EditAlways=0"
      Splits(1)._ColumnProps(129)=   "Column(16).AllowSizing=0"
      Splits(1)._ColumnProps(130)=   "Column(16)._ColStyle=65808"
      Splits(1)._ColumnProps(131)=   "Column(16).Visible=0"
      Splits(1)._ColumnProps(132)=   "Column(16).WrapText=1"
      Splits(1)._ColumnProps(133)=   "Column(16).Order=17"
      Splits(1)._ColumnProps(134)=   "Column(17).Width=2725"
      Splits(1)._ColumnProps(135)=   "Column(17).DividerColor=0"
      Splits(1)._ColumnProps(136)=   "Column(17)._WidthInPix=2646"
      Splits(1)._ColumnProps(137)=   "Column(17)._EditAlways=0"
      Splits(1)._ColumnProps(138)=   "Column(17)._ColStyle=65808"
      Splits(1)._ColumnProps(139)=   "Column(17).Visible=0"
      Splits(1)._ColumnProps(140)=   "Column(17).WrapText=1"
      Splits(1)._ColumnProps(141)=   "Column(17).Order=18"
      Splits(1)._ColumnProps(142)=   "Column(18).Width=2725"
      Splits(1)._ColumnProps(143)=   "Column(18).DividerColor=0"
      Splits(1)._ColumnProps(144)=   "Column(18)._WidthInPix=2646"
      Splits(1)._ColumnProps(145)=   "Column(18)._EditAlways=0"
      Splits(1)._ColumnProps(146)=   "Column(18)._ColStyle=65808"
      Splits(1)._ColumnProps(147)=   "Column(18).Visible=0"
      Splits(1)._ColumnProps(148)=   "Column(18).WrapText=1"
      Splits(1)._ColumnProps(149)=   "Column(18).Order=19"
      Splits(1)._ColumnProps(150)=   "Column(19).Width=2725"
      Splits(1)._ColumnProps(151)=   "Column(19).DividerColor=0"
      Splits(1)._ColumnProps(152)=   "Column(19)._WidthInPix=2646"
      Splits(1)._ColumnProps(153)=   "Column(19)._EditAlways=0"
      Splits(1)._ColumnProps(154)=   "Column(19)._ColStyle=65808"
      Splits(1)._ColumnProps(155)=   "Column(19).WrapText=1"
      Splits(1)._ColumnProps(156)=   "Column(19).Order=20"
      Splits(1)._ColumnProps(157)=   "Column(20).Width=2725"
      Splits(1)._ColumnProps(158)=   "Column(20).DividerColor=0"
      Splits(1)._ColumnProps(159)=   "Column(20)._WidthInPix=2646"
      Splits(1)._ColumnProps(160)=   "Column(20)._EditAlways=0"
      Splits(1)._ColumnProps(161)=   "Column(20)._ColStyle=65808"
      Splits(1)._ColumnProps(162)=   "Column(20).WrapText=1"
      Splits(1)._ColumnProps(163)=   "Column(20).Order=21"
      Splits(1)._ColumnProps(164)=   "Column(21).Width=2725"
      Splits(1)._ColumnProps(165)=   "Column(21).DividerColor=0"
      Splits(1)._ColumnProps(166)=   "Column(21)._WidthInPix=2646"
      Splits(1)._ColumnProps(167)=   "Column(21)._EditAlways=0"
      Splits(1)._ColumnProps(168)=   "Column(21)._ColStyle=65808"
      Splits(1)._ColumnProps(169)=   "Column(21).WrapText=1"
      Splits(1)._ColumnProps(170)=   "Column(21).Order=22"
      Splits(1)._ColumnProps(171)=   "Column(22).Width=2725"
      Splits(1)._ColumnProps(172)=   "Column(22).DividerColor=0"
      Splits(1)._ColumnProps(173)=   "Column(22)._WidthInPix=2646"
      Splits(1)._ColumnProps(174)=   "Column(22)._EditAlways=0"
      Splits(1)._ColumnProps(175)=   "Column(22)._ColStyle=65808"
      Splits(1)._ColumnProps(176)=   "Column(22).WrapText=1"
      Splits(1)._ColumnProps(177)=   "Column(22).Order=23"
      Splits(1)._ColumnProps(178)=   "Column(23).Width=2725"
      Splits(1)._ColumnProps(179)=   "Column(23).DividerColor=0"
      Splits(1)._ColumnProps(180)=   "Column(23)._WidthInPix=2646"
      Splits(1)._ColumnProps(181)=   "Column(23)._EditAlways=0"
      Splits(1)._ColumnProps(182)=   "Column(23)._ColStyle=65808"
      Splits(1)._ColumnProps(183)=   "Column(23).WrapText=1"
      Splits(1)._ColumnProps(184)=   "Column(23).Order=24"
      Splits(1)._ColumnProps(185)=   "Column(24).Width=2725"
      Splits(1)._ColumnProps(186)=   "Column(24).DividerColor=0"
      Splits(1)._ColumnProps(187)=   "Column(24)._WidthInPix=2646"
      Splits(1)._ColumnProps(188)=   "Column(24)._EditAlways=0"
      Splits(1)._ColumnProps(189)=   "Column(24)._ColStyle=65808"
      Splits(1)._ColumnProps(190)=   "Column(24).WrapText=1"
      Splits(1)._ColumnProps(191)=   "Column(24).Order=25"
      Splits(1)._ColumnProps(192)=   "Column(25).Width=2725"
      Splits(1)._ColumnProps(193)=   "Column(25).DividerColor=0"
      Splits(1)._ColumnProps(194)=   "Column(25)._WidthInPix=2646"
      Splits(1)._ColumnProps(195)=   "Column(25)._EditAlways=0"
      Splits(1)._ColumnProps(196)=   "Column(25)._ColStyle=65808"
      Splits(1)._ColumnProps(197)=   "Column(25).WrapText=1"
      Splits(1)._ColumnProps(198)=   "Column(25).Order=26"
      Splits.Count    =   2
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      DataMode        =   4
      DefColWidth     =   1,134
      HeadLines       =   2
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      DeadAreaBackColor=   12632256
      RowDividerColor =   12632256
      RowSubDividerColor=   12632256
      DirectionAfterEnter=   1
      MaxRows         =   250000
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=0,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=224,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
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
      _StyleDefs(24)  =   "Splits(0).Style:id=115,.parent=1,.valignment=2,.wraptext=-1"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=124,.parent=4"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=116,.parent=2"
      _StyleDefs(27)  =   "Splits(0).FooterStyle:id=117,.parent=3"
      _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=118,.parent=5"
      _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=120,.parent=6,.namedParent=38"
      _StyleDefs(30)  =   "Splits(0).EditorStyle:id=119,.parent=7"
      _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=121,.parent=8"
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=122,.parent=9"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=123,.parent=10"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=125,.parent=11"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=126,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=130,.parent=115,.alignment=2,.locked=-1"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=127,.parent=116,.alignment=2"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=128,.parent=117"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=129,.parent=119"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=134,.parent=115,.alignment=0,.valignment=2"
      _StyleDefs(41)  =   ":id=134,.locked=-1"
      _StyleDefs(42)  =   "Splits(0).Columns(1).HeadingStyle:id=131,.parent=116,.alignment=2"
      _StyleDefs(43)  =   "Splits(0).Columns(1).FooterStyle:id=132,.parent=117"
      _StyleDefs(44)  =   "Splits(0).Columns(1).EditorStyle:id=133,.parent=119"
      _StyleDefs(45)  =   "Splits(0).Columns(2).Style:id=138,.parent=115"
      _StyleDefs(46)  =   "Splits(0).Columns(2).HeadingStyle:id=135,.parent=116,.alignment=2"
      _StyleDefs(47)  =   "Splits(0).Columns(2).FooterStyle:id=136,.parent=117"
      _StyleDefs(48)  =   "Splits(0).Columns(2).EditorStyle:id=137,.parent=119"
      _StyleDefs(49)  =   "Splits(0).Columns(3).Style:id=142,.parent=115"
      _StyleDefs(50)  =   "Splits(0).Columns(3).HeadingStyle:id=139,.parent=116,.alignment=2"
      _StyleDefs(51)  =   "Splits(0).Columns(3).FooterStyle:id=140,.parent=117"
      _StyleDefs(52)  =   "Splits(0).Columns(3).EditorStyle:id=141,.parent=119"
      _StyleDefs(53)  =   "Splits(0).Columns(4).Style:id=146,.parent=115,.alignment=2,.valignment=2"
      _StyleDefs(54)  =   "Splits(0).Columns(4).HeadingStyle:id=143,.parent=116,.alignment=2"
      _StyleDefs(55)  =   "Splits(0).Columns(4).FooterStyle:id=144,.parent=117"
      _StyleDefs(56)  =   "Splits(0).Columns(4).EditorStyle:id=145,.parent=119"
      _StyleDefs(57)  =   "Splits(0).Columns(5).Style:id=150,.parent=115,.alignment=1"
      _StyleDefs(58)  =   "Splits(0).Columns(5).HeadingStyle:id=147,.parent=116,.alignment=2"
      _StyleDefs(59)  =   "Splits(0).Columns(5).FooterStyle:id=148,.parent=117"
      _StyleDefs(60)  =   "Splits(0).Columns(5).EditorStyle:id=149,.parent=119"
      _StyleDefs(61)  =   "Splits(0).Columns(6).Style:id=158,.parent=115,.alignment=1"
      _StyleDefs(62)  =   "Splits(0).Columns(6).HeadingStyle:id=155,.parent=116,.alignment=2"
      _StyleDefs(63)  =   "Splits(0).Columns(6).FooterStyle:id=156,.parent=117"
      _StyleDefs(64)  =   "Splits(0).Columns(6).EditorStyle:id=157,.parent=119"
      _StyleDefs(65)  =   "Splits(0).Columns(7).Style:id=162,.parent=115,.alignment=1"
      _StyleDefs(66)  =   "Splits(0).Columns(7).HeadingStyle:id=159,.parent=116,.alignment=2"
      _StyleDefs(67)  =   "Splits(0).Columns(7).FooterStyle:id=160,.parent=117"
      _StyleDefs(68)  =   "Splits(0).Columns(7).EditorStyle:id=161,.parent=119"
      _StyleDefs(69)  =   "Splits(0).Columns(8).Style:id=186,.parent=115,.alignment=1"
      _StyleDefs(70)  =   "Splits(0).Columns(8).HeadingStyle:id=183,.parent=116,.alignment=2"
      _StyleDefs(71)  =   "Splits(0).Columns(8).FooterStyle:id=184,.parent=117"
      _StyleDefs(72)  =   "Splits(0).Columns(8).EditorStyle:id=185,.parent=119"
      _StyleDefs(73)  =   "Splits(0).Columns(9).Style:id=190,.parent=115,.alignment=1"
      _StyleDefs(74)  =   "Splits(0).Columns(9).HeadingStyle:id=187,.parent=116,.alignment=2"
      _StyleDefs(75)  =   "Splits(0).Columns(9).FooterStyle:id=188,.parent=117"
      _StyleDefs(76)  =   "Splits(0).Columns(9).EditorStyle:id=189,.parent=119"
      _StyleDefs(77)  =   "Splits(0).Columns(10).Style:id=194,.parent=115,.alignment=1"
      _StyleDefs(78)  =   "Splits(0).Columns(10).HeadingStyle:id=191,.parent=116,.alignment=2"
      _StyleDefs(79)  =   "Splits(0).Columns(10).FooterStyle:id=192,.parent=117"
      _StyleDefs(80)  =   "Splits(0).Columns(10).EditorStyle:id=193,.parent=119"
      _StyleDefs(81)  =   "Splits(0).Columns(11).Style:id=198,.parent=115,.alignment=1"
      _StyleDefs(82)  =   "Splits(0).Columns(11).HeadingStyle:id=195,.parent=116,.alignment=2"
      _StyleDefs(83)  =   "Splits(0).Columns(11).FooterStyle:id=196,.parent=117"
      _StyleDefs(84)  =   "Splits(0).Columns(11).EditorStyle:id=197,.parent=119"
      _StyleDefs(85)  =   "Splits(0).Columns(12).Style:id=202,.parent=115,.alignment=1"
      _StyleDefs(86)  =   "Splits(0).Columns(12).HeadingStyle:id=199,.parent=116,.alignment=2"
      _StyleDefs(87)  =   "Splits(0).Columns(12).FooterStyle:id=200,.parent=117"
      _StyleDefs(88)  =   "Splits(0).Columns(12).EditorStyle:id=201,.parent=119"
      _StyleDefs(89)  =   "Splits(0).Columns(13).Style:id=226,.parent=115,.locked=-1"
      _StyleDefs(90)  =   "Splits(0).Columns(13).HeadingStyle:id=223,.parent=116"
      _StyleDefs(91)  =   "Splits(0).Columns(13).FooterStyle:id=224,.parent=117"
      _StyleDefs(92)  =   "Splits(0).Columns(13).EditorStyle:id=225,.parent=119"
      _StyleDefs(93)  =   "Splits(0).Columns(14).Style:id=206,.parent=115,.alignment=1"
      _StyleDefs(94)  =   "Splits(0).Columns(14).HeadingStyle:id=203,.parent=116,.alignment=2"
      _StyleDefs(95)  =   "Splits(0).Columns(14).FooterStyle:id=204,.parent=117"
      _StyleDefs(96)  =   "Splits(0).Columns(14).EditorStyle:id=205,.parent=119"
      _StyleDefs(97)  =   "Splits(0).Columns(15).Style:id=210,.parent=115"
      _StyleDefs(98)  =   "Splits(0).Columns(15).HeadingStyle:id=207,.parent=116"
      _StyleDefs(99)  =   "Splits(0).Columns(15).FooterStyle:id=208,.parent=117"
      _StyleDefs(100) =   "Splits(0).Columns(15).EditorStyle:id=209,.parent=119"
      _StyleDefs(101) =   "Splits(0).Columns(16).Style:id=218,.parent=115"
      _StyleDefs(102) =   "Splits(0).Columns(16).HeadingStyle:id=215,.parent=116"
      _StyleDefs(103) =   "Splits(0).Columns(16).FooterStyle:id=216,.parent=117"
      _StyleDefs(104) =   "Splits(0).Columns(16).EditorStyle:id=217,.parent=119"
      _StyleDefs(105) =   "Splits(0).Columns(17).Style:id=32,.parent=115"
      _StyleDefs(106) =   "Splits(0).Columns(17).HeadingStyle:id=29,.parent=116"
      _StyleDefs(107) =   "Splits(0).Columns(17).FooterStyle:id=30,.parent=117"
      _StyleDefs(108) =   "Splits(0).Columns(17).EditorStyle:id=31,.parent=119"
      _StyleDefs(109) =   "Splits(0).Columns(18).Style:id=66,.parent=115"
      _StyleDefs(110) =   "Splits(0).Columns(18).HeadingStyle:id=63,.parent=116"
      _StyleDefs(111) =   "Splits(0).Columns(18).FooterStyle:id=64,.parent=117"
      _StyleDefs(112) =   "Splits(0).Columns(18).EditorStyle:id=65,.parent=119"
      _StyleDefs(113) =   "Splits(0).Columns(19).Style:id=86,.parent=115"
      _StyleDefs(114) =   "Splits(0).Columns(19).HeadingStyle:id=83,.parent=116"
      _StyleDefs(115) =   "Splits(0).Columns(19).FooterStyle:id=84,.parent=117"
      _StyleDefs(116) =   "Splits(0).Columns(19).EditorStyle:id=85,.parent=119"
      _StyleDefs(117) =   "Splits(0).Columns(20).Style:id=154,.parent=115"
      _StyleDefs(118) =   "Splits(0).Columns(20).HeadingStyle:id=151,.parent=116"
      _StyleDefs(119) =   "Splits(0).Columns(20).FooterStyle:id=152,.parent=117"
      _StyleDefs(120) =   "Splits(0).Columns(20).EditorStyle:id=153,.parent=119"
      _StyleDefs(121) =   "Splits(0).Columns(21).Style:id=170,.parent=115"
      _StyleDefs(122) =   "Splits(0).Columns(21).HeadingStyle:id=167,.parent=116"
      _StyleDefs(123) =   "Splits(0).Columns(21).FooterStyle:id=168,.parent=117"
      _StyleDefs(124) =   "Splits(0).Columns(21).EditorStyle:id=169,.parent=119"
      _StyleDefs(125) =   "Splits(0).Columns(22).Style:id=178,.parent=115"
      _StyleDefs(126) =   "Splits(0).Columns(22).HeadingStyle:id=175,.parent=116"
      _StyleDefs(127) =   "Splits(0).Columns(22).FooterStyle:id=176,.parent=117"
      _StyleDefs(128) =   "Splits(0).Columns(22).EditorStyle:id=177,.parent=119"
      _StyleDefs(129) =   "Splits(0).Columns(23).Style:id=234,.parent=115"
      _StyleDefs(130) =   "Splits(0).Columns(23).HeadingStyle:id=231,.parent=116"
      _StyleDefs(131) =   "Splits(0).Columns(23).FooterStyle:id=232,.parent=117"
      _StyleDefs(132) =   "Splits(0).Columns(23).EditorStyle:id=233,.parent=119"
      _StyleDefs(133) =   "Splits(0).Columns(24).Style:id=242,.parent=115"
      _StyleDefs(134) =   "Splits(0).Columns(24).HeadingStyle:id=239,.parent=116"
      _StyleDefs(135) =   "Splits(0).Columns(24).FooterStyle:id=240,.parent=117"
      _StyleDefs(136) =   "Splits(0).Columns(24).EditorStyle:id=241,.parent=119"
      _StyleDefs(137) =   "Splits(0).Columns(25).Style:id=250,.parent=115"
      _StyleDefs(138) =   "Splits(0).Columns(25).HeadingStyle:id=247,.parent=116"
      _StyleDefs(139) =   "Splits(0).Columns(25).FooterStyle:id=248,.parent=117"
      _StyleDefs(140) =   "Splits(0).Columns(25).EditorStyle:id=249,.parent=119"
      _StyleDefs(141) =   "Splits(1).Style:id=43,.parent=1,.valignment=2,.wraptext=-1"
      _StyleDefs(142) =   "Splits(1).CaptionStyle:id=52,.parent=4"
      _StyleDefs(143) =   "Splits(1).HeadingStyle:id=44,.parent=2"
      _StyleDefs(144) =   "Splits(1).FooterStyle:id=45,.parent=3"
      _StyleDefs(145) =   "Splits(1).InactiveStyle:id=46,.parent=5"
      _StyleDefs(146) =   "Splits(1).SelectedStyle:id=48,.parent=6"
      _StyleDefs(147) =   "Splits(1).EditorStyle:id=47,.parent=7"
      _StyleDefs(148) =   "Splits(1).HighlightRowStyle:id=49,.parent=8"
      _StyleDefs(149) =   "Splits(1).EvenRowStyle:id=50,.parent=9"
      _StyleDefs(150) =   "Splits(1).OddRowStyle:id=51,.parent=10"
      _StyleDefs(151) =   "Splits(1).RecordSelectorStyle:id=53,.parent=11"
      _StyleDefs(152) =   "Splits(1).FilterBarStyle:id=54,.parent=12"
      _StyleDefs(153) =   "Splits(1).Columns(0).Style:id=114,.parent=43,.alignment=2"
      _StyleDefs(154) =   "Splits(1).Columns(0).HeadingStyle:id=111,.parent=44,.alignment=2"
      _StyleDefs(155) =   "Splits(1).Columns(0).FooterStyle:id=112,.parent=45"
      _StyleDefs(156) =   "Splits(1).Columns(0).EditorStyle:id=113,.parent=47"
      _StyleDefs(157) =   "Splits(1).Columns(1).Style:id=28,.parent=43,.alignment=0,.valignment=2"
      _StyleDefs(158) =   "Splits(1).Columns(1).HeadingStyle:id=25,.parent=44,.alignment=2"
      _StyleDefs(159) =   "Splits(1).Columns(1).FooterStyle:id=26,.parent=45"
      _StyleDefs(160) =   "Splits(1).Columns(1).EditorStyle:id=27,.parent=47"
      _StyleDefs(161) =   "Splits(1).Columns(2).Style:id=78,.parent=43,.locked=-1"
      _StyleDefs(162) =   "Splits(1).Columns(2).HeadingStyle:id=75,.parent=44,.alignment=2"
      _StyleDefs(163) =   "Splits(1).Columns(2).FooterStyle:id=76,.parent=45"
      _StyleDefs(164) =   "Splits(1).Columns(2).EditorStyle:id=77,.parent=47"
      _StyleDefs(165) =   "Splits(1).Columns(3).Style:id=106,.parent=43,.locked=-1"
      _StyleDefs(166) =   "Splits(1).Columns(3).HeadingStyle:id=103,.parent=44,.alignment=2"
      _StyleDefs(167) =   "Splits(1).Columns(3).FooterStyle:id=104,.parent=45"
      _StyleDefs(168) =   "Splits(1).Columns(3).EditorStyle:id=105,.parent=47"
      _StyleDefs(169) =   "Splits(1).Columns(4).Style:id=70,.parent=43,.alignment=2,.valignment=2"
      _StyleDefs(170) =   "Splits(1).Columns(4).HeadingStyle:id=67,.parent=44,.alignment=2"
      _StyleDefs(171) =   "Splits(1).Columns(4).FooterStyle:id=68,.parent=45"
      _StyleDefs(172) =   "Splits(1).Columns(4).EditorStyle:id=69,.parent=47"
      _StyleDefs(173) =   "Splits(1).Columns(5).Style:id=24,.parent=43,.alignment=1,.locked=-1"
      _StyleDefs(174) =   "Splits(1).Columns(5).HeadingStyle:id=21,.parent=44,.alignment=2"
      _StyleDefs(175) =   "Splits(1).Columns(5).FooterStyle:id=22,.parent=45"
      _StyleDefs(176) =   "Splits(1).Columns(5).EditorStyle:id=23,.parent=47"
      _StyleDefs(177) =   "Splits(1).Columns(6).Style:id=58,.parent=43,.alignment=1,.locked=-1"
      _StyleDefs(178) =   "Splits(1).Columns(6).HeadingStyle:id=55,.parent=44,.alignment=2"
      _StyleDefs(179) =   "Splits(1).Columns(6).FooterStyle:id=56,.parent=45"
      _StyleDefs(180) =   "Splits(1).Columns(6).EditorStyle:id=57,.parent=47"
      _StyleDefs(181) =   "Splits(1).Columns(7).Style:id=74,.parent=43,.alignment=1,.locked=-1"
      _StyleDefs(182) =   "Splits(1).Columns(7).HeadingStyle:id=71,.parent=44,.alignment=2"
      _StyleDefs(183) =   "Splits(1).Columns(7).FooterStyle:id=72,.parent=45"
      _StyleDefs(184) =   "Splits(1).Columns(7).EditorStyle:id=73,.parent=47"
      _StyleDefs(185) =   "Splits(1).Columns(8).Style:id=94,.parent=43,.alignment=1"
      _StyleDefs(186) =   "Splits(1).Columns(8).HeadingStyle:id=91,.parent=44,.alignment=2"
      _StyleDefs(187) =   "Splits(1).Columns(8).FooterStyle:id=92,.parent=45"
      _StyleDefs(188) =   "Splits(1).Columns(8).EditorStyle:id=93,.parent=47"
      _StyleDefs(189) =   "Splits(1).Columns(9).Style:id=98,.parent=43,.alignment=1"
      _StyleDefs(190) =   "Splits(1).Columns(9).HeadingStyle:id=95,.parent=44,.alignment=2"
      _StyleDefs(191) =   "Splits(1).Columns(9).FooterStyle:id=96,.parent=45"
      _StyleDefs(192) =   "Splits(1).Columns(9).EditorStyle:id=97,.parent=47"
      _StyleDefs(193) =   "Splits(1).Columns(10).Style:id=102,.parent=43,.alignment=1,.locked=-1"
      _StyleDefs(194) =   "Splits(1).Columns(10).HeadingStyle:id=99,.parent=44,.alignment=2"
      _StyleDefs(195) =   "Splits(1).Columns(10).FooterStyle:id=100,.parent=45"
      _StyleDefs(196) =   "Splits(1).Columns(10).EditorStyle:id=101,.parent=47"
      _StyleDefs(197) =   "Splits(1).Columns(11).Style:id=110,.parent=43,.alignment=1"
      _StyleDefs(198) =   "Splits(1).Columns(11).HeadingStyle:id=107,.parent=44,.alignment=2"
      _StyleDefs(199) =   "Splits(1).Columns(11).FooterStyle:id=108,.parent=45"
      _StyleDefs(200) =   "Splits(1).Columns(11).EditorStyle:id=109,.parent=47"
      _StyleDefs(201) =   "Splits(1).Columns(12).Style:id=16,.parent=43,.alignment=1"
      _StyleDefs(202) =   "Splits(1).Columns(12).HeadingStyle:id=13,.parent=44,.alignment=2"
      _StyleDefs(203) =   "Splits(1).Columns(12).FooterStyle:id=14,.parent=45"
      _StyleDefs(204) =   "Splits(1).Columns(12).EditorStyle:id=15,.parent=47"
      _StyleDefs(205) =   "Splits(1).Columns(13).Style:id=230,.parent=43,.alignment=1,.locked=-1"
      _StyleDefs(206) =   "Splits(1).Columns(13).HeadingStyle:id=227,.parent=44,.alignment=2"
      _StyleDefs(207) =   "Splits(1).Columns(13).FooterStyle:id=228,.parent=45"
      _StyleDefs(208) =   "Splits(1).Columns(13).EditorStyle:id=229,.parent=47"
      _StyleDefs(209) =   "Splits(1).Columns(14).Style:id=20,.parent=43,.alignment=1"
      _StyleDefs(210) =   "Splits(1).Columns(14).HeadingStyle:id=17,.parent=44,.alignment=2"
      _StyleDefs(211) =   "Splits(1).Columns(14).FooterStyle:id=18,.parent=45"
      _StyleDefs(212) =   "Splits(1).Columns(14).EditorStyle:id=19,.parent=47"
      _StyleDefs(213) =   "Splits(1).Columns(15).Style:id=214,.parent=43"
      _StyleDefs(214) =   "Splits(1).Columns(15).HeadingStyle:id=211,.parent=44"
      _StyleDefs(215) =   "Splits(1).Columns(15).FooterStyle:id=212,.parent=45"
      _StyleDefs(216) =   "Splits(1).Columns(15).EditorStyle:id=213,.parent=47"
      _StyleDefs(217) =   "Splits(1).Columns(16).Style:id=222,.parent=43"
      _StyleDefs(218) =   "Splits(1).Columns(16).HeadingStyle:id=219,.parent=44"
      _StyleDefs(219) =   "Splits(1).Columns(16).FooterStyle:id=220,.parent=45"
      _StyleDefs(220) =   "Splits(1).Columns(16).EditorStyle:id=221,.parent=47"
      _StyleDefs(221) =   "Splits(1).Columns(17).Style:id=62,.parent=43"
      _StyleDefs(222) =   "Splits(1).Columns(17).HeadingStyle:id=59,.parent=44"
      _StyleDefs(223) =   "Splits(1).Columns(17).FooterStyle:id=60,.parent=45"
      _StyleDefs(224) =   "Splits(1).Columns(17).EditorStyle:id=61,.parent=47"
      _StyleDefs(225) =   "Splits(1).Columns(18).Style:id=82,.parent=43"
      _StyleDefs(226) =   "Splits(1).Columns(18).HeadingStyle:id=79,.parent=44"
      _StyleDefs(227) =   "Splits(1).Columns(18).FooterStyle:id=80,.parent=45"
      _StyleDefs(228) =   "Splits(1).Columns(18).EditorStyle:id=81,.parent=47"
      _StyleDefs(229) =   "Splits(1).Columns(19).Style:id=90,.parent=43"
      _StyleDefs(230) =   "Splits(1).Columns(19).HeadingStyle:id=87,.parent=44"
      _StyleDefs(231) =   "Splits(1).Columns(19).FooterStyle:id=88,.parent=45"
      _StyleDefs(232) =   "Splits(1).Columns(19).EditorStyle:id=89,.parent=47"
      _StyleDefs(233) =   "Splits(1).Columns(20).Style:id=166,.parent=43"
      _StyleDefs(234) =   "Splits(1).Columns(20).HeadingStyle:id=163,.parent=44"
      _StyleDefs(235) =   "Splits(1).Columns(20).FooterStyle:id=164,.parent=45"
      _StyleDefs(236) =   "Splits(1).Columns(20).EditorStyle:id=165,.parent=47"
      _StyleDefs(237) =   "Splits(1).Columns(21).Style:id=174,.parent=43"
      _StyleDefs(238) =   "Splits(1).Columns(21).HeadingStyle:id=171,.parent=44"
      _StyleDefs(239) =   "Splits(1).Columns(21).FooterStyle:id=172,.parent=45"
      _StyleDefs(240) =   "Splits(1).Columns(21).EditorStyle:id=173,.parent=47"
      _StyleDefs(241) =   "Splits(1).Columns(22).Style:id=182,.parent=43"
      _StyleDefs(242) =   "Splits(1).Columns(22).HeadingStyle:id=179,.parent=44"
      _StyleDefs(243) =   "Splits(1).Columns(22).FooterStyle:id=180,.parent=45"
      _StyleDefs(244) =   "Splits(1).Columns(22).EditorStyle:id=181,.parent=47"
      _StyleDefs(245) =   "Splits(1).Columns(23).Style:id=238,.parent=43"
      _StyleDefs(246) =   "Splits(1).Columns(23).HeadingStyle:id=235,.parent=44"
      _StyleDefs(247) =   "Splits(1).Columns(23).FooterStyle:id=236,.parent=45"
      _StyleDefs(248) =   "Splits(1).Columns(23).EditorStyle:id=237,.parent=47"
      _StyleDefs(249) =   "Splits(1).Columns(24).Style:id=246,.parent=43"
      _StyleDefs(250) =   "Splits(1).Columns(24).HeadingStyle:id=243,.parent=44"
      _StyleDefs(251) =   "Splits(1).Columns(24).FooterStyle:id=244,.parent=45"
      _StyleDefs(252) =   "Splits(1).Columns(24).EditorStyle:id=245,.parent=47"
      _StyleDefs(253) =   "Splits(1).Columns(25).Style:id=254,.parent=43"
      _StyleDefs(254) =   "Splits(1).Columns(25).HeadingStyle:id=251,.parent=44"
      _StyleDefs(255) =   "Splits(1).Columns(25).FooterStyle:id=252,.parent=45"
      _StyleDefs(256) =   "Splits(1).Columns(25).EditorStyle:id=253,.parent=47"
      _StyleDefs(257) =   "Named:id=33:Normal"
      _StyleDefs(258) =   ":id=33,.parent=0"
      _StyleDefs(259) =   "Named:id=34:Heading"
      _StyleDefs(260) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(261) =   ":id=34,.wraptext=-1"
      _StyleDefs(262) =   "Named:id=35:Footing"
      _StyleDefs(263) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(264) =   "Named:id=36:Selected"
      _StyleDefs(265) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(266) =   "Named:id=37:Caption"
      _StyleDefs(267) =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(268) =   "Named:id=38:HighlightRow"
      _StyleDefs(269) =   ":id=38,.parent=33,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
      _StyleDefs(270) =   "Named:id=39:EvenRow"
      _StyleDefs(271) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(272) =   "Named:id=40:OddRow"
      _StyleDefs(273) =   ":id=40,.parent=33"
      _StyleDefs(274) =   "Named:id=41:RecordSelector"
      _StyleDefs(275) =   ":id=41,.parent=34"
      _StyleDefs(276) =   "Named:id=42:FilterBar"
      _StyleDefs(277) =   ":id=42,.parent=33"
   End
   Begin Threed.SSPanel PanBotoes 
      Height          =   705
      Left            =   90
      TabIndex        =   10
      Top             =   3840
      Width           =   8760
      _Version        =   65536
      _ExtentX        =   15452
      _ExtentY        =   1244
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
      Begin VB.CommandButton CmdTitulo 
         Caption         =   "&Título"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1443
         TabIndex        =   44
         Top             =   405
         Width           =   1350
      End
      Begin VB.Frame Frame2 
         Caption         =   "Total a Pagar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   690
         Left            =   7140
         TabIndex        =   48
         Top             =   30
         Width           =   1635
         Begin VB.Label LblTotal 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
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
            Left            =   90
            TabIndex        =   49
            Top             =   270
            Width           =   1440
         End
      End
      Begin VB.CommandButton CmdRecibo 
         BackColor       =   &H00000000&
         Caption         =   "Recibo &Único"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3585
         TabIndex        =   41
         Top             =   45
         Width           =   1665
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
         Height          =   300
         Left            =   2856
         TabIndex        =   45
         Top             =   405
         Width           =   1350
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
         Height          =   300
         Left            =   4269
         TabIndex        =   46
         Top             =   405
         Width           =   1350
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
         Height          =   300
         Left            =   5685
         TabIndex        =   47
         Top             =   405
         Width           =   1350
      End
      Begin VB.CommandButton CmdGravar 
         BackColor       =   &H00000000&
         Caption         =   "&Gravar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1815
         TabIndex        =   40
         Top             =   45
         Width           =   1665
      End
      Begin VB.CommandButton CmdContrato 
         Caption         =   "&Contrato"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   30
         TabIndex        =   43
         Top             =   405
         Width           =   1350
      End
      Begin VB.CommandButton CmdImprimeReciboM 
         BackColor       =   &H00000000&
         Caption         =   "Recibos &Múltiplos"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5370
         TabIndex        =   42
         Top             =   45
         Width           =   1665
      End
      Begin VB.CommandButton CmdDesmarcar 
         Caption         =   "Desmarcar Todos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   30
         TabIndex        =   39
         Top             =   45
         Width           =   1665
      End
   End
   Begin TabDlg.SSTab TabFinanceiro 
      Height          =   1965
      Left            =   15
      TabIndex        =   12
      Top             =   30
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   3466
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Dados Gerais"
      TabPicture(0)   =   "TelaPagtoMultiplos.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "PanFundo"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Dados Bancários"
      TabPicture(1)   =   "TelaPagtoMultiplos.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "LblCCorrente"
      Tab(1).Control(1)=   "DatContaCorrente"
      Tab(1).Control(2)=   "CboCCorrente"
      Tab(1).Control(3)=   "Frame1"
      Tab(1).ControlCount=   4
      Begin Threed.SSPanel PanFundo 
         Height          =   1560
         Left            =   45
         TabIndex        =   23
         Top             =   360
         Width           =   8790
         _Version        =   65536
         _ExtentX        =   15505
         _ExtentY        =   2752
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
         Begin VB.Timer Timer1 
            Left            =   2880
            Top             =   480
         End
         Begin Threed.SSPanel SSPanel2 
            Height          =   315
            Left            =   5160
            TabIndex        =   51
            Top             =   840
            Width           =   2520
            _Version        =   65536
            _ExtentX        =   4445
            _ExtentY        =   556
            _StockProps     =   15
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Begin VB.CheckBox ChkResiduo 
               Caption         =   "Gerar Parcela de Resíduo?"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   90
               TabIndex        =   5
               Top             =   60
               Value           =   1  'Checked
               Width           =   2280
            End
         End
         Begin MSComCtl2.DTPicker DtpPagamento 
            Height          =   300
            Left            =   1125
            TabIndex        =   3
            Top             =   840
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "dd/MM/yy"
            Format          =   63111171
            CurrentDate     =   37180
         End
         Begin MSComCtl2.DTPicker DtpBasePagto 
            Height          =   300
            Left            =   3645
            TabIndex        =   4
            Top             =   840
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "dd/MM/yy"
            Format          =   63111171
            CurrentDate     =   37180
         End
         Begin VB.Frame FraCorrecao 
            Caption         =   "Correção"
            ForeColor       =   &H00404000&
            Height          =   450
            Left            =   75
            TabIndex        =   38
            Top             =   1095
            Width           =   5955
            Begin Threed.SSPanel SSPanel1 
               Height          =   270
               Left            =   2715
               TabIndex        =   50
               Top             =   135
               Width           =   3180
               _Version        =   65536
               _ExtentX        =   5609
               _ExtentY        =   476
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
               Begin VB.OptionButton OptIndexador1 
                  Caption         =   "Indexador 1"
                  Height          =   195
                  Left            =   270
                  TabIndex        =   7
                  Top             =   45
                  Value           =   -1  'True
                  Width           =   1200
               End
               Begin VB.OptionButton OptIndexador2 
                  Caption         =   "Indexador 2"
                  Height          =   195
                  Left            =   1695
                  TabIndex        =   8
                  Top             =   45
                  Width           =   1140
               End
            End
            Begin VB.CheckBox ChkCorrecao 
               Caption         =   "Pagto. c/ Parcela Corrigida"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   135
               TabIndex        =   6
               Top             =   165
               Value           =   1  'Checked
               Width           =   2505
            End
         End
         Begin MSMask.MaskEdBox MskEmpreendimento 
            Height          =   300
            Left            =   1545
            TabIndex        =   0
            Top             =   15
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   4
            Mask            =   "AAAA"
            PromptChar      =   " "
         End
         Begin MSDataListLib.DataCombo CboEmpreendimento 
            Bindings        =   "TelaPagtoMultiplos.frx":0038
            Height          =   315
            Left            =   2145
            TabIndex        =   1
            Top             =   0
            Width           =   3945
            _ExtentX        =   6959
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            Style           =   2
            ListField       =   "empd_tx_nome"
            BoundColumn     =   "empd_cd_empreendimento"
            Text            =   "CboEmpreendimento"
         End
         Begin MSDataListLib.DataCombo CboImovel 
            Bindings        =   "TelaPagtoMultiplos.frx":0058
            Height          =   315
            Left            =   6870
            TabIndex        =   2
            Top             =   0
            Width           =   1860
            _ExtentX        =   3281
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            MatchEntry      =   -1  'True
            Style           =   2
            ListField       =   "imov_cd_imovel"
            BoundColumn     =   "imov_cd_imovel"
            Text            =   ""
         End
         Begin MSAdodcLib.Adodc DatEmpreendimento 
            Height          =   330
            Left            =   2505
            Top             =   75
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
         Begin MSAdodcLib.Adodc DatImovel 
            Height          =   330
            Left            =   6825
            Top             =   75
            Visible         =   0   'False
            Width           =   1665
            _ExtentX        =   2937
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
            Caption         =   "DatImovel"
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
         Begin Threed.SSCommand CmdPesquisar 
            Height          =   300
            Left            =   7755
            TabIndex        =   9
            Top             =   1230
            Width           =   990
            _Version        =   65536
            _ExtentX        =   1746
            _ExtentY        =   529
            _StockProps     =   78
            Caption         =   "&Pesquisar"
            Enabled         =   0   'False
         End
         Begin VB.Label LblDtContrato2 
            Caption         =   "01/11/02"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   7695
            TabIndex        =   26
            Top             =   585
            Width           =   1095
         End
         Begin VB.Label LblDtVenda2 
            Caption         =   "01/11/02"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   240
            Left            =   4800
            TabIndex        =   28
            Top             =   585
            Width           =   1065
         End
         Begin VB.Label LblContrato 
            Caption         =   "Contrato:"
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
            Left            =   7590
            TabIndex        =   37
            Top             =   330
            Width           =   795
         End
         Begin VB.Label LblContrato2 
            Caption         =   "01"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   8430
            TabIndex        =   36
            Top             =   345
            Width           =   255
         End
         Begin VB.Label LblCliente 
            Alignment       =   1  'Right Justify
            Caption         =   "Cliente:"
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
            Left            =   75
            TabIndex        =   35
            Top             =   330
            Width           =   645
         End
         Begin VB.Label LblRazaoSocial 
            Caption         =   "ELANE SOUZA LAGO"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   225
            Left            =   780
            TabIndex        =   34
            Top             =   345
            Width           =   6615
         End
         Begin VB.Label LblImovel 
            Alignment       =   1  'Right Justify
            Caption         =   "Imóvel:"
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
            Left            =   6015
            TabIndex        =   33
            Top             =   75
            Width           =   810
         End
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
            Height          =   240
            Left            =   0
            TabIndex        =   32
            Top             =   60
            Width           =   1530
         End
         Begin VB.Label LblTipoImovel 
            Alignment       =   1  'Right Justify
            Caption         =   "Tipo do Imóvel:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   30
            TabIndex        =   31
            Top             =   585
            Width           =   1365
         End
         Begin VB.Label LblDtVenda 
            Alignment       =   1  'Right Justify
            Caption         =   "Data Venda:"
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
            Left            =   3690
            TabIndex        =   30
            Top             =   585
            Width           =   1095
         End
         Begin VB.Label LblTipoImovel2 
            Caption         =   "Tipo Imovel"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   1395
            TabIndex        =   29
            Top             =   585
            Width           =   2145
         End
         Begin VB.Label LblDtContrato 
            Alignment       =   1  'Right Justify
            Caption         =   "Data Base Contrato:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   5925
            TabIndex        =   27
            Top             =   585
            Width           =   1755
         End
         Begin VB.Label LblDtPagamento 
            Caption         =   "Pagamento:"
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
            Left            =   75
            TabIndex        =   25
            Top             =   870
            Width           =   1095
         End
         Begin VB.Label LblDtRepasse 
            Alignment       =   1  'Right Justify
            Caption         =   "Base Pagto:"
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
            Left            =   2295
            TabIndex        =   24
            Top             =   870
            Width           =   1320
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Dados do Pagamento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1035
         Left            =   -74880
         TabIndex        =   13
         Top             =   360
         Width           =   8655
         Begin VB.CheckBox ChkPreDatado 
            Caption         =   "Pré-Datado"
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
            Height          =   315
            Left            =   2835
            TabIndex        =   16
            Top             =   240
            Width           =   1335
         End
         Begin VB.ComboBox CboLocalPagto 
            Height          =   315
            ItemData        =   "TelaPagtoMultiplos.frx":0070
            Left            =   6510
            List            =   "TelaPagtoMultiplos.frx":0080
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   225
            Width           =   2040
         End
         Begin VB.TextBox TxtDocPagto 
            Height          =   315
            Left            =   2250
            MaxLength       =   40
            TabIndex        =   20
            Top             =   630
            Width           =   6300
         End
         Begin MSComCtl2.DTPicker DtpDeposito 
            Height          =   300
            Left            =   990
            TabIndex        =   14
            Top             =   255
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   529
            _Version        =   393216
            DateIsNull      =   -1  'True
            Format          =   63111169
            CurrentDate     =   37180
         End
         Begin VB.Label LblDtDeposito 
            Caption         =   "Depósito:"
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
            Left            =   150
            TabIndex        =   19
            Top             =   285
            Width           =   915
         End
         Begin VB.Label LblLocalPgto 
            Caption         =   "Forma de Pagamento:"
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
            Left            =   4605
            TabIndex        =   17
            Top             =   270
            Width           =   1875
         End
         Begin VB.Label LblDocPagto 
            Caption         =   "Documento Pagamento:"
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
            Left            =   165
            TabIndex        =   15
            Top             =   660
            Width           =   2115
         End
      End
      Begin MSDataListLib.DataCombo CboCCorrente 
         Bindings        =   "TelaPagtoMultiplos.frx":00B2
         Height          =   330
         Left            =   -73350
         TabIndex        =   21
         Top             =   1470
         Width           =   7035
         _ExtentX        =   12409
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
         Left            =   -73395
         Top             =   1425
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
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
      Begin VB.Label LblCCorrente 
         Alignment       =   1  'Right Justify
         Caption         =   "Conta Corrente:"
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
         Height          =   270
         Left            =   -74775
         TabIndex        =   22
         Top             =   1500
         Width           =   1365
      End
   End
End
Attribute VB_Name = "TelaFinPagtoMultiplos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CarregouGrid As Boolean 'Indica se o grid já foi carregado ou não
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

Dim FRM_XInserir As String        'Flag para identificar se entrou na rotina para inserir ou alterar ('I'Inserir, 'A'Alterar, 'C'Consultar)
Dim FRM_XOrigem As String         'Identifica a origem (que formulario) chamou o formulário ativo
Dim FRM_SQLAtivo As String        'Atualizar o SQL Ativo após modificações de filtro e ordem

Dim Vetor1 As New XArray
Dim XFF_SALDODEVEDOR As Double  'Armazena o saldo devedor do cliente a medida que vai sendo desmarcado no grid

Dim XFT_SQL As String       'Prepara Select
Dim XFT_SQLBASE As String

Dim XFO_PAGAMENTO As New ADODB.Recordset


Sub subTelaValoresGlobais(GuardaOUPegaValores As String)
    Dim Count As Integer, Count2 As Integer
    If GuardaOUPegaValores = "G" Then 'Guardando variáveis de definição da tela
        FRM_ContFil = ContFil
        FRM_Ordem = Ordem
        FRM_Campos_Obrig = Campos_Obrig
        'FRM_TipoVar = TipoVar
        FRM_TipoVar = XGT_CAMPOFILTRO
        FRM_FiltroOrig = FiltroOrig
        FRM_FiltroAtual = FiltroAtual
        FRM_filtra_especial = filtra_especial
        FRM_Tem_Filtro = Tem_Filtro
        FRM_ConTelaOrdem = ConTelaOrdem
        FRM_Quantid = Quantid
        FRM_Filtrou = Filtrou
        FRM_XInserir = XInserir
        FRM_XOrigem = XOrigem
        FRM_SQLAtivo = SQLAtivo
        
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
        'TipoVar = FRM_TipoVar
        XGT_CAMPOFILTRO = FRM_TipoVar
        FiltroOrig = FRM_FiltroOrig
        FiltroAtual = FRM_FiltroAtual
        filtra_especial = FRM_filtra_especial
        Tem_Filtro = FRM_Tem_Filtro
        ConTelaOrdem = FRM_ConTelaOrdem
        Quantid = FRM_Quantid
        Filtrou = FRM_Filtrou
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


'Retorna o valor calculado do residuo
Function funCalculaResiduo(XLF_TOTALTITULO As Double, XLF_VLCORRIGIDO As Double) As Double
    'Dim XLF_TOTALTITULO As Double  'Valor total a pagar do título
    Dim XLF_PERCENT As Double    'Percentual encontrado da divisão entre valor pago e valor a pagar
    Dim XLF_MULTA
    Dim XLF_JUROS
    Dim XLF_DESAGIO
    'Dim XLF_VLCORRIGIDO As Double
   
    XLF_TOTALTITULO = (((XLF_TOTALTITULO * 100) \ 1) / 100) 'Resolve problemas de aproximação
    If CDbl(FunNuloVal(TDBGrid1.Columns("Valor a Pagar"))) < XLF_TOTALTITULO Then
        XLF_PERCENT = CDbl(FunNuloVal(TDBGrid1.Columns("Valor a Pagar"))) / XLF_TOTALTITULO
        XLF_MULTA = CDbl(FunNuloVal(TDBGrid1.Columns("Multa"))) * XLF_PERCENT
        XLF_JUROS = CDbl(FunNuloVal(TDBGrid1.Columns("Juros"))) * XLF_PERCENT
        XLF_DESAGIO = CDbl(FunNuloVal(TDBGrid1.Columns("Deságio"))) * XLF_PERCENT
        
        funCalculaResiduo = XLF_VLCORRIGIDO - XLF_DESAGIO - CDbl(FunNuloVal(TDBGrid1.Columns("Desconto"))) - CDbl(FunNuloVal(TDBGrid1.Columns("Valor a Pagar"))) + XLF_MULTA + XLF_JUROS + CDbl(FunNuloVal(TDBGrid1.Columns("Outros(+)")))
    ElseIf ChkCorrecao.Value = 0 Then
        funCalculaResiduo = XLF_VLCORRIGIDO - CDbl(TDBGrid1.Columns("Valor Histórico"))
    End If
End Function
Sub subCalculaSeguro(XLI_X As Variant)
'    Dim XLF_TAXASEGURO As Double    'Taxa de seguro
'    Dim XLT_TIPOSEGURO As String    'Tipo do seguro do contrato
'    Dim XLT_SQL As String           'Prepara select
'    Dim XLO_TITULO As Object        'Recordset para tabela titulos
'    Dim XLB_CALCSEGUROS As Boolean  'indica se o seguros deverá ou não ser calculado
'    Dim XLB_ACHEI As Boolean        'Indica se já localizou o registro no grid
'    Dim XLF_VLSEGURO As Double      'Valor calculado do seguro
'    Dim XLF_VALOR As Double         'Valor corrigido com seguro
'    Dim XLF_JUROS As Double         'Valor corrigido com seguro
'    Dim XLF_MULTA As Double         'Valor corrigido com seguro
'    Dim XLI_Y As Integer            'Contador
'
'    'Entende-se que nesse grid todos os registros serão pagos, portanto a principio os
'    'registros serão apresentados na tela sem cálculo do seguros. Este será calculado quando
'    'um registro for desmarcado
'    DatImovel.Recordset.bookmark = CboImovel.SelectedItem
'    XLF_TAXASEGURO = DatImovel.Recordset.Fields("cont_vl_TaxaSegVida")
'    XLT_TIPOSEGURO = DatImovel.Recordset.Fields("cont_nr_TipoSeguro")
'
'    If XLF_TAXASEGURO <> 0 Then 'DatTitulos.Recordset.Fields!natureza <> "FINANCIAMENTO" Then
'        'Verificar se o seguro já foi pago no mês
'        XLT_SQL = "SELECT titu_vl_seguro FROM Titulos WHERE empr_cd_empresa = " & PCodEmpresa & "" & _
'                    " AND empd_cd_Empreendimento= '" & MskEmpreendimento.Text & "'" & _
'                    " AND imov_cd_Imovel='" & CboImovel.BoundText & "'" & _
'                    " AND cont_cd_Contrato='" & LblContrato2.Caption & "'" & _
'                    " AND month(titu_dt_pagamento)=" & Month(DtpPagamento.Value) & "" & _
'                    " AND titu_vl_Seguro<>0 "
'        SubQOpenRecordset XLO_TITULO, XLT_SQL, Estatico
'        If XLO_TITULO.EOF Then 'Se não encontrou registros que atendesse a condição o seguro deve ser calculado
'            XLB_CALCSEGUROS = True
'        Else
'            XLB_CALCSEGUROS = False
'        End If
'        XLO_TITULO.Close
'
'        If XLB_CALCSEGUROS = True Then
'            If XLT_TIPOSEGURO = "2" Then 'Cálculo do seguro pela parcela
'                Vetor1(TDBGrid1.bookmark, 8) = Format(CDbl(TDBGrid1.Columns("Valor Histórico").Value) * XLF_TAXASEGURO, "standard")
'            ElseIf XLT_TIPOSEGURO = "1" Then 'Cálculo do seguro pela saldo devedor
'                XLF_VLSEGURO = Format(XFF_SALDODEVEDOR * XLF_TAXASEGURO, "standard")
'
'                'Abrir os registro (mensais) não pagos por ordem de dt. de criação
'                XLT_SQL = "SELECT * FROM consCAPTitulos WHERE empr_cd_empresa = " & PCodEmpresa & "" & _
'                              " AND empd_cd_Empreendimento= '" & Trim(MskEmpreendimento.Text) & "'" & _
'                              " AND imov_cd_Imovel='" & Trim(CboImovel.Text) & "'" & _
'                              " AND cont_cd_Contrato='" & Trim(LblContrato2.Caption) & "'" & _
'                              " AND titu_dt_Pagamento is null AND tipl_tx_Descricao = 'MENSAL'" & _
'                              " AND Natureza <> 'FINANCIAMENTO' ORDER BY titu_dt_Criacao"
'                SubQOpenRecordset XLO_TITULO, XLT_SQL, Estatico
'                'Localizar no grid e verificar se está selecionado
'                If Not XLO_TITULO.EOF Then
'                    While Not XLO_TITULO.EOF And Not XLB_ACHEI
'                        For XLI_Y = 0 To Vetor1.UpperBound(1)
'                            If Vetor1(XLI_Y, 1) = XLO_TITULO!Titulo And Vetor1(XLI_Y, 4) = "S" Then
'                                Vetor1(XLI_Y, 8) = Format(XFF_SALDODEVEDOR * XLF_TAXASEGURO, "standard")
'                                XLB_ACHEI = True
'                            Else
'                                Vetor1(XLI_Y, 8) = ""
'                            End If
'
'                            '//////////////RECALCULA VALORES////////////////
'                            If ChkCorrecao.Value = 0 Then
'                                XLF_VALOR = CDbl(Vetor1(XLI_Y, 5)) + CDbl(FunNuloVal(Vetor1(XLI_Y, 8)))
'                            Else
'                                XLF_VALOR = CDbl(Vetor1(XLI_Y, 7)) + CDbl(FunNuloVal(Vetor1(XLI_Y, 8)))
'                            End If
'
'                            If DtpBasePagto.Value > CDate(Vetor1(XLI_Y, 0)) Then
'                                If (DtpBasePagto.Value - CDate(Vetor1(XLI_Y, 0))) > 2 _
'                                    Or ((DtpBasePagto.Value - CDate(Vetor1(XLI_Y, 0))) = 2 And Weekday(CDate(Vetor1(XLI_Y, 0))) <> 7) _
'                                    Or ((DtpBasePagto.Value - CDate(Vetor1(XLI_Y, 0))) = 1 And Weekday(CDate(Vetor1(XLI_Y, 0))) <> 1) Then 'EM ATRASO E VENC NÃO É SÁBADO NEM DOMINGO
'                                    Vetor1(XLI_Y, 11) = "" 'DESAGIO
'
'                                    funCalculaJurosMulta XLF_JUROS, XLF_MULTA, Mid(Vetor1(XLI_Y, 1), 1, 22), DtpBasePagto.Value, XLF_VALOR
'                                    Vetor1(XLI_Y, 9) = XLF_JUROS
'                                    Vetor1(XLI_Y, 10) = XLF_MULTA
'                                Else
'                                    Vetor1(XLI_Y, 9) = ""
'                                    Vetor1(XLI_Y, 10) = ""
'                                    Vetor1(XLI_Y, 11) = "" 'DESAGIO
'                                End If
'                            ElseIf DtpBasePagto.Value < CDate(Vetor1(XLI_Y, 0)) Then 'PAGO ANTECIPADAMENTE
'                                Vetor1(XLI_Y, 9) = ""
'                                Vetor1(XLI_Y, 10) = ""
'                                'CÁLCULO DO DESÁGIO
'                                Vetor1(XLI_Y, 11) = Format(funCalculaDesagio(Mid(Vetor1(XLI_Y, 1), 1, 22), DtpBasePagto.Value, XLF_VALOR - CDbl(FunNuloVal(Vetor1(XLI_Y, 8)))), "standard")
'                            Else 'PAGO NO VENCIMENTO
'                                Vetor1(XLI_Y, 9) = ""
'                                Vetor1(XLI_Y, 10) = ""
'                                Vetor1(XLI_Y, 11) = ""
'                            End If
'                            Vetor1(XLI_Y, 14) = Format(XLF_VALOR + CDbl(FunNuloVal(Vetor1(XLI_Y, 9))) + CDbl(FunNuloVal(Vetor1(XLI_Y, 10))) + CDbl(FunNuloVal(Vetor1(XLI_Y, 13))) - CDbl(FunNuloVal(Vetor1(XLI_Y, 11))) - CDbl(FunNuloVal(Vetor1(XLI_Y, 12))), "standard")
'                            Vetor1(XLI_Y, 15) = Format(XLF_VALOR + CDbl(FunNuloVal(Vetor1(XLI_Y, 9))) + CDbl(FunNuloVal(Vetor1(XLI_Y, 10))) + CDbl(FunNuloVal(Vetor1(XLI_Y, 13))) - CDbl(FunNuloVal(Vetor1(XLI_Y, 11))) - CDbl(FunNuloVal(Vetor1(XLI_Y, 12))), "standard")
'                            '////////////////////////////
'                        Next
'                        XLO_TITULO.MoveNext
'                    Wend
'                    TDBGrid1.Refresh
'                    XLO_TITULO.Close
'
'                    If Not XLB_ACHEI Then
'                        'Abrir outro select para outros tipos de planos (<> Mensal)
'                        XLT_SQL = "SELECT * FROM consCAPTitulos WHERE empr_cd_empresa = " & PCodEmpresa & "" & _
'                                      " AND empd_cd_Empreendimento= '" & Trim(MskEmpreendimento.Text) & "'" & _
'                                      " AND imov_cd_Imovel='" & Trim(CboImovel.Text) & "'" & _
'                                      " AND cont_cd_Contrato='" & Trim(LblContrato2.Caption) & "'" & _
'                                      " AND titu_dt_Pagamento is null AND tipl_tx_Descricao <> 'MENSAL'" & _
'                                      " AND Natureza <> 'FINANCIAMENTO' ORDER BY titu_dt_Criacao"
'                        SubQOpenRecordset XLO_TITULO, XLT_SQL, Estatico
'
'                        While Not XLO_TITULO.EOF And Not XLB_ACHEI
'                            TDBGrid1.MoveFirst
'                            While Not TDBGrid1.EOF And Not XLB_ACHEI
'                                If TDBGrid1.Columns("Título") = XLO_TITULO!Titulo And TDBGrid1.Columns("Sel.").Value = "S" Then
'                                    Vetor1(TDBGrid1.bookmark, 8) = Format(XFF_SALDODEVEDOR * XLF_TAXASEGURO, "standard")
'                                    XLB_ACHEI = True
'                                Else
'                                    Vetor1(TDBGrid1.bookmark, 8) = ""
'                                End If
'                                TDBGrid1.MoveNext
'                            Wend
'                            XLO_TITULO.MoveNext
'                        Wend
'                        TDBGrid1.Refresh
'                    End If
'                Else 'Abrir outro select para outros tipos de planos (<> Mensal)
'                    XLO_TITULO.Close
'                    XLT_SQL = "SELECT * FROM consCAPTitulos WHERE empr_cd_empresa = " & PCodEmpresa & "" & _
'                                  " AND empd_cd_Empreendimento= '" & Trim(MskEmpreendimento.Text) & "'" & _
'                                  " AND imov_cd_Imovel='" & Trim(CboImovel.Text) & "'" & _
'                                  " AND cont_cd_Contrato='" & Trim(LblContrato2.Caption) & "'" & _
'                                  " AND titu_dt_Pagamento is null " & _
'                                  " AND Natureza <> 'FINANCIAMENTO' ORDER BY titu_dt_Criacao"
'                    SubQOpenRecordset XLO_TITULO, XLT_SQL, Estatico
'
'                    While Not XLO_TITULO.EOF And Not XLB_ACHEI
'                        TDBGrid1.MoveFirst
'                        While Not TDBGrid1.EOF And Not XLB_ACHEI
'                            If TDBGrid1.Columns("Título") = XLO_TITULO!Titulo And TDBGrid1.Columns("Sel.").Value = "S" Then
'                                Vetor1(TDBGrid1.bookmark, 8) = Format(XFF_SALDODEVEDOR * XLF_TAXASEGURO, "standard")
'                                XLB_ACHEI = True
'                            Else
'                                Vetor1(TDBGrid1.bookmark, 8) = ""
'                            End If
'                            TDBGrid1.MoveNext
'                        Wend
'                        XLO_TITULO.MoveNext
'                    Wend
'                    TDBGrid1.Refresh
'                End If
'            End If
'        Else
'            Vetor1(XLI_X, 8) = ""
'        End If
'    End If
End Sub

Private Sub CboCCorrente_Click(Area As Integer)
  
  If CboCCorrente.Text <> "" Then
     DatContaCorrente.Recordset.bookmark = CboCCorrente.SelectedItem
   End If
  
End Sub

Private Sub ChkCorrecao_Click()
    Dim XLI_X As Integer
    
    If CarregouGrid Then
        XFO_PAGAMENTO.MoveFirst
        For XLI_X = 0 To Vetor1.UpperBound(1)
            subCalculaCampos XFO_PAGAMENTO, XLI_X
            XFO_PAGAMENTO.MoveNext
        Next
        LblTotal.Caption = funSomaColuna(Vetor1, 14)
        TDBGrid1.Refresh
    End If
End Sub

Private Sub CmdContrato_Click()
    
    If Not FunVerifica_Permissao("TelaManContratos", "1") Then
        Exit Sub
    End If
    
    XInserir = "C"
    Chave = (MskEmpreendimento.Text + "." + CboImovel.Text + "." + LblContrato2.Caption)
    FrmManContratos.Show 1
End Sub

Private Sub CmdDesmarcar_Click()
    Dim XLF_VLCORRSEGURO As Double
    
    LblTotal.Caption = "0,00"
    If CmdDesmarcar.Caption = "Desmarcar Todos" Then
        'Desmarcar os registros
        TDBGrid1.MoveFirst
        LblTotal.Caption = "0,00"
        If Not TDBGrid1.EOF And Not TDBGrid1.BOF Then
            While Not TDBGrid1.EOF
                TDBGrid1.Columns("Sel.").Value = "N"
                If ChkCorrecao.Value = 0 Then
                    TDBGrid1.Columns("Valor Total") = Format(CDbl(TDBGrid1.Columns("Valor Histórico")) + CDbl(FunNuloVal(TDBGrid1.Columns("juros"))) + CDbl(FunNuloVal(TDBGrid1.Columns("Multa"))) + CDbl(FunNuloVal(TDBGrid1.Columns("Outros(+)"))) - CDbl(FunNuloVal(TDBGrid1.Columns("Deságio"))) - CDbl(FunNuloVal(TDBGrid1.Columns("Desconto"))), "standard")
                Else
                    TDBGrid1.Columns("Valor Total") = Format(CDbl(TDBGrid1.Columns("Valor Atual")) + CDbl(FunNuloVal(TDBGrid1.Columns("juros"))) + CDbl(FunNuloVal(TDBGrid1.Columns("Multa"))) + CDbl(FunNuloVal(TDBGrid1.Columns("Outros(+)"))) - CDbl(FunNuloVal(TDBGrid1.Columns("Deságio"))) - CDbl(FunNuloVal(TDBGrid1.Columns("Desconto"))), "standard")
                End If
                TDBGrid1.Columns("Valor a Pagar") = ""
                TDBGrid1.MoveNext
            Wend
        End If
        TDBGrid1.Refresh
        TDBGrid1.MoveFirst
        CmdDesmarcar.Caption = "Marcar Todos"
    ElseIf CmdDesmarcar.Caption = "Marcar Todos" Then
        
        TDBGrid1.MoveFirst

        If Not TDBGrid1.EOF And Not TDBGrid1.BOF Then
            While Not TDBGrid1.EOF
                TDBGrid1.Columns("Sel.").Value = "S"
                If ChkCorrecao.Value = 0 Then
                    TDBGrid1.Columns("Valor Total") = Format(CDbl(TDBGrid1.Columns("Valor Histórico")) + CDbl(FunNuloVal(TDBGrid1.Columns("juros"))) + CDbl(FunNuloVal(TDBGrid1.Columns("Multa"))) + CDbl(FunNuloVal(TDBGrid1.Columns("Outros(+)"))) - CDbl(FunNuloVal(TDBGrid1.Columns("Deságio"))) - CDbl(FunNuloVal(TDBGrid1.Columns("Desconto"))), "standard")
                    TDBGrid1.Columns("Valor a Pagar") = Format(CDbl(TDBGrid1.Columns("Valor Histórico")) + CDbl(FunNuloVal(TDBGrid1.Columns("juros"))) + CDbl(FunNuloVal(TDBGrid1.Columns("Multa"))) + CDbl(FunNuloVal(TDBGrid1.Columns("Outros(+)"))) - CDbl(FunNuloVal(TDBGrid1.Columns("Deságio"))) - CDbl(FunNuloVal(TDBGrid1.Columns("Desconto"))), "standard")
                Else
                    TDBGrid1.Columns("Valor Total") = Format(CDbl(TDBGrid1.Columns("Valor Atual")) + CDbl(FunNuloVal(TDBGrid1.Columns("juros"))) + CDbl(FunNuloVal(TDBGrid1.Columns("Multa"))) + CDbl(FunNuloVal(TDBGrid1.Columns("Outros(+)"))) - CDbl(FunNuloVal(TDBGrid1.Columns("Deságio"))) - CDbl(FunNuloVal(TDBGrid1.Columns("Desconto"))), "standard")
                    TDBGrid1.Columns("Valor a Pagar") = Format(CDbl(TDBGrid1.Columns("Valor Atual")) + CDbl(FunNuloVal(TDBGrid1.Columns("juros"))) + CDbl(FunNuloVal(TDBGrid1.Columns("Multa"))) + CDbl(FunNuloVal(TDBGrid1.Columns("Outros(+)"))) - CDbl(FunNuloVal(TDBGrid1.Columns("Deságio"))) - CDbl(FunNuloVal(TDBGrid1.Columns("Desconto"))), "standard")
                End If
                LblTotal.Caption = Format(CDbl(LblTotal.Caption) + CDbl(FunNuloVal(TDBGrid1.Columns("Valor a Pagar"))), "standard")
                TDBGrid1.MoveNext
            Wend
        End If
        TDBGrid1.Refresh
        TDBGrid1.MoveFirst
        CmdDesmarcar.Caption = "Desmarcar Todos"
    End If
End Sub

Private Sub cmdGravar_Click()
    Dim XLO_TITULO As New ADODB.Recordset
    Dim XLT_SQL As String
    
    Dim XLT_RESIDUO As String
    Dim XLT_PARCELA As String
    
    Dim XLO_MOEDA As ADODB.Recordset
    Dim XLB_QUITADO As Boolean
    Dim XLT_CODPLANO As String     'Código do TITULO
    Dim XLT_CODPARCELA As String   'Código da Parcela
    Dim XLT_CODRESIDUO As String   'Código do resíduo
    Dim XLF_VLRESIDUO As Double    'Valor do Resíduo
    Dim XLT_INDEXACAO1 As String   'tipo de indexação da moeda1
    Dim XLF_VLINDEX1 As Double     'Valor do Indexador 1
    Dim XLT_INDEXACAO2 As String   'tipo de indexação da moeda2
    Dim XLF_VLINDEX2 As Double     'Valor do Indexador 2
    Dim XLD_VENCRESIDUO As Date    'Data de vencimento do resíduo
    Dim XLF_TOTALTITULO As Double  'Valor total a pagar do título
    Dim XLF_VLCORRIGIDO As Double
    Dim XLT_CORRIGE As String      'Indica se corrige ou não a parcela
    Dim XLT_PRORATA As String     'Indica se a correcao é prorata
    Dim XLO_TIPOPLANO As ADODB.Recordset  'Recordset para a tabela de planos
    Dim XLI_TIPOPLANO As Integer   'Código do tipo do plano
    Dim XLO_QUITACAO As New ADODB.Recordset
    Dim XLI_X As Integer
    Dim XLB_TRANSACAOATIVA As Boolean
    Dim XLF_VALORCORRIGIDO As Double
    Dim XLT_VERIFICADATA As String
    Dim XLT_VERIFICAMESATUAL As String
    Dim XLT_VERIFICADATAATUAL As String
             
    XLB_TRANSACAOATIVA = False
    XLB_QUITADO = False
    
    XLT_VERIFICADATAATUAL = Format(Now, "dd/MM/yy")
     XLT_VERIFICAMESATUAL = Mid$(XLT_VERIFICADATAATUAL, 4, 2)
     XLT_VERIFICADATA = Format(DtpPagamento.Value, "dd/MM/yy")
     XLT_VERIFICADATA = Mid$(XLT_VERIFICADATA, 4, 2)
        If XLT_VERIFICADATA <> XLT_VERIFICAMESATUAL Then
          If Not FunVerifica_Permissao_SemMsg("FrmFinPagamentos", "5") Then
                MsgBox "Você não possui permissão para realizar pagamentos fora do mês corrente.", 16, "AVISO"
                DtpPagamento.SetFocus
            Exit Sub
          End If
        End If
    'Verifica a data de deposito
        XLT_VERIFICADATA = Format(DtpDeposito.Value, "dd/MM/yy")
        XLT_VERIFICADATA = Mid$(XLT_VERIFICADATA, 4, 2)
        If XLT_VERIFICADATA <> XLT_VERIFICAMESATUAL Then
            If Not FunVerifica_Permissao_SemMsg("FrmFinPagamentos", "5") Then
                MsgBox "Você não possui permissão para realizar pagamentos na data de depósito informada.", 16, "AVISO"
                DtpDeposito.SetFocus
                Exit Sub
            End If
        End If
        
    
    On Error GoTo TrataErroTitulos
     
    If XGB_PQUALIFIN Then 'Integrado com o Financeiro
       If FunObrigatorioCBO(CboCCorrente, "A Conta Corrente é Obrigatória") Then
         TabFinanceiro.Tab = 1
         CboCCorrente.SetFocus
         Exit Sub
       End If
       
       For XLI_X = 0 To Vetor1.UpperBound(1)
           If Vetor1(XLI_X, 4) = "S" Then
               If Vetor1(XLI_X, 17) = "" Then
                   MsgBox "O título " + Vetor1(XLI_X, 1) + " deve possuir uma Conta Contábil.", vbCritical, "ATENÇÃO"
                   Exit Sub
               End If
           
               If Vetor1(XLI_X, 18) = "" Then
                   MsgBox "O título " + Vetor1(XLI_X, 1) + " deve possuir um Centro de Custo.", vbCritical, "ATENÇÃO"
                   Exit Sub
               End If
           End If
       Next
    End If

    Me.MousePointer = VBRUN.MousePointerConstants.vbHourglass
    
    XLT_RESIDUO = "SELECT tipl_cd_TipoPlano FROM TiposPlanos WHERE tipl_tx_sigla = 'R'"
    SubQOpenRecordset XLO_TIPOPLANO, XLT_RESIDUO, Estatico
    XLI_TIPOPLANO = XLO_TIPOPLANO!tipl_cd_TipoPlano
    XLO_TIPOPLANO.Close
    Set XLO_TIPOPLANO = Nothing
    
    TDBGrid1.MoveFirst
    XFO_PAGAMENTO.MoveFirst
    While Not TDBGrid1.EOF
        If TDBGrid1.Columns("Sel.").Value = "S" Then
        
            XLT_CODPLANO = Mid$(Trim(TDBGrid1.Columns("Título")), 14, 2)
            XLT_CODPARCELA = Mid$(Trim(TDBGrid1.Columns("Título")), 17, 3)
            XLT_CODRESIDUO = Mid$(Trim(TDBGrid1.Columns("Título")), 21, 2)

            Conexao.BeginTrans
            
            XLB_TRANSACAOATIVA = True
            
            XLT_SQL = "UPDATE TITULOS SET "
                    
            'titu_vl_Desconto
            XLT_SQL = XLT_SQL & "titu_vl_Desconto = " & FunTrataFloat(FunNuloVal(TDBGrid1.Columns("Desconto"))) & ", "
            'titu_vl_Outros
            XLT_SQL = XLT_SQL & "titu_vl_Outros = " & FunTrataFloat(FunNuloVal(TDBGrid1.Columns("Outros(+)"))) & ", "
            'titu_vl_ValorPago
            XLT_SQL = XLT_SQL & "titu_vl_ValorPago = " & FunTrataFloat(FunNuloVal(TDBGrid1.Columns("Valor a Pagar"))) & ", "
            
            If OptIndexador1.Value = True Then
                'Valor Atual
                XLT_SQL = XLT_SQL & "titu_vl_Corrigido1 = " & FunTrataFloat(FunNuloVal(TDBGrid1.Columns(7))) & ", "
                
                If TDBGrid1.Columns("Moeda2").Value <> "" Then
                    XLF_VALORCORRIGIDO = Format(funCalculaCorrecaoMonetaria(XFO_PAGAMENTO, 2, DtpBasePagto.Value), "standard")
                    
                    'titu_vl_Corrigido2
                    XLT_SQL = XLT_SQL & "titu_vl_Corrigido2 = " & FunTrataFloat(FunNuloVal(Format(XLF_VALORCORRIGIDO + Format(FunCalculaJurosCEF(XFO_PAGAMENTO, DtpBasePagto.Value, XLF_VALORCORRIGIDO), "standard"), "standard"))) & ", "
                Else
                    'titu_vl_Corrigido2
                    XLT_SQL = XLT_SQL & "titu_vl_Corrigido2 = 0, "
                End If
                
                
            Else
                'Valor Atual
                'titu_vl_Corrigido2
                XLT_SQL = XLT_SQL & "titu_vl_Corrigido2 = " & FunTrataFloat(FunNuloVal(TDBGrid1.Columns(7))) & ", "
                
                If TDBGrid1.Columns("Moeda1").Value <> "" Then
                    XLF_VALORCORRIGIDO = Format(funCalculaCorrecaoMonetaria(XFO_PAGAMENTO, 1, DtpBasePagto.Value), "standard")
                    
                    'titu_vl_Corrigido1
                    XLT_SQL = XLT_SQL & "titu_vl_Corrigido1 = " & FunTrataFloat(FunNuloVal(Format(XLF_VALORCORRIGIDO + Format(FunCalculaJurosCEF(XFO_PAGAMENTO, DtpBasePagto.Value, XLF_VALORCORRIGIDO), "standard"), "standard"))) & ", "
                Else
                    'titu_vl_Corrigido1
                    XLT_SQL = XLT_SQL & "titu_vl_Corrigido1 = 0, "
                End If
                
            End If
            
            XLT_SQL = XLT_SQL & "titu_dt_Pagamento = " & FunNuloData(DtpPagamento.Value, NomeSgbd) & ", "
            XLT_SQL = XLT_SQL & "titu_dt_BasePagto = " & FunNuloData(DtpBasePagto.Value, NomeSgbd) & ", "
            XLT_SQL = XLT_SQL & "titu_dt_Deposito = " & FunNuloData(DtpDeposito.Value, NomeSgbd) & ", "
            XLT_SQL = XLT_SQL & "titu_tx_LocalPagto = " & CboLocalPagto.ItemData(CboLocalPagto.ListIndex) & ", "
            
            If (TxtDocPagto.Text <> "") Then
                XLT_SQL = XLT_SQL & "titu_tx_DocPagto = '" & TxtDocPagto.Text & "', "
            Else
                XLT_SQL = XLT_SQL & "titu_tx_DocPagto = NULL, "
            End If
                
            XLT_SQL = XLT_SQL & "coco_cd_Codigo = " & funNull(CboCCorrente.BoundText) & ", "
                         
            If ChkPreDatado.Value = 1 Then
                XLT_SQL = XLT_SQL & "titu_tx_PreDatado = 'S', "
            Else
                XLT_SQL = XLT_SQL & "titu_tx_PreDatado = 'N', "
            End If
            
            If OptIndexador1.Value = True Then 'escolhido valor indexado 1
                XLT_SQL = XLT_SQL & "titu_tx_IndicePagamento = 1, "
                'XLT_SQL = XLT_SQL & "moed_cd_Moeda3 = " & funNull(TDBGrid1.Columns("Moeda1")) & ", "
                XLT_SQL = XLT_SQL & "moed_cd_Moeda3 = " & FunNuloBanco(TDBGrid1.Columns("Moeda1")) & ", "
            Else 'escolhido valor indexado 2
                XLT_SQL = XLT_SQL & "titu_tx_IndicePagamento = 2, "
                XLT_SQL = XLT_SQL & "moed_cd_Moeda3 = " & funNull(TDBGrid1.Columns("Moeda2")) & ", "
            End If
            
            XLF_TOTALTITULO = (CDbl(TDBGrid1.Columns("Valor Atual")) + CDbl(FunNuloVal(TDBGrid1.Columns("Juros")) + CDbl(FunNuloVal(TDBGrid1.Columns("Multa"))) + CDbl(FunNuloVal(TDBGrid1.Columns("Outros(+)"))) - CDbl(FunNuloVal(TDBGrid1.Columns("Deságio"))) - CDbl(FunNuloVal(TDBGrid1.Columns("Desconto")))))
            XLF_VLCORRIGIDO = CDbl(TDBGrid1.Columns("Valor Atual"))
                        
            If ChkCorrecao.Value = 1 Then
                XLT_SQL = XLT_SQL & "titu_tx_CorrigeParcela = 'S', "
                XLT_CORRIGE = "1"
            Else
                XLT_SQL = XLT_SQL & "titu_tx_CorrigeParcela = 'N', "
                XLT_CORRIGE = "0"
                XLF_TOTALTITULO = (CDbl(TDBGrid1.Columns("Valor Histórico")) + CDbl(FunNuloVal(TDBGrid1.Columns("Juros")) + CDbl(FunNuloVal(TDBGrid1.Columns("Multa"))) + CDbl(FunNuloVal(TDBGrid1.Columns("Outros(+)"))) - CDbl(FunNuloVal(TDBGrid1.Columns("Deságio"))) - CDbl(FunNuloVal(TDBGrid1.Columns("Desconto")))))
            End If
            
            XLT_SQL = XLT_SQL & "titu_tx_TipoPag = " & funGeraTipoPag(XLT_CORRIGE, CDbl(FunNuloVal(TDBGrid1.Columns("Valor a Pagar"))), XLF_TOTALTITULO, XLF_VLCORRIGIDO, FunNuloVal(TDBGrid1.Columns("Vencimento")), DtpPagamento.Value) & " "
            
            '////////VERIFICAR SE HÁ RESÍDUO////////////
            If ChkResiduo.Value = 1 Or (XLT_CORRIGE = "0" And XLF_TOTALTITULO = (CDbl(FunNuloVal(TDBGrid1.Columns("Valor a Pagar"))))) Then
                'Deseja gerar resíduo ou não corrige e o valor pago não foi alterado (nesse último caso gera resíduo automaticamente)
                XLF_VLRESIDUO = funCalculaResiduo(XLF_TOTALTITULO, XLF_VLCORRIGIDO)
                If XLF_VLRESIDUO <> 0 Then
                    XLT_SQL = XLT_SQL & ", titu_vl_Desagio = " & ((CDbl(FunNuloVal(TDBGrid1.Columns("Valor a Pagar").Value)) * CDbl(FunNuloVal(TDBGrid1.Columns("Deságio").Value))) / XLF_TOTALTITULO) & ", "
                    XLT_SQL = XLT_SQL & "titu_vl_Multa = " & ((CDbl(FunNuloVal(TDBGrid1.Columns("Valor a Pagar").Value)) * CDbl(FunNuloVal(TDBGrid1.Columns("Multa").Value))) / XLF_TOTALTITULO) & ", "
                    XLT_SQL = XLT_SQL & "titu_vl_Juros = " & ((CDbl(FunNuloVal(TDBGrid1.Columns("Valor a Pagar").Value)) * CDbl(FunNuloVal(TDBGrid1.Columns("Juros").Value))) / XLF_TOTALTITULO) & " "
                    
                    XLT_PRORATA = DatImovel.Recordset.Fields("cont_tx_correcProRata")
                    
                    XLD_VENCRESIDUO = FunNuloVal(TDBGrid1.Columns("Prorrogação").Value)
                    
                    'Verificar o tipo de indexação
                    If TDBGrid1.Columns("Moeda1") <> "" Then
                        SubQOpenRecordset XLO_MOEDA, "SELECT * FROM Moedas WHERE Moedas.moed_cd_CodMoeda = " & FunNuloVal(TDBGrid1.Columns("Moeda1").Value) & "", Estatico
                        XLT_INDEXACAO1 = XLO_MOEDA!moed_tx_Indexacao
                        XLO_MOEDA.Close
                        
                        XLF_VLINDEX1 = FunCalcularValorIndexado(Format(XLF_VLRESIDUO, "standard"), DtpBasePagto.Value, FunNuloVal(TDBGrid1.Columns("Moeda1").Value), XLT_INDEXACAO1, XLT_PRORATA, TelaCotacoes)
                    Else
                        XLF_VLINDEX1 = 0
                    End If
                    
                    If TDBGrid1.Columns("Moeda2") <> "" Then
                        SubQOpenRecordset XLO_MOEDA, "SELECT * FROM Moedas WHERE Moedas.moed_cd_CodMoeda = " & FunNuloVal(TDBGrid1.Columns("Moeda2").Value) & "", Estatico
                        XLT_INDEXACAO2 = XLO_MOEDA!moed_tx_Indexacao
                        XLO_MOEDA.Close
                       Set XLO_MOEDA = Nothing
                        XLF_VLINDEX2 = FunCalcularValorIndexado(Format(XLF_VLRESIDUO, "standard"), DtpBasePagto.Value, FunNuloVal(TDBGrid1.Columns("Moeda2").Value), XLT_INDEXACAO2, XLT_PRORATA, TelaCotacoes)
                    Else
                        XLF_VLINDEX2 = 0
                    End If
    
                    'GRAVAR RESÍDUO
                     
                    
                    XLT_RESIDUO = "INSERT INTO Titulos (empr_cd_Empresa, empd_cd_Empreendimento, imov_cd_Imovel,cont_cd_Contrato, " & _
                                    "titu_cd_Plano,titu_cd_Parcela,titu_cd_Residuo,obse_cd_Observacao, " & _
                                    "moed_cd_Moeda1, moed_cd_Moeda2, titu_vl_Parcela, " & _
                                    "titu_vl_ParcelaIndex1, titu_vl_ParcelaIndex2, titu_dt_Vencimento, tipl_cd_TipoPlano, " & _
                                    "titu_vl_PercJurosFin, titu_vl_PercMultaMora, titu_vl_PercJurosMora, " & _
                                    "cere_cd_Pcr,plco_cd_Conta,coco_cd_Codigo,titu_dt_Base," & _
                                    "titu_dt_Prorrogacao,titu_tx_NaturezaPlano,titu_dt_Criacao,iden_cd_Projeto)" & _
                            " SELECT " & PCodEmpresa & ",'" & MskEmpreendimento.Text & "', " & _
                                    "'" & CboImovel.BoundText & "','" & LblContrato2.Caption & "'," & _
                                    "'" & XLT_CODPLANO & "','" & XLT_CODPARCELA & "'," & _
                                    "'" & Format(Int(XLT_CODRESIDUO) + 1, "00") & "'," & FunNuloBancoVal(TDBGrid1.Columns("Observação").Value) & ", " & _
                                    "" & FunNuloBancoVal(TDBGrid1.Columns("moed_cd_moeda1").Value) & "," & FunNuloBancoVal(TDBGrid1.Columns("moed_cd_moeda2").Value) & "," & _
                                    "" & FunNuloVal(FunTrataFloat(XLF_VLRESIDUO)) & "," & FunNuloVal(FunTrataFloat(XLF_VLINDEX1)) & "," & _
                                    "" & FunNuloVal(FunTrataFloat(XLF_VLINDEX2)) & "," & FunNuloData(TDBGrid1.Columns("Vencimento").Value, NomeSgbd) & "," & _
                                    "" & XLI_TIPOPLANO & "," & FunNuloVal(FunTrataFloat(TDBGrid1.Columns("titu_vl_PercJurosFin").Value)) & "," & _
                                    "" & FunNuloVal(FunTrataFloat(TDBGrid1.Columns("titu_vl_PercMultaMora").Value)) & "," & FunNuloVal(FunTrataFloat(TDBGrid1.Columns("titu_vl_PercJurosMora").Value)) & ", " & _
                                    "" & FunNuloBanco(TDBGrid1.Columns("CCusto").Value) & "," & FunNuloBanco(TDBGrid1.Columns("CConta").Value) & "," & FunNuloBancoVal(CboCCorrente.BoundText) & "," & _
                                    "" & FunNuloData(DtpBasePagto.Value, NomeSgbd) & "," & FunNuloData(TDBGrid1.Columns("Vencimento").Value, NomeSgbd) & "," & _
                                    "'" & TDBGrid1.Columns("titu_tx_naturezaplano").Value & "'," & FunNuloData(DtpPagamento.Value, NomeSgbd) & ", " & FunNuloBancoVal(TDBGrid1.Columns("iden_cd_Projeto").Value) & ""
            
            
                   Conexao.Execute XLT_RESIDUO
                Else
                    'titu_vl_Desagio
                    XLT_SQL = XLT_SQL & ", titu_vl_Desagio = " & FunTrataFloat(FunNuloVal(TDBGrid1.Columns("Deságio"))) & ", "
                    'titu_vl_Multa
                    XLT_SQL = XLT_SQL & "titu_vl_Multa = " & FunTrataFloat(FunNuloVal(TDBGrid1.Columns("Multa"))) & ", "
                    'titu_vl_Juros
                    XLT_SQL = XLT_SQL & "titu_vl_Juros = " & FunTrataFloat(FunNuloVal(TDBGrid1.Columns("Juros"))) & " "
                End If
            End If
            
            
            'WHERE
            XLT_SQL = XLT_SQL & "WHERE empr_cd_empresa = " & PCodEmpresa & "" & _
                 " AND empd_cd_Empreendimento= '" & MskEmpreendimento.Text & "' AND imov_cd_Imovel='" & CboImovel.BoundText & "'" & _
                 " AND cont_cd_Contrato='" & LblContrato2.Caption & "' AND titu_cd_Plano='" & XLT_CODPLANO & "'" & _
                 " AND titu_cd_Parcela='" & XLT_CODPARCELA & "' AND titu_cd_Residuo='" & XLT_CODRESIDUO & "'"
                 
            Conexao.Execute XLT_SQL
            
            '*********GRAVAÇÃO NO FINANCEIRO**********
            If XGB_PQUALIFIN Then
                
                If XGT_EXPCOMPROMCAPI = "S" Then
                     'LANCA O RESIDUO NO FINANCEIRO COMO COMPROMISSAR
                     If XLF_VLRESIDUO <> 0 Then
                        XLT_SQL = "SELECT * FROM ConsCAPTitulos WITH (ROWLOCK)  WHERE empr_cd_empresa = " & PCodEmpresa & "" & _
                            " AND empd_cd_Empreendimento= '" & MskEmpreendimento.Text & "' AND imov_cd_Imovel='" & CboImovel.BoundText & "'" & _
                            " AND cont_cd_Contrato='" & LblContrato2 & "' AND titu_cd_Plano='" & XLT_CODPLANO & "'" & _
                            " AND titu_cd_Parcela='" & XLT_CODPARCELA & "' AND titu_cd_Residuo='" & Format(Int(XLT_CODRESIDUO) + 1, "00") & "'"
                        SubQOpenRecordset XLO_TITULO, XLT_SQL, Estatico
                    
                        subGravaFinanceiro XLO_TITULO, "COMPROMISSAR", DatImovel.Recordset.Fields("focl_cd_FornCli"), XLO_TITULO!tipl_tx_Descricao, XLO_TITULO!Natureza
                        XLO_TITULO.Close
                        Set XLO_TITULO = Nothing
                    End If
                End If
                
                'LANCA O TITULO PAGO NO FINANCEIRO
                XLT_SQL = "SELECT * FROM ConsCAPTitulos WITH (ROWLOCK) WHERE empr_cd_empresa = " & PCodEmpresa & "" & _
                     " AND empd_cd_Empreendimento= '" & MskEmpreendimento.Text & "' AND imov_cd_Imovel='" & CboImovel.BoundText & "'" & _
                     " AND cont_cd_Contrato='" & LblContrato2 & "' AND titu_cd_Plano='" & XLT_CODPLANO & "'" & _
                     " AND titu_cd_Parcela='" & XLT_CODPARCELA & "' AND titu_cd_Residuo='" & XLT_CODRESIDUO & "'"
                SubQOpenRecordset XLO_TITULO, XLT_SQL, Estatico
                
                subGravaFinanceiro XLO_TITULO, "PAGAR", DatImovel.Recordset.Fields("focl_cd_FornCli"), XLO_TITULO!tipl_tx_Descricao, XLO_TITULO!Natureza
                XLO_TITULO.Close
                Set XLO_TITULO = Nothing
            End If
            
            Conexao.CommitTrans
            XLB_TRANSACAOATIVA = False
            
        End If
           
        TDBGrid1.MoveNext
        XFO_PAGAMENTO.MoveNext
        
        DoEvents
    Wend
    
On Error GoTo TrataErroContrato
    
     SubQOpenRecordset XLO_QUITACAO, "SELECT titu_dt_Vencimento FROM Titulos WITH (ROWLOCK) " & _
     " WHERE empr_cd_empresa = " & PCodEmpresa & " AND " & _
     " empd_cd_Empreendimento= '" & MskEmpreendimento.Text & "' AND " & _
     " imov_cd_Imovel='" & CboImovel.BoundText & "'" & " AND " & _
     " cont_cd_Contrato='" & LblContrato2 & "' AND " & _
     " titu_dt_Pagamento is null", Estatico

    If XLO_QUITACAO.EOF Then
        Conexao.Execute "UPDATE Contratos WITH (ROWLOCK) SET cont_dt_Quitacao=" & FunNuloData(DtpPagamento.Value, NomeSgbd) & _
        " WHERE empr_cd_empresa = " & PCodEmpresa & " AND " & _
        " empd_cd_Empreendimento= '" & MskEmpreendimento.Text & "' AND " & _
        " imov_cd_Imovel='" & CboImovel.BoundText & "'" & " AND " & _
        " cont_cd_Contrato='" & LblContrato2 & "'"
    
        XLB_QUITADO = True
    
        XLO_QUITACAO.Close
        Set XLO_QUITACAO = Nothing
    End If
              
    If XLB_QUITADO = True Then
        MsgBox "O Contrato foi Quitado com sucesso!", vbInformation, "ATENÇÃO"
    Else
        MsgBox "Títulos pagos com sucesso!", vbInformation, "INFORMAÇÃO"
    End If
     
    CmdRecibo.Enabled = True
    CmdImprimeReciboM.Enabled = True
    CmdPesquisar.Enabled = True
    CmdGravar.Enabled = False
    
    Me.MousePointer = VBRUN.MousePointerConstants.vbDefault
    
    Exit Sub
    
TrataErroTitulos:
    
    CmdPesquisar.Enabled = True
    
    Me.MousePointer = VBRUN.MousePointerConstants.vbDefault

    If XLB_TRANSACAOATIVA Then
        Conexao.RollbackTrans
    End If
    
    MsgBox "Não foi possível pagar o título " & TDBGrid1.Columns("Título") & Chr(13) & "Detalhe do Erro: " & Err.Number & " - " & Err.Description, vbCritical, "ATENÇÃO"
    
   Exit Sub
       
TrataErroContrato:

    Me.MousePointer = VBRUN.MousePointerConstants.vbDefault
   
    MsgBox "Os títulos foram pagos mas não foi possível verificar se o contrato foi quitado. Entre em contato com o suporte. " & Chr(13) & "Detalhe do Erro: " & Err.Number & " - " & Err.Description, vbCritical, "ATENÇÃO"
    
    Exit Sub
End Sub


'Ajustar a Tela à resolução do monitor
Sub subAjustaTelaLoc()
    Formulario.Top = 0
    Formulario.Height = Resolucaoy - 1380
    Formulario.Width = ResolucaoX - 60
    Formulario.Left = 5
    Formulario.TDBGrid1.Top = Formulario.TabFinanceiro.Top + Formulario.TabFinanceiro.Height + 50
    Formulario.TDBGrid1.Left = 75
    Formulario.TDBGrid1.Width = Formulario.Width - 150
    Formulario.TDBGrid1.Height = Formulario.Height - Formulario.TDBGrid1.Top - Formulario.PanBotoes.Height - 100
    Formulario.PanBotoes.Top = Formulario.TDBGrid1.Top + Formulario.TDBGrid1.Height + 50
    Formulario.PanBotoes.Left = 75 + (Formulario.TDBGrid1.Width - Formulario.PanBotoes.Width) / 2
    Formulario.TabFinanceiro.Left = 75
    Formulario.TabFinanceiro.Width = Formulario.Width - 150
End Sub

Private Sub subHabilitaBotoes()

    'Objetivo: Habilita os Botões da Tela
    CmdGravar.Enabled = True
    'CmdRecibo.Enabled = True
    CmdSair.Enabled = True
    CmdImprimir.Enabled = True
    
    'Verifica se tem Filtro Ativo
    If ContFil = 0 Then
       CmdRemoverFiltro.Enabled = False
    Else
       CmdRemoverFiltro.Enabled = True
    End If
End Sub

Private Sub CboEmpreendimento_Change()
    If CboEmpreendimento.Text <> "" Then
        DatEmpreendimento.Recordset.bookmark = CboEmpreendimento.SelectedItem
        MskEmpreendimento.Text = CboEmpreendimento.BoundText
        CboImovel.Enabled = True
        CboImovel.Text = ""
        subConectarControleDadosNV DatImovel, "SELECT * FROM ConsCapContratos WHERE empr_cd_Empresa= " & PCodEmpresa & " AND empd_cd_Empreendimento='" & MskEmpreendimento.Text & "' and cont_tx_Status='A' ORDER BY imov_cd_Imovel", Estatico
        CmdPesquisar.Enabled = False
        TDBGrid1.Enabled = False
    End If
End Sub

Private Sub CboEmpreendimento_DblClick(Area As Integer)
    TelaManEmpreendimentos.Show
    If FunLocalizaJanelasAtivas("TelaManEmpreendimentos") <> -1 Then
        TelaManEmpreendimentos.SetFocus
    End If
End Sub

Private Sub CboImovel_Change()
    Dim XLT_SQL As String
    If CboImovel.Text <> "" Then
        DatImovel.Recordset.bookmark = CboImovel.SelectedItem
        LblRazaoSocial.Caption = DatImovel.Recordset.Fields("Cliente")
        If Not IsNull(DatImovel.Recordset.Fields("cont_dt_base")) Then LblDtContrato2.Caption = Format(DatImovel.Recordset.Fields("cont_dt_base"), "dd/mm/yy")
        LblContrato2.Caption = DatImovel.Recordset.Fields!cont_cd_Contrato
        LblDtVenda2.Caption = Format(DatImovel.Recordset.Fields!cont_dt_Venda, "dd/mm/yy")
        LblTipoImovel2.Caption = FunNulo(DatImovel.Recordset.Fields!tiim_tx_Descricao)
        CmdPesquisar.Enabled = True
        TDBGrid1.Enabled = False
    End If
End Sub

Private Sub CmdImprimeReciboM_Click()
  
  Dim XLT_SQL As String
  Dim XLT_CONJUNCAO As String
  Dim XLT_TITULO As String
  Dim XLT_DOCPGTO As String
  Dim XGO_RSCABECALHO As New ADODB.Recordset
  Dim XGO_RSRELATORIO As New ADODB.Recordset
  Dim XLT_FORMAPGTO As String
  Dim XLT_DEPOSITO As String
  
  Me.MousePointer = vbHourglass
  
  XLT_SQL = " SELECT consCAPRecibo.*," & _
    "(SELECT COUNT(*) FROM Titulos AS tit WHERE tit.empr_cd_empresa=consCAPRecibo.empr_cd_empresa " & _
    " AND tit.empd_cd_Empreendimento = consCAPRecibo.empd_cd_Empreendimento " & _
    " AND tit.imov_cd_Imovel = consCAPRecibo.imov_cd_Imovel AND  tit.cont_cd_Contrato = consCAPRecibo.cont_cd_Contrato " & _
    " AND tit.titu_cd_Plano = consCAPRecibo.titu_cd_Plano AND tit.titu_cd_residuo = '00') AS TotalParcelas, " & _
    " IsNull((SELECT titu_vl_Parcela FROM ConscapTitulos as tit WHERE tit.empr_cd_empresa=consCAPRecibo.empr_cd_empresa" & _
    " AND tit.empd_cd_Empreendimento = consCAPRecibo.empd_cd_Empreendimento" & _
    " AND tit.imov_cd_Imovel = consCAPRecibo.imov_cd_Imovel AND tit.cont_cd_Contrato = consCAPRecibo.cont_cd_Contrato" & _
    " AND tit.titu_cd_Plano = consCAPRecibo.titu_cd_Plano AND tit.titu_cd_Parcela = consCAPRecibo.titu_cd_Parcela" & _
    " AND (tit.titu_cd_Residuo = (consCAPRecibo.titu_cd_residuo + 1))), 0) AS ValorResiduo" & _
    " FROM consCAPRecibo WHERE " & _
    " empr_cd_empresa = " & PCodEmpresa & " AND ("
   
  XLT_CONJUNCAO = ""
  TDBGrid1.MoveFirst
  While Not TDBGrid1.EOF
     If TDBGrid1.Columns("Sel.").Value = "S" Then
         XLT_TITULO = TDBGrid1.Columns("Título").Value
         XLT_SQL = XLT_SQL + XLT_CONJUNCAO + "Titulo = '" & XLT_TITULO & "'"
         XLT_CONJUNCAO = " OR "
     End If
     TDBGrid1.MoveNext
  Wend
  XLT_SQL = XLT_SQL + " )"
  CmdPesquisar.Enabled = True
  CmdGravar.Enabled = False
  
  XLT_DOCPGTO = TxtDocPagto.Text
  Set XGO_RSRELATORIO = ConexaoRelatorio.Execute(XLT_SQL)
  If XGO_RSRELATORIO.EOF Then
        MsgBox "O relatório não possui registros para serem visualizados", vbInformation, "Relatório sem dados"
        TelaFinPagtoMultiplos.MousePointer = vbDefault
        Exit Sub
  End If
  Set XGO_RSCABECALHO = ConexaoRelatorio.Execute("SELECT * FROM Empresas WHERE empr_cd_empresa = " & PCodEmpresa)
    
  If CboCCorrente.BoundText <> "" And CboLocalPagto.Text = "Cheque" Then
    XLT_DEPOSITO = DatContaCorrente.Recordset!banc_tx_Descricao & " - AG: " & DatContaCorrente.Recordset!coco_cd_Agencia & ""
    If FunNulo(DatContaCorrente.Recordset!coco_nr_Dag) <> "" Then
      XLT_DEPOSITO = XLT_DEPOSITO & "-" & DatContaCorrente.Recordset!coco_nr_Dag
    End If
    XLT_DEPOSITO = XLT_DEPOSITO & " - C/C: " & DatContaCorrente.Recordset!coco_tx_Conta
    If FunNulo(DatContaCorrente.Recordset!coco_nr_Dac) <> "" Then
      XLT_DEPOSITO = XLT_DEPOSITO & "-" & DatContaCorrente.Recordset!coco_nr_Dac
    End If
  Else
    XLT_DEPOSITO = ""
  End If

  'Carrega os relatórios
  FrmVisRelatorios.Caption = "Impressão de Recibos Multiplos"
  If XGT_MODELORECIBO = "1" Then
    Set XGR_RELATORIO = XGR_APLICACAO.OpenReport(App.Path + "\Relatorios\RelRecibo.rpt")
  Else
    Set XGR_RELATORIO = XGR_APLICACAO.OpenReport(App.Path + "\Relatorios\RelRecibo2.rpt")
  End If
  
  'Passa o recordset para o relatório
  XGR_RELATORIO.Database.SetDataSource XGO_RSRELATORIO

  XLT_FORMAPGTO = CboLocalPagto.Text
  With XGR_RELATORIO.FormulaFields
    .GetItemByName("CodEmpresa").Text = Chr(34) & CStr(PCodEmpresa) & Chr(34)
    .GetItemByName("docpgto").Text = Chr(34) & XLT_DOCPGTO & Chr(34)
    .GetItemByName("DepositoBanco").Text = Chr(34) & XLT_DEPOSITO & Chr(34)
  End With
  
  FrmVisRelatorios.SubAlterarDataSource_Subrelatorios XGR_RELATORIO, XGO_RSCABECALHO, "Empresas"
  FrmVisRelatorios.SubAlterarDataSource_Subrelatorios XGR_RELATORIO, XGO_RSCABECALHO, "Rodape"
  FrmVisRelatorios.SubVerRelatorio XGR_RELATORIO, ""
  Me.MousePointer = vbDefault
  
End Sub

Private Sub CmdImprimir_Click()
    
    On Error GoTo TrataErro:
    Dim GuardaY As Single
    Dim LargPapel As Single, AltPapel As Single, AreaImpressao As Single
    Dim TitRel As String
    Dim RecebeCampo As Variant
    Dim GuardaX As Single, GuardaYInicial As Single
    Dim QuebraY As Single
    Dim ImprimirLinha As Boolean
    Dim XLF_VALOR As Double
    Dim XLF_TOTALPAGAR As Double
    Dim XLI_QTDTITULOS As Integer
        
    TitRel = "Listagem de Títulos do Pagamento Múltiplo"
    
    Printer.ScaleMode = 7  'vbCentimeters
    
    Printer.Orientation = 1  'Portrait
    LargPapel = Printer.Width / 567   'Convertendo Twips para Cms
    AltPapel = Printer.Height / 567   'Convertendo Twips para Cms
    
    subRodape 2, AltPapel, TitRel
    subCabecalhoListagemRelatorio TitRel
    GuardaY = Printer.CurrentY
    GuardaYInicial = Printer.CurrentY
    
    AreaImpressao = AltPapel - AltRodape      ' Variável setada em Rodape()
    
    GuardaY = Printer.CurrentY + 0.3
    Printer.CurrentY = GuardaY
    GuardaY = Printer.CurrentY + 0.2
    subImprimeLinha 1.5, 20, 3
    subImprimeTexto "TÍTULO", 1.5, GuardaY, 7, "Arial", True, False
    subImprimeTexto "VENCTO.", 4.7, GuardaY, 7, "Arial", True, False
    subImprimeTexto "MULTA", 7, GuardaY, 7, "Arial", True, False
    subImprimeTexto "JUROS", 9, GuardaY, 7, "Arial", True, False
    subImprimeTexto "DESÁGIO", 11, GuardaY, 7, "Arial", True, False
    subImprimeTexto "DESCONTO", 13, GuardaY, 7, "Arial", True, False
    subImprimeTexto "OUTROS", 15.5, GuardaY, 7, "Arial", True, False
    subImprimeTexto "VALOR A PAGAR", 17.7, GuardaY, 7, "Arial", True, False
    subImprimeLinha 11.6, 15.7, 1
    subImprimeLinha 15.9, 20, 1
    GuardaY = Printer.CurrentY + 0.2
    subImprimeLinha 1.5, 20, 3
    Printer.Print ""
    
    TDBGrid1.MoveFirst
    
    XLI_QTDTITULOS = 0
    XLF_TOTALPAGAR = 0
    
    If Not (TDBGrid1.EOF And TDBGrid1.BOF) Then
        
        While Not (TDBGrid1.EOF)
            If (TDBGrid1.Columns("Sel.").Value = "S") Then
                If Printer.CurrentY >= (AreaImpressao - 0.35) Then
                    Printer.NewPage
                    subRodape 2, AltPapel, TitRel
                    subCabecalhoListagemRelatorio TitRel
                    
                    subImprimeLinha 1.5, 20, 3
                    GuardaY = Printer.CurrentY + 0.1
                    subImprimeTexto "TÍTULO", 1.5, GuardaY, 7, "Arial", True, False
                    subImprimeTexto "VENCTO.", 4.7, GuardaY, 7, "Arial", True, False
                    subImprimeTexto "MULTA", 7, GuardaY, 7, "Arial", True, False
                    subImprimeTexto "JUROS", 9, GuardaY, 7, "Arial", True, False
                    subImprimeTexto "DESÁGIO", 11, GuardaY, 7, "Arial", True, False
                    subImprimeTexto "DESCONTO", 13, GuardaY, 7, "Arial", True, False
                    subImprimeTexto "OUTROS", 15.5, GuardaY, 7, "Arial", True, False
                    subImprimeTexto "VALOR A PAGAR", 17.7, GuardaY, 7, "Arial", True, False
                    subImprimeLinha 11.6, 15.7, 1
                    subImprimeLinha 15.9, 20, 1
                    GuardaY = Printer.CurrentY + 0.1
                    subImprimeLinha 1.5, 20, 3
                    Printer.Print ""
                End If
                
                GuardaY = Printer.CurrentY
                RecebeCampo = TDBGrid1.Columns("Título")
                subQuebraTexto FunNulo(RecebeCampo), 1.5, 3, 7, "Arial", False, False
                QuebraY = Printer.CurrentY
                   
                Printer.CurrentY = GuardaY
                RecebeCampo = TDBGrid1.Columns("Vencimento")
                subQuebraTexto FunNulo(RecebeCampo), 4.7, 1.2, 7, "Arial", False, False
                
                Printer.CurrentY = GuardaY
                RecebeCampo = TDBGrid1.Columns("Multa")
                subQuebraTexto FunNulo(RecebeCampo), subRetornaImpNum(FunNulo(RecebeCampo), 5.9, 2, 7, "Arial", False, False), 2, 7, "Arial", False, False
                
                Printer.CurrentY = GuardaY
                RecebeCampo = TDBGrid1.Columns("Juros")
                subQuebraTexto FunNulo(RecebeCampo), subRetornaImpNum(FunNulo(RecebeCampo), 7.9, 2, 7, "Arial", False, False), 2, 7, "Arial", False, False
                
                Printer.CurrentY = GuardaY
                RecebeCampo = TDBGrid1.Columns("Deságio")
                subQuebraTexto FunNulo(RecebeCampo), subRetornaImpNum(FunNulo(RecebeCampo), 10.2, 2, 7, "Arial", False, False), 2, 7, "Arial", False, False
                
                Printer.CurrentY = GuardaY
                RecebeCampo = TDBGrid1.Columns("Desconto")
                subQuebraTexto FunNulo(RecebeCampo), subRetornaImpNum(FunNulo(RecebeCampo), 12.5, 2, 7, "Arial", False, False), 2, 7, "Arial", False, False
                
                Printer.CurrentY = GuardaY
                RecebeCampo = TDBGrid1.Columns("Outros(+)")
                subQuebraTexto FunNulo(RecebeCampo), subRetornaImpNum(FunNulo(RecebeCampo), 14.15, 2, 7, "Arial", False, False), 2, 7, "Arial", False, False
                
                Printer.CurrentY = GuardaY
                RecebeCampo = TDBGrid1.Columns("Valor A Pagar")
                subQuebraTexto FunNulo(RecebeCampo), subRetornaImpNum(FunNulo(RecebeCampo), 16.8, 3, 7, "Arial", False, False), 3, 7, "Arial", False, False
                              
                Printer.CurrentY = QuebraY + 0.2
                
                XLF_TOTALPAGAR = XLF_TOTALPAGAR + CCur(TDBGrid1.Columns("Valor A Pagar"))
                XLI_QTDTITULOS = XLI_QTDTITULOS + 1
            End If
            TDBGrid1.MoveNext
        Wend
        GuardaY = Printer.CurrentY + 0.4
        subImprimeLinha 1.5, 20, 3
        GuardaY = Printer.CurrentY + 0.2
        Printer.CurrentY = GuardaY
        RecebeCampo = "TOTAL: "
        subQuebraTexto FunNulo(RecebeCampo), 1.5, 1.6, 7, "Arial", True, False
        QuebraY = Printer.CurrentY
    
        Printer.CurrentY = GuardaY
        RecebeCampo = Format$(XLF_TOTALPAGAR, "STANDARD")
        subQuebraTexto FunNulo(RecebeCampo), subRetornaImpNum(FunNulo(RecebeCampo), 16.8, 3, 7, "Arial", False, False), 3, 7, "Arial", True, False
        If QuebraY < Printer.CurrentY Then
            QuebraY = Printer.CurrentY
        End If
        
        GuardaY = Printer.CurrentY + 0.3
        subImprimeTexto "TOTAL DE TÍTULOS:", 1.5, GuardaY, 7, "Arial", True, False
        
        Printer.CurrentY = GuardaY
        RecebeCampo = Format(XLI_QTDTITULOS, "000")
        subQuebraTexto FunNulo(RecebeCampo), 4, 4, 7, "Arial", True, False
        If QuebraY < Printer.CurrentY Then
            QuebraY = Printer.CurrentY
        End If
        GuardaY = Printer.CurrentY + 1#
        Printer.CurrentY = GuardaY
        
        TDBGrid1.MoveFirst
    Else
        Printer.Print ""
    End If
    Printer.EndDoc
   
Exit Sub

TrataErro:
    funTrataErros (ComMensagem)
    Exit Sub
    Resume
    
End Sub

Private Sub CmdPesquisar_Click()
    Me.MousePointer = vbHourglass
    
    'Recorset auxiliar
    XFT_SQLBASE = FunCriaConsultaBase(DtpBasePagto, NomeSgbd, "ConsCAPBasePagamento", 1)
    XFT_SQLBASE = XFT_SQLBASE & " empr_cd_empresa=" & PCodEmpresa
    XFT_SQL = XFT_SQLBASE & " AND left(Titulo,12) = " & _
      "'" & MskEmpreendimento.Text & "." & CboImovel.BoundText & "." & _
      DatImovel.Recordset.Fields!cont_cd_Contrato & "'" & _
      " AND titu_dt_pagamento is null AND (titu_tx_ExpBanco='N' OR titu_tx_ExpBanco is null) " & _
      " ORDER BY titu_dt_Vencimento,Titulo"
    
    If XFO_PAGAMENTO.State = adStateOpen Then
      XFO_PAGAMENTO.Close
    End If
    SubQOpenRecordset XFO_PAGAMENTO, XFT_SQL, Estatico
       
    If XFO_PAGAMENTO.EOF Then
      MsgBox "Não existem títulos pendentes para este contrato.", vbInformation, "ATENÇÃO"
      Me.MousePointer = vbDefault
      Exit Sub
    End If
       
    If OptIndexador1.Value = True Then
       
      If Not funVerirficaCotacao(XFO_PAGAMENTO, 1, DtpBasePagto.Value, True, TelaCotacoes) Then
        XFO_PAGAMENTO.Requery adCmdText
        If Not funVerirficaCotacao(XFO_PAGAMENTO, 1, DtpBasePagto.Value, False, TelaCotacoes) Then
          Exit Sub
        End If
      End If
      
    Else
    
      If Not funVerirficaCotacao(XFO_PAGAMENTO, 1, DtpBasePagto.Value, True, TelaCotacoes) Then
        XFO_PAGAMENTO.Requery adCmdText
        If Not funVerirficaCotacao(XFO_PAGAMENTO, 1, DtpBasePagto.Value, False, TelaCotacoes) Then
          Exit Sub
        End If
      End If
    End If
                
    XFT_SQL = "SELECT " & _
                "Titulo, Moeda1, Moeda2, titu_cd_Plano, titu_cd_Parcela, titu_cd_Residuo, empr_cd_Empresa, empd_cd_Empreendimento, imov_cd_Imovel, cont_cd_Contrato, tipl_cd_TipoPlano, coco_cd_Codigo, " & _
                "obse_cd_Observacao, moed_cd_Moeda1, moed_cd_Moeda2, titu_vl_Parcela, titu_vl_ParcelaIndex1, titu_vl_Corrigido1, titu_vl_ParcelaIndex2, titu_vl_Corrigido2, titu_dt_Vencimento, " & _
                "titu_vl_ValorPago, titu_dt_Pagamento, titu_vl_Multa, titu_vl_Juros, titu_vl_Seguro, titu_tx_IndicePagamento, titu_vl_Outros, titu_vl_Desconto, titu_vl_Desagio, " & _
                "titu_tx_Financia, titu_vl_PercJurosFin, titu_vl_PercMultaMora, titu_vl_PercJurosMora, " & _
                "titu_dt_Deposito, titu_dt_Repasse, titu_tx_LocalPagto, titu_tx_DocPagto, titu_dt_Prorrogacao, titu_tx_TipoPag, titu_dt_Base, cere_cd_Pcr, plco_cd_Conta, " & _
                "titu_tx_NaturezaPlano, moed_cd_Moeda3, titu_tx_CorrigeParcela, titu_tx_PreDatado, titu_dt_BasePagto, titu_dt_Criacao, nofi_cd_NotaFiscal, iden_cd_Projeto, " & _
                "titu_tx_ExpBanco , titu_tx_EntregaChave, titu_nr_ContratoBanco, obse_tx_Observacao, tipl_tx_Descricao, cont_tx_MultaProRata, focl_cd_FornCli, Natureza, 'S' as Selecionar " & _
                "FROM ConsCAPTitulos " & _
                "WHERE empr_cd_empresa = " & PCodEmpresa & " " & _
                "AND empd_cd_Empreendimento='" & MskEmpreendimento.Text & "' " & _
                "AND titu_dt_Pagamento is null AND (titu_tx_ExpBanco='N' OR titu_tx_ExpBanco is null) AND status='ATIVO' " & _
                "AND imov_cd_Imovel = '" & CboImovel.BoundText & "'" ' ORDER BY titu_dt_vencimento"
    
    subCarregaVetor XFT_SQL & " ORDER BY titu_dt_vencimento, Titulo ", Array("titu_dt_vencimento", "Titulo", "tipl_tx_descricao", "natureza", "Selecionar", "titu_vl_parcela", "titu_vl_ParcelaIndex1", "", "", "", "", "", "", "", "", "moed_cd_moeda1", "moed_cd_moeda2", "cere_cd_pcr", "plco_cd_conta", "titu_dt_Prorrogacao", "obse_cd_Observacao", "titu_tx_naturezaplano", "titu_vl_PercJurosFin", "titu_vl_PercMultaMora", "titu_vl_PercJurosMora", "iden_cd_Projeto"), Vetor1, TDBGrid1
       
    subCarregaVariaveis XFT_SQL
    
    TDBGrid1.Enabled = True
    TDBGrid1.SetFocus
    TDBGrid1.Refresh
    
    LblTotal.Caption = funSomaColuna(Vetor1, 14)
    
    filtra_especial = True
    
    If Vetor1(0, 0) <> Empty Then
        subHabilitaBotoes
        TDBGrid1.MoveFirst
    End If
    
    subTelaValoresGlobais "G"
    
    CmdRecibo.Enabled = False
    CmdPesquisar.Enabled = False
    CmdDesmarcar.Caption = "Desmarcar Todos"
    CarregouGrid = True
    Me.MousePointer = vbDefault
End Sub

'Carrega os valores no vetor
Sub subCarregaVetor(Sql As String, VetorCampos As Variant, Vetor As XArray, Grid As TDBGrid)   'Carrega o vetores
    Dim ResAvaliacao As Object
    Dim Cont As Integer
    Dim x As Integer
    
    'sql - string sql com a tabela ou consulta para carregar o vetor e o grid
    'VetorCampos - vetor simples com os campos na respectiva ordem do grid
    'Vetor - vetor que será preenchido
    'Grid - grid que será preenchido
    
    SubQOpenRecordset ResAvaliacao, Sql, Estatico
    Vetor.ReDim 0, 0, 0, Grid.Columns.Count - 1
    Vetor.Clear
    Grid.Array = Vetor
    XFO_PAGAMENTO.MoveFirst
    If Not (ResAvaliacao.EOF And ResAvaliacao.BOF) Then
        ResAvaliacao.MoveFirst
        Do While Not ResAvaliacao.EOF
            If Vetor.UpperBound(1) <> 0 Or Vetor(0, 0) <> Empty Then
                Vetor.Insert 1, Vetor.UpperBound(1) + 1
            End If
            Cont = Vetor.UpperBound(1)
            For x = 0 To Grid.Columns.Count - 1
                If VetorCampos(x) <> "" Then
                    subSetaVal Vetor, Cont, x, FunNulo(ResAvaliacao(VetorCampos(x)))
                End If
            Next
            subCalculaCampos XFO_PAGAMENTO, Vetor.UpperBound(1)
            XFO_PAGAMENTO.MoveNext
            ResAvaliacao.MoveNext
        Loop
    End If
    
    With Grid
        .ReBind
    End With
    
    ResAvaliacao.Close
End Sub


Sub subCalculaCampos(XLO_TITULO As ADODB.Recordset, XLI_X As Integer)
    Dim XLT_SQL As String           'Armazena SQL's
    Dim XLF_VALOR As Double         'Armazena o valor sobre o qual deverá incidir juros
    'Dim XLO_TITULO As Object        'Tabela com os dados do contrato e dos títulos correspondentes
    Dim XLF_VLCORRIGIDO As Double   'Valor corrigido para ser utilizado no cálculo do seguro com base no saldo devedor
    Dim XLF_VLTOTCORRIGIDO As Double 'Valor total corrigido e desagiado para ser utilizado no cálculo do seguro com base no saldo devedor
    Dim XLF_TAXASEGURO As Double    'Taxa de seguro
    Dim XLF_VLSEGURO As Double      'Valor do Seguro calculado com base no Saldo Devedor
    Dim XLT_TIPOSEGURO As String    'Tipo de seguro cadastrado no contrato 1-Pelo Saldo Devedor 2-Pela Parcela ""Não possui seguro
    Dim XLB_CALCSEGUROS As String   'Informa se deseja ou nao calcular o seguro
    Dim XLF_JUROS As Double         'Valor do Juros calculado na funcao CalculaJurosMulta
    Dim XLF_MULTA As Double         'Valor da multa calculado na funcao CalculaJurosMulta
    
    Dim XLF_VALORCORRIGIDO As Double
    
    '/////Definição dos campos do vetor/////
    'Vetor(0,0)=Vencimento
    'Vetor(0,1)=Titulo
    'Vetor(0,2)=Tipo Plano
    'Vetor(0,3)=Natureza
    'Vetor(0,4)=Selcionar
    'Vetor(0,5)=Valor
    'Vetor(0,6)=Corrige Parcela (excluído por Newton)
    'Vetor(0,6)=Correção 1
    'Vetor(0,7)=Vl Atual 1
    'Vetor(0,9)=Marcar 1 (excluído por Newton)
    'Vetor(0,8)=Correção 2 (excluído por Newton)
    'Vetor(0,9)=Vl Atual 2 (excluído por Newton)
    'Vetor(0,12)=Marcar 2 (excluído por Newton)
    'Vetor(0,8)=Seguro (EXCLUÍDO - AUTORIZADO POR D. Mª JOSÉ - C. ANDRADE
    'Vetor(0,8)=Juros
    'Vetor(0,9)=multa
    'Vetor(0,10)=deságio
    'Vetor(0,11)=desconto
    'Vetor(0,12)=outros
    'Vetor(0,13)=Valor Total
    'Vetor(0,14)=Valor Pagto
    'Vetor(0,15)=Código da Moeda1
    'Vetor(0,16)=Código da Moeda2
    '///////////////////////////////////////
    
    XLB_CALCSEGUROS = True
    'CÁLCULO DA CORREÇÃO MONETÁRIA
    If (OptIndexador1.Value = True And Vetor1(XLI_X, 15) <> "") Then  'Não possui o indexador
        'Correção Monetária
        XLF_VALORCORRIGIDO = Format(funCalculaCorrecaoMonetaria(XLO_TITULO, 1, DtpBasePagto.Value), "standard")
        
        'Correção
        Vetor1(XLI_X, 6) = Format(XLF_VALORCORRIGIDO - FunNuloVal(Vetor1(XLI_X, 5)), "standard")
        
        'Juros CEF
        Vetor1(XLI_X, 7) = Format(XLF_VALORCORRIGIDO + Format(FunCalculaJurosCEF(XLO_TITULO, DtpBasePagto.Value, XLF_VALORCORRIGIDO), "standard"), "standard")
        
    ElseIf (OptIndexador2.Value = True And Vetor1(XLI_X, 16) <> "") Then
        'Correção Monetária
        XLF_VALORCORRIGIDO = Format(funCalculaCorrecaoMonetaria(XLO_TITULO, 2, DtpBasePagto.Value), "standard")
        
        'Correçção
        Vetor1(XLI_X, 6) = Format(XLF_VALORCORRIGIDO - FunNuloVal(Vetor1(XLI_X, 5)), "standard")
        
        'Juros CEF
        Vetor1(XLI_X, 7) = Format(XLF_VALORCORRIGIDO + Format(FunCalculaJurosCEF(XLO_TITULO, DtpBasePagto.Value, XLF_VALORCORRIGIDO), "standard"), "standard")
    Else
        XLF_VALORCORRIGIDO = Format(Vetor1(XLI_X, 5), "standard")
        Vetor1(XLI_X, 7) = Format(XLF_VALORCORRIGIDO + Format(FunCalculaJurosCEF(XLO_TITULO, DtpBasePagto.Value, XLF_VALORCORRIGIDO), "standard"), "standard")
        Vetor1(XLI_X, 6) = "0,00"
    End If
    
    If ChkCorrecao.Value = 0 Then
        XLF_VALOR = CDbl(Vetor1(XLI_X, 5))  'Valor
    Else
        XLF_VALOR = CDbl(FunNuloVal(Vetor1(XLI_X, 7))) 'Corrigido
    End If
      
    If DtpBasePagto.Value > CDate(Vetor1(XLI_X, 0)) Then
    
      XLF_JUROS = FunCalculaJuros(XLO_TITULO, DtpBasePagto.Value, XLF_VALOR)
      XLF_MULTA = FunCalculaMulta(XLO_TITULO, DtpBasePagto.Value, XLF_VALOR)
      
      Vetor1(XLI_X, 8) = FunTrataZero(Format(XLF_JUROS, "standard"))
      Vetor1(XLI_X, 9) = FunTrataZero(Format(XLF_MULTA, "standard"))
      Vetor1(XLI_X, 10) = "" 'DESAGIO
        
    ElseIf DtpBasePagto.Value < CDate(Vetor1(XLI_X, 0)) Then 'PAGO ANTECIPADAMENTE
        
        Vetor1(XLI_X, 8) = "" 'Juros
        Vetor1(XLI_X, 9) = "" 'Multa
        'CÁLCULO DO DESÁGIO
        Vetor1(XLI_X, 10) = Format(funCalculaDesagio(XLO_TITULO, DtpBasePagto.Value, XLF_VALOR - CDbl(FunNuloVal(Vetor1(XLI_X, 9)))), "standard")
    Else 'PAGO NO VENCIMENTO
        Vetor1(XLI_X, 8) = ""
        Vetor1(XLI_X, 9) = ""
        Vetor1(XLI_X, 10) = ""
    End If
    
    Vetor1(XLI_X, 13) = Format(CDbl(FunNuloVal(Vetor1(XLI_X, 7))) + CDbl(FunNuloVal(Vetor1(XLI_X, 8)) + CDbl(FunNuloVal(Vetor1(XLI_X, 9))) - CDbl(FunNuloVal(Vetor1(XLI_X, 10)))), "standard")
    If Vetor1(XLI_X, 4) = "S" Then
        Vetor1(XLI_X, 14) = Format(XLF_VALOR + CDbl(FunNuloVal(Vetor1(XLI_X, 8)) + CDbl(FunNuloVal(Vetor1(XLI_X, 9))) - CDbl(FunNuloVal(Vetor1(XLI_X, 10)))), "standard")
    Else
        Vetor1(XLI_X, 14) = ""
    End If
End Sub


Private Sub CmdRecibo_Click()
  Dim XLT_SQL As String
  Dim XLT_CONJUNCAO As String
  Dim XLT_TITULO As String
  Dim XLT_DOCPGTO As String
  Dim XGO_RSCABECALHO As New ADODB.Recordset
  Dim XGO_RSRELATORIO As New ADODB.Recordset
  Dim XLT_FORMAPGTO As String
  Dim XLT_DEPOSITO As String
    
  Me.MousePointer = vbHourglass
  
  XLT_SQL = " SELECT consCAPRecibo.*," & _
    "(SELECT COUNT(*) FROM Titulos AS tit WHERE  tit.empd_cd_Empreendimento = consCAPRecibo.empd_cd_Empreendimento " & _
    " AND tit.imov_cd_Imovel = consCAPRecibo.imov_cd_Imovel AND  tit.cont_cd_Contrato = consCAPRecibo.cont_cd_Contrato " & _
    " AND tit.titu_cd_Plano = consCAPRecibo.titu_cd_Plano AND tit.titu_cd_residuo = '00') AS TotalParcelas, " & _
    " IsNull((SELECT titu_vl_Parcela FROM ConscapTitulos as tit WHERE   tit.empd_cd_Empreendimento = consCAPRecibo.empd_cd_Empreendimento" & _
    " AND tit.imov_cd_Imovel = consCAPRecibo.imov_cd_Imovel AND tit.cont_cd_Contrato = consCAPRecibo.cont_cd_Contrato" & _
    " AND tit.titu_cd_Plano = consCAPRecibo.titu_cd_Plano AND tit.titu_cd_Parcela = consCAPRecibo.titu_cd_Parcela" & _
    " AND (tit.titu_cd_Residuo = (consCAPRecibo.titu_cd_residuo + 1))), 0) AS ValorResiduo" & _
    " FROM consCAPRecibo WHERE " & _
    " empr_cd_empresa = " & PCodEmpresa & "AND ("
   
  XLT_CONJUNCAO = ""
  TDBGrid1.MoveFirst
  While Not TDBGrid1.EOF
     If TDBGrid1.Columns("Sel.").Value = "S" Then
         XLT_TITULO = TDBGrid1.Columns("Título").Value
         XLT_SQL = XLT_SQL + XLT_CONJUNCAO + "Titulo = '" & XLT_TITULO & "'"
         XLT_CONJUNCAO = " OR "
     End If
     TDBGrid1.MoveNext
  Wend
  XLT_SQL = XLT_SQL + " )"
  CmdPesquisar.Enabled = True
  CmdGravar.Enabled = False
  
  XLT_DOCPGTO = TxtDocPagto.Text
  
  If CboCCorrente.BoundText <> "" And CboLocalPagto.Text = "Cheque" Then
    XLT_DEPOSITO = DatContaCorrente.Recordset!banc_tx_Descricao & " - AG: " & DatContaCorrente.Recordset!coco_cd_Agencia & ""
    If FunNulo(DatContaCorrente.Recordset!coco_nr_Dag) <> "" Then
      XLT_DEPOSITO = XLT_DEPOSITO & "-" & DatContaCorrente.Recordset!coco_nr_Dag
    End If
    XLT_DEPOSITO = XLT_DEPOSITO & " - C/C: " & DatContaCorrente.Recordset!coco_tx_Conta
    If FunNulo(DatContaCorrente.Recordset!coco_nr_Dac) <> "" Then
      XLT_DEPOSITO = XLT_DEPOSITO & "-" & DatContaCorrente.Recordset!coco_nr_Dac
    End If
  Else
    XLT_DEPOSITO = ""
  End If

   Set XGO_RSRELATORIO = ConexaoRelatorio.Execute(XLT_SQL)
   If XGO_RSRELATORIO.EOF Then
        MsgBox "O relatório não possui registros para serem visualizados", vbInformation, "Relatório sem dados"
        TelaFinPagtoMultiplos.MousePointer = vbDefault
        Exit Sub
    End If
   Set XGO_RSCABECALHO = ConexaoRelatorio.Execute("SELECT * FROM Empresas WHERE empr_cd_empresa = " & PCodEmpresa)

  FrmVisRelatorios.Caption = "Impressão de Recibos Multiplos"
  
  'Carrega os relatórios de acordo com modelo
  If XGT_MODELORECIBO = "1" Then
    Set XGR_RELATORIO = XGR_APLICACAO.OpenReport(App.Path + "\Relatorios\RelReciboMultiplo.rpt")
  Else
    Set XGR_RELATORIO = XGR_APLICACAO.OpenReport(App.Path + "\Relatorios\RelReciboMultiplo2.rpt")
  End If

  'Passa o recordset para o relatório
  XGR_RELATORIO.Database.SetDataSource XGO_RSRELATORIO

'  'Ativa os Relatórios
  XLT_FORMAPGTO = CboLocalPagto.Text
  With XGR_RELATORIO.FormulaFields
    .GetItemByName("CodEmpresa").Text = Chr(34) & CStr(PCodEmpresa) & Chr(34)
    .GetItemByName("docpgto").Text = Chr(34) & XLT_DOCPGTO & Chr(34)
    .GetItemByName("DepositoBanco").Text = Chr(34) & XLT_DEPOSITO & Chr(34)
  End With
    
  FrmVisRelatorios.SubAlterarDataSource_Subrelatorios XGR_RELATORIO, XGO_RSCABECALHO, "Empresas"
  FrmVisRelatorios.SubAlterarDataSource_Subrelatorios XGR_RELATORIO, XGO_RSCABECALHO, "Rodape"
  FrmVisRelatorios.SubVerRelatorio XGR_RELATORIO, ""
  Me.MousePointer = vbDefault

End Sub

Private Sub CmdRemoverFiltro_Click()
    
    Set Formulario = TelaFinPagtoMultiplos
    subTelaValoresGlobais "P"
    TelaFiltro.Show 1 'Mostra a tela de filtros no modo modal
    subTelaValoresGlobais "G"

    If Filtrou = True Then
        subCarregaVetor FRM_FiltroAtual, Array("titu_dt_Vencimento", "Titulo", "tipl_tx_descricao", _
                    "natureza", "Selecionar", "titu_vl_parcela", "titu_vl_ParcelaIndex1", _
                    "", "", "", "", "", "", "", "", "moed_cd_moeda1", "moed_cd_moeda2"), Vetor1, TDBGrid1
    End If
    
    subHabilitaBotoes
   
    TDBGrid1.SetFocus
End Sub

Private Sub CmdRemoverFiltro_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    MdiPrincipal.BarraStatus.Panels(1).Text = "Mostra os filtros ativos da tabela"
End Sub

Private Sub CmdImprimir_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    MdiPrincipal.BarraStatus.Panels(1).Text = "Imprime os dados da tabela"
End Sub

Private Sub CmdSair_Click()
    MdiPrincipal.BarraStatus.Panels(3).Text = ""
    MdiPrincipal.BarraStatus.Panels(1).Text = ""
    Unload Me
End Sub

Private Sub CmdSair_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    MdiPrincipal.BarraStatus.Panels(1).Text = "Retorna à tela principal"
End Sub


Private Sub CmdOrdem_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    MdiPrincipal.BarraStatus.Panels(1).Text = "Ordena os registros atuais da tabela"
End Sub

Private Sub CmdTitulo_Click()
    Chave = TDBGrid1.Columns("Titulo")
    XInserir = "C"
    FrmFinTitulos.Show 1
End Sub

Private Sub DtpBasePagto_Change()
    CmdPesquisar.Enabled = True
    TDBGrid1.Enabled = False
End Sub

Private Sub DtpPagamento_LostFocus()
    CmdPesquisar.Enabled = True
    TDBGrid1.Enabled = False
    DtpBasePagto.Value = DtpPagamento.Value
    DtpDeposito.Value = DtpPagamento.Value
    DtpBasePagto.SetFocus
End Sub


Private Sub MskEmpreendimento_GotFocus()
    Call subSelecionaMSK(MskEmpreendimento)
End Sub

Private Sub MskEmpreendimento_LostFocus()
    If MskEmpreendimento.Text <> "" Then
        SubPreencheComZero MskEmpreendimento
        CboEmpreendimento.BoundText = MskEmpreendimento.Text
        If CboEmpreendimento.Text = "" Then
            MsgBox "Não existe Empreendimento com este código.", vbCritical + vbOKOnly, "ATENÇÃO"
            MskEmpreendimento.SetFocus
            Exit Sub
        End If
    End If
End Sub

Private Sub MskEmpreendimento_KeyPress(KeyAscii As Integer)
If KeyAscii >= 97 And KeyAscii <= 122 Then
      KeyAscii = KeyAscii - 32
    End If
End Sub

Private Sub OptIndexador1_Click()
    Dim XLI_X As Integer
    
    If CarregouGrid Then
      XFO_PAGAMENTO.MoveFirst
      For XLI_X = 0 To Vetor1.UpperBound(1)
          subCalculaCampos XFO_PAGAMENTO, XLI_X
          XFO_PAGAMENTO.MoveNext
      Next
      
      LblTotal.Caption = funSomaColuna(Vetor1, 14)
      TDBGrid1.Refresh
    End If
End Sub

Private Sub OptIndexador2_Click()
    Dim XLI_X As Integer
    
    If CarregouGrid Then
      XFO_PAGAMENTO.MoveFirst
      For XLI_X = 0 To Vetor1.UpperBound(1)
          subCalculaCampos XFO_PAGAMENTO, XLI_X
          XFO_PAGAMENTO.MoveNext
      Next
      LblTotal.Caption = funSomaColuna(Vetor1, 14)
      TDBGrid1.Refresh
    End If
End Sub

Private Sub TDBGrid1_AfterColUpdate(ByVal ColIndex As Integer)
    Dim XLF_VLCORRSEGURO As Double  'Calcular o vl corrido para o seguro (cotação cheia)
    Dim XLF_TOTALAPAGAR As Double  'Totaliza o valor a pagar
    Dim XLI_POSICAO As Integer      'Posicao no grid
    Dim XLF_VALOR As Double
    Dim XLF_JUROS As Double
    Dim XLF_MULTA As Double
    
    XLI_POSICAO = TDBGrid1.bookmark
        If TDBGrid1.Columns("Sel.").Value = "S" Then
            Vetor1(TDBGrid1.bookmark, 4) = "S"
            If ChkCorrecao.Value = 0 Then
                TDBGrid1.Columns("Valor Total") = Format(CDbl(TDBGrid1.Columns("Valor Histórico")) + CDbl(FunNuloVal(TDBGrid1.Columns("juros"))) + CDbl(FunNuloVal(TDBGrid1.Columns("Multa"))) + CDbl(FunNuloVal(TDBGrid1.Columns("Outros(+)"))) - CDbl(FunNuloVal(TDBGrid1.Columns("Deságio"))) - CDbl(FunNuloVal(TDBGrid1.Columns("Desconto"))), "standard")
                If ColIndex <> 14 Then 'Não é a coluna Valor a Pagar que está sendo alterada
                    TDBGrid1.Columns("Valor a Pagar") = Format(CDbl(TDBGrid1.Columns("Valor Histórico")) + CDbl(FunNuloVal(TDBGrid1.Columns("juros"))) + CDbl(FunNuloVal(TDBGrid1.Columns("Multa"))) + CDbl(FunNuloVal(TDBGrid1.Columns("Outros(+)"))) - CDbl(FunNuloVal(TDBGrid1.Columns("Deságio"))) - CDbl(FunNuloVal(TDBGrid1.Columns("Desconto"))), "standard")
                End If
                Vetor1(TDBGrid1.bookmark, 14) = TDBGrid1.Columns("Valor a Pagar").Value
            Else
                
                TDBGrid1.Columns("Valor Total") = Format(CDbl(TDBGrid1.Columns("Valor Atual")) + CDbl(FunNuloVal(TDBGrid1.Columns("juros"))) + CDbl(FunNuloVal(TDBGrid1.Columns("Multa"))) + CDbl(FunNuloVal(TDBGrid1.Columns("Outros(+)"))) - CDbl(FunNuloVal(TDBGrid1.Columns("Deságio"))) - CDbl(FunNuloVal(TDBGrid1.Columns("Desconto"))), "standard")
                If ColIndex <> 14 Then 'Não é a coluna Valor a Pagar que está sendo alterada
                    TDBGrid1.Columns("Valor a Pagar") = Format(CDbl(TDBGrid1.Columns("Valor Atual")) + CDbl(FunNuloVal(TDBGrid1.Columns("juros"))) + CDbl(FunNuloVal(TDBGrid1.Columns("Multa"))) + CDbl(FunNuloVal(TDBGrid1.Columns("Outros(+)"))) - CDbl(FunNuloVal(TDBGrid1.Columns("Deságio"))) - CDbl(FunNuloVal(TDBGrid1.Columns("Desconto"))), "standard")
                End If
                Vetor1(TDBGrid1.bookmark, 14) = TDBGrid1.Columns("Valor a Pagar").Value
            End If
        Else
            Vetor1(TDBGrid1.bookmark, 4) = "N"
            TDBGrid1.Columns("Valor a Pagar") = ""
            Vetor1(TDBGrid1.bookmark, 14) = ""
        End If
    LblTotal.Caption = funSomaColuna(Vetor1, 14)
End Sub

Private Sub TDBGrid1_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    If ColIndex = 9 Or ColIndex = 10 Or ColIndex = 11 Or ColIndex = 13 Or ColIndex = 14 Or ColIndex = 16 Then
        If TDBGrid1.Columns(ColIndex) <> "" Then
            If Not IsNumeric(TDBGrid1.Columns(ColIndex)) Then
                MsgBox "O conteúdo de " & TDBGrid1.Columns(ColIndex).Caption & " deve ser numérico", vbCritical, "ATENÇÃO"
                Cancel = True
            Else
                If CSng(TDBGrid1.Columns(ColIndex)) < 0 Then
                    MsgBox "O valor de " & TDBGrid1.Columns(ColIndex).Caption & " deve ser maior ou igual a zero", vbCritical, "ATENÇÃO"
                    Cancel = True
                End If
            End If
        End If
    End If
End Sub

Private Sub TDBGrid1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then   'Verifica se o botão da direita foi pressionado
        If TDBGrid1.ColContaining(x) = 4 Or TDBGrid1.ColContaining(x) = 7 _
            Or TDBGrid1.ColContaining(x) = 8 Or TDBGrid1.ColContaining(x) = 9 _
            Or TDBGrid1.ColContaining(x) = 10 Or TDBGrid1.ColContaining(x) = 11 _
            Or TDBGrid1.ColContaining(x) = 12 Or TDBGrid1.ColContaining(x) = 13 _
            Or TDBGrid1.ColContaining(x) = 14 Then
            MsgBox "Esta coluna não pode ser filtrada."
        Else
            Set Formulario = TelaFinPagtoMultiplos
            subTelaValoresGlobais "P"
            
            FunExecutaFiltroUnbound Array("titu_dt_vencimento", "Titulo", "tipl_tx_descricao", _
                    "natureza", "Selecionar", "titu_vl_parcela", "titu_vl_ParcelaIndex1", _
                    "", "", "", "", "", "", "", "", "moed_cd_moeda1", "moed_cd_moeda2"), Vetor1, TDBGrid1, Formulario, x, XFT_SQL
            subTelaValoresGlobais "G"
        End If
    End If
    subHabilitaBotoes
End Sub

Private Sub Form_Activate()
    Set Formulario = TelaFinPagtoMultiplos
    If ResolucaoX < 10000 Then   '640x480
        TDBGrid1.Splits(0).Size = 4
        TDBGrid1.Splits(1).Size = 7
    ElseIf ResolucaoX > 10000 And ResolucaoX < 13000 Then  '800x600
        TDBGrid1.Splits(0).Size = 4
        TDBGrid1.Splits(1).Size = 10
    Else        '1024x728 ou superior
        TDBGrid1.Splits(0).Size = 4
        TDBGrid1.Splits(1).Size = 14
    End If
    
    If CarregouGrid Then
        If Not PrimeiraVez Then
            subTelaValoresGlobais "P"
        End If

        PrimeiraVez = False
    End If
End Sub

Private Sub Form_Load()
    
    Set Formulario = TelaFinPagtoMultiplos
    
    Call subAjustaTelaLoc
    
    subManutencaoJanelasAtivas "I", "TelaFinPagtoMultiplos"
               
    subConectarControleDadosNV DatEmpreendimento, "SELECT empd_cd_Empreendimento,empd_tx_Nome FROM Empreendimentos WHERE empr_cd_empresa = " & PCodEmpresa & " ORDER BY empd_tx_nome", Estatico
    subConectarControleDadosNV DatContaCorrente, "SELECT * FROM ConsGENCCCombo WHERE empr_cd_empresa = " & PCodEmpresa & " ORDER BY descCombo", Estatico
    
    DtpPagamento.Value = Format(Now, "dd/MM/yyyy")
    DtpBasePagto.Value = Format(Now, "dd/MM/yyyy")
    DtpDeposito.Value = Format(Now, "dd/MM/yy")
    LblRazaoSocial.Caption = ""
    LblTipoImovel2.Caption = ""
    LblContrato2.Caption = ""
    LblDtVenda2.Caption = ""
    LblDtContrato2.Caption = ""
    CboLocalPagto.ListIndex = 0
    subHabilitaBotoes
    filtra_especial = True
    
    If Not XGB_PQUALIFIN Then ' Se não possuir o financeiro o combo de conta corrente estará desabilitado.
        CboCCorrente.Enabled = False
    End If

    If BarraFerramentasPressionada = True Then
       BarraFerramentasPressionada = False
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    XFO_PAGAMENTO.Close
    Set XFO_PAGAMENTO = Nothing
    
    subManutencaoJanelasAtivas "R", "TelaFinPagtoMultiplos"
End Sub

Private Sub TDBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If TDBGrid1.Columns("Sel.").Value = "S" Then
        TDBGrid1.Splits(1).Columns("Desconto").Locked = False
        TDBGrid1.Splits(1).Columns("Outros(+)").Locked = False
        TDBGrid1.Splits(1).Columns("Valor a Pagar").Locked = False
        If TDBGrid1.Columns("Vencimento").Value <> "" Then
            If CDate(TDBGrid1.Columns("Vencimento").Value) <= DtpBasePagto.Value Then
                TDBGrid1.Splits(1).Columns("Juros").Locked = False
                TDBGrid1.Splits(1).Columns("Multa").Locked = False
            Else
                TDBGrid1.Splits(1).Columns("Juros").Locked = True
                TDBGrid1.Splits(1).Columns("Multa").Locked = True
            End If
        End If
    ElseIf TDBGrid1.Columns("Sel.").Value = "N" Then
        TDBGrid1.Splits(1).Columns("Juros").Locked = True
        TDBGrid1.Splits(1).Columns("Multa").Locked = True
        TDBGrid1.Splits(1).Columns("Desconto").Locked = True
        TDBGrid1.Splits(1).Columns("Outros(+)").Locked = True
        TDBGrid1.Splits(1).Columns("Valor a Pagar").Locked = True
    End If
End Sub
