VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{6FBA474E-43AC-11CE-9A0E-00AA0062BB4C}#1.0#0"; "SYSINFO.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm MdiPrincipal 
   BackColor       =   &H00C0C0C0&
   Caption         =   "QualiCapi"
   ClientHeight    =   4995
   ClientLeft      =   1140
   ClientTop       =   2775
   ClientWidth     =   9420
   Icon            =   "MdiPrincipal.frx":0000
   LinkTopic       =   "MDIForm1"
   NegotiateToolbars=   0   'False
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar BarraFerramentas 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   9420
      _ExtentX        =   16616
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ListaImagens"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   19
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Empreendimentos"
            Object.ToolTipText     =   "Empreendimentos"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Imoveis"
            Object.ToolTipText     =   "Imóveis"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "PessoaFisica"
            Object.ToolTipText     =   "Pessoa Física"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "PessoaJuridica"
            Object.ToolTipText     =   "Pessoa Jurídica"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Cotacoes"
            Object.ToolTipText     =   "Cotações"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Planos"
            Object.ToolTipText     =   "Planos"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Titulos"
            Object.ToolTipText     =   "Títulos"
            Object.Tag             =   ""
            ImageIndex      =   7
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "PagamentoUnico"
            Object.ToolTipText     =   "Pagamento Único"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "PagamentoMultiplo"
            Object.ToolTipText     =   "Pagamento Múltiplo"
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button14 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "EmpresaAtiva"
            Object.ToolTipText     =   "Alterar Empres Ativa"
            Object.Tag             =   ""
            ImageIndex      =   10
         EndProperty
         BeginProperty Button15 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button16 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Impressoras"
            Object.Tag             =   "Impressoras"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button17 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button18 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Sobre"
            Object.ToolTipText     =   "Sobre"
            Object.Tag             =   ""
            ImageIndex      =   12
         EndProperty
         BeginProperty Button19 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Sair"
            Object.ToolTipText     =   "Sair"
            Object.Tag             =   ""
            ImageIndex      =   13
         EndProperty
      EndProperty
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   375
         HelpContextID   =   10
         Left            =   6840
         Picture         =   "MdiPrincipal.frx":0E42
         ScaleHeight     =   375
         ScaleWidth      =   1815
         TabIndex        =   5
         Top             =   0
         Width           =   1815
      End
      Begin Threed.SSPanel Label1 
         Height          =   420
         Left            =   8640
         TabIndex        =   4
         Top             =   -15
         Width           =   750
         _Version        =   65536
         _ExtentX        =   1323
         _ExtentY        =   741
         _StockProps     =   15
         Caption         =   "Não escreva a versão aqui"
         BackColor       =   -2147483633
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
      End
      Begin VB.PictureBox p_hlp_indice 
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   5775
         ScaleHeight     =   330
         ScaleWidth      =   390
         TabIndex        =   3
         Top             =   15
         Width           =   390
      End
   End
   Begin ComctlLib.StatusBar BarraStatus 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   1
      Top             =   4665
      Width           =   9420
      _ExtentX        =   16616
      _ExtentY        =   529
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   7
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   10585
            MinWidth        =   10585
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "Mensagens do Sistema"
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   3069
            MinWidth        =   3069
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "Empresa Atual"
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Width           =   2822
            MinWidth        =   2822
            TextSave        =   ""
            Key             =   ""
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
            Key             =   ""
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
            Key             =   ""
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
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "Modo Insert"
         EndProperty
         BeginProperty Panel7 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   3528
            MinWidth        =   3528
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "Usuário Atual"
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
      Top             =   4965
      Width           =   9420
      _ExtentX        =   16616
      _ExtentY        =   53
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
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
      MaskColor       =   65280
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   13
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MdiPrincipal.frx":1699
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MdiPrincipal.frx":1BEB
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MdiPrincipal.frx":213D
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MdiPrincipal.frx":268F
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MdiPrincipal.frx":2BE1
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MdiPrincipal.frx":2EFB
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MdiPrincipal.frx":374D
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MdiPrincipal.frx":3A67
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MdiPrincipal.frx":3FB9
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MdiPrincipal.frx":42D3
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MdiPrincipal.frx":4825
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MdiPrincipal.frx":4937
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MdiPrincipal.frx":4BC9
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu MnuManutencao 
      Caption         =   "&Manutenção"
      Begin VB.Menu MnuManEmpreendimentos 
         Caption         =   "&Empreendimentos"
      End
      Begin VB.Menu MnuManImoveis 
         Caption         =   "&Imóveis"
      End
      Begin VB.Menu MnuManClientes 
         Caption         =   "&Clientes"
         Begin VB.Menu MnuManCliPessoaFisica 
            Caption         =   "Pessoa &Física"
         End
         Begin VB.Menu MnuManCliPessoaJuridica 
            Caption         =   "Pessoa &Juridica"
         End
      End
      Begin VB.Menu MnuManFiadores 
         Caption         =   "&Fiadores"
      End
      Begin VB.Menu MnuManTabelas 
         Caption         =   "&Tabelas"
         Begin VB.Menu MnuManTabTiposImoveis 
            Caption         =   "&Tipos de Imóveis"
         End
         Begin VB.Menu MnuManTabTiposPlanos 
            Caption         =   "Tipos de P&lanos"
         End
         Begin VB.Menu MnuManTabIndexadores 
            Caption         =   "&Indexadores"
         End
         Begin VB.Menu MnuManTabCorretores 
            Caption         =   "&Corretores"
         End
         Begin VB.Menu MnuManTabObservacoes 
            Caption         =   "&Observações"
         End
         Begin VB.Menu MnuManTabProfissoes 
            Caption         =   "&Profissões"
         End
         Begin VB.Menu MnuManTabCartas 
            Caption         =   "C&artas"
         End
         Begin VB.Menu MnuManTabBancos 
            Caption         =   "&Bancos"
         End
         Begin VB.Menu MnuManTabCCorrentes 
            Caption         =   "Contas Co&rrentes"
         End
         Begin VB.Menu MnuManIdentificadores_Click 
            Caption         =   "Identi&ficadores"
         End
         Begin VB.Menu MnuManCnae_Click 
            Caption         =   "C&nae"
            Begin VB.Menu MnuManSecaoCnae_Click 
               Caption         =   "Seção"
            End
            Begin VB.Menu MnuManDivisaoCnae_Click 
               Caption         =   "Divisão"
            End
            Begin VB.Menu MnuManGrupoCnae_Click 
               Caption         =   "Grupo"
            End
            Begin VB.Menu MnuManClasseCnae_Click 
               Caption         =   "Classe"
            End
            Begin VB.Menu MnuManSubClasseCnae_Click 
               Caption         =   "SubClasse"
            End
         End
         Begin VB.Menu SEP 
            Caption         =   "-"
         End
         Begin VB.Menu mniFeriados 
            Caption         =   "Feriados"
         End
      End
   End
   Begin VB.Menu MnuFinanceiro 
      Caption         =   "&Financeiro"
      Begin VB.Menu MnuFinPlanos 
         Caption         =   "&Planos"
      End
      Begin VB.Menu MnuFinTitulos 
         Caption         =   "&Títulos"
      End
      Begin VB.Menu MnuFinCotacoes 
         Caption         =   "&Cotações"
      End
      Begin VB.Menu MnuEspaco 
         Caption         =   "-"
      End
      Begin VB.Menu MnuFinPagtoUnico 
         Caption         =   "Pagamento &Único"
      End
      Begin VB.Menu MnuFinPagtoMultiplo 
         Caption         =   "Pagamento &Múltiplo"
      End
      Begin VB.Menu MnuAtualizarExpBancaria 
         Caption         =   "&Exportacao Bancária"
         Begin VB.Menu MnuExportacaoBradesco 
            Caption         =   "&Bradesco"
         End
         Begin VB.Menu MnuExportacaoItau 
            Caption         =   "&Itaú"
         End
         Begin VB.Menu MnuexportacaoHSBC 
            Caption         =   "&HSBC"
         End
         Begin VB.Menu MnuExportacaoReal 
            Caption         =   "&Real"
         End
         Begin VB.Menu MnuExportacaoSantander 
            Caption         =   "&Santander"
         End
         Begin VB.Menu MnuExportacaoSudameris 
            Caption         =   "&Sudameris"
         End
         Begin VB.Menu MnuExportacaoCaixa 
            Caption         =   "&Caixa"
         End
         Begin VB.Menu MnuExportacaoUnibanco 
            Caption         =   "&Unibanco"
         End
      End
      Begin VB.Menu MnuExportacaoDimob 
         Caption         =   "Exportação &Dimob"
      End
      Begin VB.Menu MnuEspaco2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuFinAtualizarQualifin 
         Caption         =   "Atualização do &Qualifin"
      End
   End
   Begin VB.Menu MnuRelatorios 
      Caption         =   "&Relatórios"
      Begin VB.Menu MnuItau 
         Caption         =   "Itaú"
         Begin VB.Menu MnuAnexoFluxoRecebido 
            Caption         =   "&Anexo de Fluxo Recebido"
         End
         Begin VB.Menu MnuPosicaoComercializacao 
            Caption         =   "&Posição de Comercialização"
         End
         Begin VB.Menu MnuRelacaoCompradores 
            Caption         =   "&Relação de Compradores"
         End
      End
      Begin VB.Menu mnuReal 
         Caption         =   "Real"
         Begin VB.Menu mnuFluxoRecebiboReal 
            Caption         =   "&Fluxo Recebíveis"
         End
      End
      Begin VB.Menu MnuSantender 
         Caption         =   "Santander"
         Begin VB.Menu MnuPosicaoComercializacaoSantander 
            Caption         =   "&Posição de Comercialização"
         End
         Begin VB.Menu MnuFluxoRecebidoSantander 
            Caption         =   "&Fluxo Recebíveis"
         End
      End
      Begin VB.Menu MnuBradesco 
         Caption         =   "Bradesco"
         Begin VB.Menu MnuPosicaoComercializacaoBradesco 
            Caption         =   "&Posição de Comercialização"
         End
      End
      Begin VB.Menu MnuSpace 
         Caption         =   "-"
      End
      Begin VB.Menu MnuRelAssistenciaTecnica 
         Caption         =   "Assistência Técnica"
      End
      Begin VB.Menu MnuRelSituacaoImoveis 
         Caption         =   "&Situação dos Imóveis"
      End
      Begin VB.Menu MnuRelSituacaoContratos 
         Caption         =   "Situação dos Contratos"
      End
      Begin VB.Menu MnuMnuRelClientes 
         Caption         =   "Relação de Clientes"
         Begin VB.Menu MnuRelClientes 
            Caption         =   "Analítico"
         End
         Begin VB.Menu MnuRelCompradores 
            Caption         =   "Sintético"
         End
      End
      Begin VB.Menu MnuRelDemonstrativoPorCliente 
         Caption         =   "&Demonstrativo Por Empreendimento"
      End
      Begin VB.Menu MnuRelExtratoContrato 
         Caption         =   "&Extrato do Contrato"
         Begin VB.Menu MnuRelExtratoContratoAnalitico 
            Caption         =   "&Analítico"
         End
         Begin VB.Menu MnuRelExtratoContratoAnaliticoCliente 
            Caption         =   "Analítico (&Cliente)"
         End
         Begin VB.Menu MnuRelExtratoContratoReal 
            Caption         =   "Analítico &Real"
         End
         Begin VB.Menu MnuRelExtratoContratoSintetico 
            Caption         =   "&Sintético"
         End
         Begin VB.Menu MnuRelExtratoContratoResumoDebitos 
            Caption         =   "&Resumo dos Débitos"
         End
         Begin VB.Menu MnuRelExtratoContratoPagamentos 
            Caption         =   "&Pagamentos Realizados"
         End
      End
      Begin VB.Menu MnuRelPrevisaoRecebimento 
         Caption         =   "&Previsão de Recebimento"
         Begin VB.Menu MnuAnaliticoDiario 
            Caption         =   "&Analítico Diário"
         End
         Begin VB.Menu MnuAnaliticoMensal 
            Caption         =   "Analítico &Mensal"
         End
         Begin VB.Menu MnuSinteticoDiario 
            Caption         =   "&Sintetico Diário"
         End
         Begin VB.Menu MnuSinteticoMensal 
            Caption         =   "Sintetico &Mensal"
         End
      End
      Begin VB.Menu MnuRelTitulosVencidos 
         Caption         =   "Títulos &Vencidos"
         Begin VB.Menu MnuRelTitulosVencidosAnalitico 
            Caption         =   "&Analítico"
         End
         Begin VB.Menu MnuRelTitulosVencidosPorCliente 
            Caption         =   "Por &Cliente"
         End
      End
      Begin VB.Menu MnuRelTitulosLiquidados 
         Caption         =   "&Títulos Liquidados"
         Begin VB.Menu MnuRelTitulosLiquidadosAnalitico 
            Caption         =   "&Analítico"
         End
         Begin VB.Menu MnuRelTitulosLiquidadosSintetico 
            Caption         =   "&Sintético"
         End
         Begin VB.Menu MnuRelPorNatureza 
            Caption         =   "&Por Natureza"
         End
         Begin VB.Menu MnuRelTitulosLiquidadosPIS 
            Caption         =   "P&IS/COFINS"
         End
      End
      Begin VB.Menu MnuRelFluxoGeral 
         Caption         =   "&Fluxo Geral"
         Begin VB.Menu MnuRelfluxoGeralVencimento 
            Caption         =   "Mês de &Vencimento"
         End
         Begin VB.Menu MnuRelFluxoGeralPagamento 
            Caption         =   "Mês de &Pagamento"
         End
      End
      Begin VB.Menu MnuResumoMensal 
         Caption         =   "Resumo &Mensal"
      End
      Begin VB.Menu MnuResumoFGeral 
         Caption         =   "Resumo &Geral"
      End
      Begin VB.Menu MnuRelMapaGeral 
         Caption         =   "&Mapa Geral"
      End
      Begin VB.Menu MnuRelSeguros 
         Caption         =   "&Seguros"
         Begin VB.Menu MnuRelContratosComSeguro 
            Caption         =   "&Contratos Com Seguro"
         End
         Begin VB.Menu MnuRelSegurosAPagar 
            Caption         =   "&A Pagar"
         End
         Begin VB.Menu MnuRelSegurosPagos 
            Caption         =   "&Pagos"
         End
      End
      Begin VB.Menu MnuRelComissoes 
         Caption         =   "&Comissões dos Corretores"
      End
      Begin VB.Menu MnuRelImoveisQuitados 
         Caption         =   "Imóveis Quitados"
      End
      Begin VB.Menu mnuRelReapasse 
         Caption         =   "Repasse Contratual"
      End
      Begin VB.Menu MnuRealMapaContabil 
         Caption         =   "Mapa Contábil"
      End
      Begin VB.Menu mnuRelCartas 
         Caption         =   "Cartas"
         Begin VB.Menu mnuRelCartCobranca 
            Caption         =   "&Cobrança"
         End
         Begin VB.Menu mnuRelCartEspeciais 
            Caption         =   "&Especiais"
         End
      End
      Begin VB.Menu MnuRelEtiquetas 
         Caption         =   "&Etiquetas e Envelopes"
      End
      Begin VB.Menu MnuRelCotacoes 
         Caption         =   "&Cotações"
         Begin VB.Menu MnuRelListagemCotacao 
            Caption         =   "&Listagem"
         End
         Begin VB.Menu MnuRelListagemProRata 
            Caption         =   "Listagem &Pró-Rata"
         End
      End
      Begin VB.Menu MnuRelPromissorias 
         Caption         =   "&Promissórias"
      End
      Begin VB.Menu MnuRelDeclaracoes 
         Caption         =   "&Declarações"
         Begin VB.Menu MnuRelItiv 
            Caption         =   "&Declaração ITIV"
         End
         Begin VB.Menu MnuRelCOAF 
            Caption         =   "Transações &Imobiliárias"
         End
      End
      Begin VB.Menu MnuRecibos 
         Caption         =   "&Recibos"
      End
   End
   Begin VB.Menu MnuControle 
      Caption         =   "&Controle"
      Begin VB.Menu MnuAtualizar 
         Caption         =   "A&tualizar QualiCapi"
      End
      Begin VB.Menu Sepz 
         Caption         =   "-"
      End
      Begin VB.Menu MnuExportacaoBancaria 
         Caption         =   "A&tualizar Exportação Bancária"
         Begin VB.Menu MnuAtualizarExpBradesco 
            Caption         =   "&ExpBradesco"
         End
         Begin VB.Menu MnuAtualizarExpHSBC 
            Caption         =   "&ExpHSBC"
         End
         Begin VB.Menu MnuAtualizarExpItau 
            Caption         =   "&ExpItau"
         End
         Begin VB.Menu MnuAtualizarExpReal 
            Caption         =   "&ExpReal"
         End
         Begin VB.Menu MnuAtualizarExpSantander 
            Caption         =   "&ExpSantander"
         End
         Begin VB.Menu MnuAtualizarExpSudameris 
            Caption         =   "&ExpSudameris"
         End
         Begin VB.Menu MnuAtualizarExpUnibanco 
            Caption         =   "&ExpUnibanco"
         End
      End
      Begin VB.Menu sept 
         Caption         =   "-"
      End
      Begin VB.Menu MnuConImpressora 
         Caption         =   "&Configurar Impressora"
      End
      Begin VB.Menu MnuConSenhas 
         Caption         =   "&Senhas"
      End
      Begin VB.Menu MnuConEmpresa 
         Caption         =   "&Padrões da Empresa"
      End
      Begin VB.Menu MnuAtiva 
         Caption         =   "&Alterar Empresa Ativa"
      End
      Begin VB.Menu MnuConBanco 
         Caption         =   "Configurar Sis&tema"
      End
      Begin VB.Menu MnuConForm 
         Caption         =   "Configurações &Formularios"
      End
      Begin VB.Menu MnuEspaco3 
         Caption         =   "-"
      End
      Begin VB.Menu MnuExportEmpresas 
         Caption         =   "Exportação Entre Empresas"
      End
      Begin VB.Menu MnuImportEmpresas 
         Caption         =   "Importação Entre Empresas"
      End
      Begin VB.Menu MnuConImportacao 
         Caption         =   "Importacao"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu MnuJanelas 
      Caption         =   "&Janelas"
      WindowList      =   -1  'True
   End
   Begin VB.Menu MnuFinalizar 
      Caption         =   "Fi&nalizar"
   End
   Begin VB.Menu MnuAjuda 
      Caption         =   "&Ajuda"
   End
End
Attribute VB_Name = "MdiPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long

Option Explicit

Private Sub BarraFerramentas_ButtonClick(ByVal Button As ComctlLib.Button)
    
    Dim T As Variant
    BarraFerramentasPressionada = True
    Select Case Button.Key
    
    Case "Empreendimentos"
       
       If Not FunVerifica_Permissao("TelaManEmpreendimentos", "1") Then
         Exit Sub
       End If
       
       TelaManEmpreendimentos.Show
       If FunLocalizaJanelasAtivas("TelaManEmpreendimentos") <> -1 Then
          TelaManEmpreendimentos.SetFocus
       End If
    
    Case "Imoveis"
       
       If Not FunVerifica_Permissao("TelaManImoveis", "1") Then
         Exit Sub
       End If
       
       FrmManImoveis.Show vbModal
       BarraFerramentasPressionada = False
    
    Case "PessoaFisica"
    
      XInserir = "I"
      FornClie = "C" 'Cliente
      Pessoa = "F" 'Pessoa Física
      Chave = -1
      
      XFormForn = "MdiCapi"
      FrmFornecedores.Show 1
        
    Case "PessoaJuridica"
    
      XInserir = "I"
      FornClie = "C" 'Cliente
      Pessoa = "J" 'Pessoa Juridica
      Chave = -1
      
      XFormForn = "MdiCapi"
      FrmFornecedores.Show 1
        
    Case "Cotacoes"
         
         If Not FunVerifica_Permissao("TelaCotacoes", "1") Then
           Exit Sub
         End If
         
         TelaCotacoes.Show vbModal
         BarraFerramentasPressionada = False
    
    Case "Planos"
         
         If Not FunVerifica_Permissao("FrmFinPlanos", "1") Then
           Exit Sub
         End If
         
         FrmFinPlanos.Show vbModal
         BarraFerramentasPressionada = False
    
    Case "Titulos"
        
        If Not FunVerifica_Permissao("TelaFinTitulos", "1") Then
           Exit Sub
        End If
         
        TelaFinTitulos.Show
        If FunLocalizaJanelasAtivas("TelaFinTitulos") <> -1 Then
            TelaFinTitulos.SetFocus
        End If
    
    Case "PagamentoUnico"
        
        If Not FunVerifica_Permissao("FrmFinPagamentos", "1") Then
           Exit Sub
        End If
        XFormulario = ""
        FrmFinPagamentos.Show vbModal
        BarraFerramentasPressionada = False
    
    Case "PagamentoMultiplo"
        
        If Not FunVerifica_Permissao("TelaFinPagtoMultiplos", "1") Then
           Exit Sub
        End If
        
        TelaFinPagtoMultiplos.Show
        If FunLocalizaJanelasAtivas("TelaFinPagtoMultiplos") <> -1 Then
            TelaFinPagtoMultiplos.SetFocus
        End If
        
    Case "EmpresaAtiva"
        Call MnuAtiva_Click
        BarraFerramentasPressionada = False
    Case "Impressoras"
        Call MnuConImpressora_Click
        BarraFerramentasPressionada = False
    Case "Sobre"
        FrmSobre.Show 1
    Case "Sair"
        Unload Me
    End Select
    
End Sub

Private Sub MDIForm_Activate()
   
    ' Carrega formulario para entrada da senha do usuário
    If func_cd_operador = 0 Then
       func_cd_operador = 9999
       'XGB_PQUALIFIN = True
       FrmSenha.Show 1
       'FrmMultiEmpresa.Show 1
    End If

End Sub

Private Sub mniFeriados_Click()
 
    If Not FunVerifica_Permissao("TelaManFeriados", "1") Then
        Exit Sub
    End If
    
    TelaManFeriados.Show
    If FunLocalizaJanelasAtivas("TelaManFeriados") <> -1 Then
        TelaManFeriados.SetFocus
    End If
End Sub

Private Sub MnuAjuda_Click()
  
  FrmSobre.Show vbModal
  
End Sub

Private Sub MnuAnaliticoDiario_Click()
    
  If Not FunVerifica_Permissao("FrmRelPrevisaoRecebimento", "1") Then
       Exit Sub
  End If
  
  XOrigem = "Analitico Diario"
  FrmRelPrevisaoRecebimento.Show
  If FunLocalizaJanelasAtivas("FrmRelPrevisaoRecebimento") <> -1 Then
    FrmRelPrevisaoRecebimento.SetFocus
  End If
    
End Sub

Private Sub MnuAnaliticoMensal_Click()

  If Not FunVerifica_Permissao("FrmRelPrevisaoRecebimento", "1") Then
       Exit Sub
  End If
  
  XOrigem = "Analitico Mensal"
  FrmRelPrevisaoRecebimento.Show
  If FunLocalizaJanelasAtivas("FrmRelPrevisaoRecebimento") <> -1 Then
    FrmRelPrevisaoRecebimento.SetFocus
  End If

End Sub

Private Sub MnuAnexoFluxoRecebido_Click()
    FrmRelAnexoFluxoRecebiveisItau.Show
End Sub

Private Sub MnuAtualizar_Click()
    Conexao.Close
    ConexaoRelatorio.Close
    
    Set Conexao = Nothing
    Set ConexaoRelatorio = Nothing
        
    Call Shell(Mid(App.Path, 1, (Len(App.Path) - Len("\QualiCapi"))) & "\Atualizador\Atualizador.exe QualiCapi """, vbNormalFocus)
    End
End Sub

Private Sub MnuAtualizarExpBradesco_Click()
    Conexao.Close
    ConexaoRelatorio.Close
    
    Set Conexao = Nothing
    Set ConexaoRelatorio = Nothing
        
    Call Shell(Mid(App.Path, 1, (Len(App.Path) - Len("\QualiCapi"))) & "\Atualizador\Atualizador.exe QualiCapi ExpBradesco", vbNormalFocus)
    End
End Sub

Private Sub MnuAtualizarExpHSBC_Click()
    Conexao.Close
    ConexaoRelatorio.Close
    
    Set Conexao = Nothing
    Set ConexaoRelatorio = Nothing
        
    Call Shell(Mid(App.Path, 1, (Len(App.Path) - Len("\QualiCapi"))) & "\Atualizador\Atualizador.exe QualiCapi ExpHSBC", vbNormalFocus)
    End
End Sub

Private Sub MnuAtualizarExpItau_Click()
    Conexao.Close
    ConexaoRelatorio.Close
    
    Set Conexao = Nothing
    Set ConexaoRelatorio = Nothing
    
    Call Shell(Mid(App.Path, 1, (Len(App.Path) - Len("\QualiCapi"))) & "\Atualizador\Atualizador.exe QualiCapi ExpItau", vbNormalFocus)
    End
End Sub

Private Sub MnuAtualizarExpReal_Click()
    Conexao.Close
    ConexaoRelatorio.Close
    
    Set Conexao = Nothing
    Set ConexaoRelatorio = Nothing
    
    Call Shell(Mid(App.Path, 1, (Len(App.Path) - Len("\QualiCapi"))) & "\Atualizador\Atualizador.exe QualiCapi ExpReal", vbNormalFocus)
    End
End Sub

Private Sub MnuAtualizarExpSantander_Click()
    Conexao.Close
    ConexaoRelatorio.Close
    
    Set Conexao = Nothing
    Set ConexaoRelatorio = Nothing
    
    Call Shell(Mid(App.Path, 1, (Len(App.Path) - Len("\QualiCapi"))) & "\Atualizador\Atualizador.exe QualiCapi ExpSantander", vbNormalFocus)
    End
End Sub

Private Sub MnuAtualizarExpSudameris_Click()
    Conexao.Close
    ConexaoRelatorio.Close
    
    Set Conexao = Nothing
    Set ConexaoRelatorio = Nothing
    
    Call Shell(Mid(App.Path, 1, (Len(App.Path) - Len("\QualiCapi"))) & "\Atualizador\Atualizador.exe QualiCapi ExpSudameris", vbNormalFocus)
    End
End Sub

Private Sub MnuAtualizarExpUnibanco_Click()
    Conexao.Close
    ConexaoRelatorio.Close
    
    Call Shell(Mid(App.Path, 1, (Len(App.Path) - Len("\QualiCapi"))) & "\Atualizador\Atualizador.exe QualiCapi ExpUnibanco", vbNormalFocus)
    End
End Sub

Private Sub MnuConBanco_Click()
  FrmConfBanco.Show vbModal
End Sub

Private Sub MnuConImportacao_Click()
    FrmImporta.Show
End Sub

Private Sub MnuExportacaoBradesco_Click()
    If Not FunVerifica_Permissao("ExportacaoBradesco", "1") Then
       Exit Sub
    End If

    Shell App.Path + "\ExpBradesco\ExpBradesco.exe" + Str(PCodEmpresa) + "," + Trim(Str(func_cd_operador)), vbNormalFocus
End Sub

Private Sub MnuExportacaoCaixa_Click()

    If Not FunVerifica_Permissao("ExportacaoCaixa", "1") Then
       Exit Sub
    End If
  
  Shell App.Path + "\ExpCaixa\ExpCaixa.exe" + Str(PCodEmpresa) + "," + Trim(Str(func_cd_operador)), vbNormalFocus
End Sub

Private Sub MnuExportacaoDimob_Click()
  
  If Not FunVerifica_Permissao("ExportacaoDimob", "1") Then
       Exit Sub
  End If
  
  TelaExportacaoDimob.Show
  If FunLocalizaJanelasAtivas("TelaExportacaoDimob") <> -1 Then
    TelaExportacaoDimob.SetFocus
  End If
  
End Sub

Private Sub MnuexportacaoHSBC_Click()
  
  If Not FunVerifica_Permissao("ExportacaoHsbc", "1") Then
       Exit Sub
  End If
  
  Shell App.Path + "\ExpHSBC\ExpHSBC.exe" + Str(PCodEmpresa) + "," + Trim(Str(func_cd_operador)), vbNormalFocus
End Sub

Private Sub MnuExportacaoItau_Click()
  If Not FunVerifica_Permissao("ExportacaoItau", "1") Then
    Exit Sub
  End If
  Shell App.Path + "\ExpItau\ExpItau.exe" + Str(PCodEmpresa) + "," + Trim(Str(func_cd_operador)), vbNormalFocus
End Sub

Private Sub MnuExportacaoReal_Click()
  
  If Not FunVerifica_Permissao("ExportacaoReal", "1") Then
       Exit Sub
  End If
  
  Shell App.Path + "\ExpReal\ExpReal.exe" + Str(PCodEmpresa) + "," + Trim(Str(func_cd_operador)), vbNormalFocus
End Sub

Private Sub MnuExportacaoSantander_Click()

  If Not FunVerifica_Permissao("ExportacaoSantander", "1") Then
       Exit Sub
  End If

  Shell App.Path + "\ExpSantander\ExpSantander.exe" + Str(PCodEmpresa) + "," + Trim(Str(func_cd_operador)), vbNormalFocus
End Sub

Private Sub MnuExportacaoSudameris_Click()
  
  If Not FunVerifica_Permissao("ExportacaoSudameris", "1") Then
       Exit Sub
  End If
  
  Shell App.Path + "\ExpSudameris\ExpSudameris.exe" + Str(PCodEmpresa) + "," + Trim(Str(func_cd_operador)), vbNormalFocus
End Sub

Private Sub MnuExportacaoUnibanco_Click()
  
  If Not FunVerifica_Permissao("ExportacaoUnibanco", "1") Then
       Exit Sub
  End If
  
  Shell App.Path + "\ExpUnibanco\ExpUnibanco.exe" + Str(PCodEmpresa) + "," + Trim(Str(func_cd_operador)), vbNormalFocus
End Sub

Private Sub MnuExportEmpresas_Click()

    FrmExportEmpresas.Show
    
    If FunLocalizaJanelasAtivas("FrmExportEmpresas") <> -1 Then
        FrmExportEmpresas.SetFocus
    End If
End Sub

Private Sub MnuFinAtualizarQualifin_Click()
    
    If Not FunVerifica_Permissao("FrmConAlteraQualifin", "1") Then
        Exit Sub
    End If
    
    If XGB_PQUALIFIN = True Then
        XFormulario = ""
        FrmFinAlteraQualifin.Show vbModal
        If FunLocalizaJanelasAtivas("FrmConAlteraQualifin") <> -1 Then
            FrmFinAlteraQualifin.SetFocus
        End If
    Else
        MsgBox "O sistema financeiro não está integrado. Contate seu administrador.", 16, "AVISO"
    End If
End Sub

Private Sub MnuConForm_Click()
    
    If Not FunVerifica_Permissao("FrmConConfigDoc", "1") Then
       Exit Sub
    End If
    
    FrmConConfigDoc.Show
    If FunLocalizaJanelasAtivas("FrmConConfigDoc") <> -1 Then
        FrmConConfigDoc.SetFocus
    End If
    MdiPrincipal.MousePointer = vbDefault
End Sub

Private Sub MnuFinCotacoes_Click()

  If Not FunVerifica_Permissao("TelaCotacoes", "1") Then
       Exit Sub
  End If
  
  XFormulario = ""
  TelaCotacoes.Show vbModal
  If FunLocalizaJanelasAtivas("TelaCotacoes") <> -1 Then
    TelaCotacoes.SetFocus
  End If
  
End Sub

Private Sub mnuFluxoRecebiboReal_Click()
    FrmRelAnexoFluxoRecebiveisReal.Show
End Sub

Private Sub MnuFluxoRecebidoSantander_Click()
    FrmRelAnexoFluxoRecebiveisSantander.Show
End Sub

Private Sub MnuImportEmpresas_Click()
    FrmImportEmpresas.Show
    
    If FunLocalizaJanelasAtivas("FrmImportEmpresas") <> -1 Then
        FrmImportEmpresas.SetFocus
    End If
End Sub

Private Sub MnuManClasseCnae_Click_Click()
   
   If Not FunVerifica_Permissao("TelaIdentificadores", "1") Then
        Exit Sub
    End If
    
    XGT_CNAE = "Classe"
    
    MdiPrincipal.MousePointer = vbHourglass
    
    TelaCnae.Show
    
    If FunLocalizaJanelasAtivas("TelaCnae") <> -1 Then
        TelaCnae.SetFocus
    End If
    
    MdiPrincipal.MousePointer = vbDefault

End Sub

Private Sub MnuManDivisaoCnae_Click_Click()

   If Not FunVerifica_Permissao("TelaIdentificadores", "1") Then
        Exit Sub
    End If
    
    XGT_CNAE = "Divisao"
    
    MdiPrincipal.MousePointer = vbHourglass
    
    TelaCnae.Show
    
    If FunLocalizaJanelasAtivas("TelaCnae") <> -1 Then
        TelaCnae.SetFocus
    End If
    
    MdiPrincipal.MousePointer = vbDefault
End Sub

Private Sub MnuManGrupoCnae_Click_Click()
   
   If Not FunVerifica_Permissao("TelaIdentificadores", "1") Then
        Exit Sub
    End If
    
    XGT_CNAE = "Grupo"
    
    MdiPrincipal.MousePointer = vbHourglass
    
    TelaCnae.Show
    
    If FunLocalizaJanelasAtivas("TelaCnae") <> -1 Then
        TelaCnae.SetFocus
    End If
    
    MdiPrincipal.MousePointer = vbDefault

End Sub

Private Sub MnuManIdentificadores_Click_Click()
    
    If Not FunVerifica_Permissao("TelaIdentificadores", "1") Then
        Exit Sub
    End If

    MdiPrincipal.MousePointer = vbHourglass
    TelaIdentificadores.Show
    If FunLocalizaJanelasAtivas("TelaIdentificadores") <> -1 Then
        TelaIdentificadores.SetFocus
    End If
    MdiPrincipal.MousePointer = vbDefault
End Sub

Private Sub MnuManSecaoCnae_Click_Click()
    
    If Not FunVerifica_Permissao("TelaIdentificadores", "1") Then
        Exit Sub
    End If
    
    XGT_CNAE = "Secao"
    
    MdiPrincipal.MousePointer = vbHourglass
    
    TelaCnae.Show
    
    If FunLocalizaJanelasAtivas("TelaCnae") <> -1 Then
        TelaCnae.SetFocus
    End If
    
    MdiPrincipal.MousePointer = vbDefault
End Sub

Private Sub MnuManSubClasseCnae_Click_Click()
   If Not FunVerifica_Permissao("TelaIdentificadores", "1") Then
        Exit Sub
    End If
    
    XGT_CNAE = "SubClasse"
    
    MdiPrincipal.MousePointer = vbHourglass
    
    TelaCnae.Show
    
    If FunLocalizaJanelasAtivas("TelaCnae") <> -1 Then
        TelaCnae.SetFocus
    End If
    
    MdiPrincipal.MousePointer = vbDefault
End Sub

Private Sub MnuManTabBancos_Click()
    
    If Not FunVerifica_Permissao("TelaBancos", "1") Then
        Exit Sub
    End If
    
    TelaBancos.Show
    If FunLocalizaJanelasAtivas("TelaBancos") <> -1 Then
        TelaBancos.SetFocus
    End If
End Sub

Private Sub MnuManTabCartas_Click()
  
  If Not FunVerifica_Permissao("TelaManTabCartas", "1") Then
      Exit Sub
  End If
  
  TelaManTabCartas.Show
  If FunLocalizaJanelasAtivas("TelaManTabCartas") <> -1 Then
    TelaManTabCartas.SetFocus
  End If
  
End Sub

Private Sub MnuManTabCCorrentes_Click()
    
    If Not FunVerifica_Permissao("TelaContaCorrente", "1") Then
        Exit Sub
    End If

    MdiPrincipal.MousePointer = vbHourglass
    TelaContaCorrente.Show
    If FunLocalizaJanelasAtivas("TelaContaCorrente") <> -1 Then
        TelaContaCorrente.SetFocus
    End If
    MdiPrincipal.MousePointer = vbDefault
End Sub

Private Sub MnuPosicaoComercializacao_Click()
    If Not FunVerifica_Permissao("FrmRelRelacaoClientes", "1") Then
         Exit Sub
    End If
      
    FrmRelPosComercializacaoItau.Show
    If FunLocalizaJanelasAtivas("FrmRelRelacaoClientes") <> -1 Then
        FrmRelPosComercializacaoItau.SetFocus
    End If

End Sub

Private Sub MnuPosicaoComercializacaoBradesco_Click()
    If Not FunVerifica_Permissao("FrmRelRelacaoClientes", "1") Then
        Exit Sub
    End If
      
    FrmRelPosComercializacaoBradesco.Show
    If FunLocalizaJanelasAtivas("FrmRelRelacaoClientes") <> -1 Then
        FrmRelPosComercializacaoBradesco.SetFocus
    End If

End Sub

Private Sub MnuPosicaoComercializacaoSantander_Click()
    If Not FunVerifica_Permissao("FrmRelRelacaoClientes", "1") Then
        Exit Sub
    End If
      
    FrmRelPosComercializacaoSantander.Show
    If FunLocalizaJanelasAtivas("FrmRelRelacaoClientes") <> -1 Then
        FrmRelPosComercializacaoSantander.SetFocus
    End If
End Sub

Private Sub MnuRealMapaContabil_Click()
  
  FrmRelMapaContabil.Show
  If FunLocalizaJanelasAtivas("FrmRelMapaContabil") <> -1 Then
    FrmRelMapaContabil.SetFocus
  End If
   
End Sub

Private Sub MnuRecibos_Click()
  If Not FunVerifica_Permissao("FrmRelRecibos", "1") Then
       Exit Sub
  End If
  
FrmRelRecibos.Show
End Sub

Private Sub MnuRelacaoCompradores_Click()
    If Not FunVerifica_Permissao("FrmRelRelacaoClientes", "1") Then
           Exit Sub
      End If
      
      FrmRelCompradoresCadastrados.Show
      If FunLocalizaJanelasAtivas("FrmRelRelacaoClientes") <> -1 Then
        FrmRelCompradoresCadastrados.SetFocus
      End If
End Sub

Private Sub MnuRelAssistenciaTecnica_Click()

 ' If Not FunVerifica_Permissao("FrmRelAssistenciaTecnica", "1") Then
   '    Exit Sub
'  End If
  
  FrmRelAssistenciaTecnica.Show
  If FunLocalizaJanelasAtivas("FrmRelAssistenciaTecnica") <> -1 Then
    FrmRelAssistenciaTecnica.SetFocus
  End If
End Sub

Private Sub mnuRelCartCobranca_Click()
  
  If Not FunVerifica_Permissao("FrmRelCartasCobranca", "1") Then
       Exit Sub
  End If
  
  FrmRelCartasCobranca.Show
  If FunLocalizaJanelasAtivas("FrmRelCartasCobranca") <> -1 Then
    FrmRelCartasCobranca.SetFocus
  End If
  
End Sub

Private Sub mnuRelCartEspeciais_Click()
  
  If Not FunVerifica_Permissao("FrmRelCartas", "1") Then
       Exit Sub
  End If
  
  FrmRelCartas.Show
  If FunLocalizaJanelasAtivas("FrmRelCartas") <> -1 Then
    FrmRelCartas.SetFocus
  End If
  
End Sub

Private Sub MnuRelCOAF_Click()
    FrmRelCOAF.Show
  
  If FunLocalizaJanelasAtivas("FrmRelCOAF") <> -1 Then
    FrmRelCOAF.SetFocus
  End If
End Sub

Private Sub MnuRelCompradores_Click()
    
    If Not FunVerifica_Permissao("FrmRelRelacaoClientes", "1") Then
       Exit Sub
    End If
    
    FrmRelCompradores.Show
    If FunLocalizaJanelasAtivas("FrmRelCompradores") <> -1 Then
        FrmRelCompradores.SetFocus
    End If
End Sub

Private Sub MnuRelContratosComSeguro_Click()
  
  If Not FunVerifica_Permissao("FrmRelContratosComSeguro", "1") Then
       Exit Sub
  End If
  
  FrmRelContratosComSeguro.Show
  If FunLocalizaJanelasAtivas("FrmRelContratosComSeguro") <> -1 Then
    FrmRelContratosComSeguro.SetFocus
  End If
  
End Sub

Private Sub MnuRelDemonstrativoPorCliente_Click()
    
  If Not FunVerifica_Permissao("FrmRelDemonstrativoCliente", "1") Then
       Exit Sub
  End If
  
  FrmRelDemonstrativoCliente.Show
  If FunLocalizaJanelasAtivas("FrmRelDemonstrativoCliente") <> -1 Then
    FrmRelDemonstrativoCliente.SetFocus
  End If
  
End Sub

Private Sub MnuRelEtiquetas_Click()
  
  If Not FunVerifica_Permissao("FrmRelEtiquetas", "1") Then
    Exit Sub
  End If
  
  FrmRelEtiquetas.Show
  
  If FunLocalizaJanelasAtivas("FrmRelEtiquetas") <> -1 Then
    FrmRelExtratoContratoAnalitico.SetFocus
  End If
End Sub

Private Sub MnuRelExtratoContratoAnalitico_Click()
  
  If Not FunVerifica_Permissao("FrmRelExtratoContratoAnalitico", "1") Then
       Exit Sub
  End If
  
  XOrigem = "Analitico"
  FrmRelExtratoContratoAnalitico.Show
  If FunLocalizaJanelasAtivas("FrmRelExtratoContratoAnalitico") <> -1 Then
    FrmRelExtratoContratoAnalitico.SetFocus
  End If
  
End Sub

Private Sub MnuRelExtratoContratoAnaliticoCliente_Click()
  
 ' If Not FunVerifica_Permissao("FrmRelExtratoContratoAnaliticoCliente", "1") Then
 '      Exit Sub
 ' End If
  
  XOrigem = "Analitico"
  FrmRelExtratoContratoAnaliticoCliente.Show
  If FunLocalizaJanelasAtivas("FrmRelExtratoContratoAnaliticoCliente") <> -1 Then
    FrmRelExtratoContratoAnaliticoCliente.SetFocus
  End If

End Sub

Private Sub MnuRelExtratoContratoPagamentos_Click()
  
  If Not FunVerifica_Permissao("FrmRelExtratoContratoPgtosRealizados", "1") Then
       Exit Sub
  End If
  
  XOrigem = "Pagamentos"
  FrmRelExtratoContratoPgtosRealizados.Show
  If FunLocalizaJanelasAtivas("FrmRelExtratoContratoPgtosRealizados") <> -1 Then
    FrmRelExtratoContratoPgtosRealizados.SetFocus
  End If
  
End Sub

Private Sub MnuRelExtratoContratoReal_Click()

  If Not FunVerifica_Permissao("FrmRelExtratoContratoAnalitico", "1") Then
       Exit Sub
  End If
  
  XOrigem = "Analitico"
  FrmRelExtratoContratoAnaliticoReal.Show
  If FunLocalizaJanelasAtivas("FrmRelExtratoContratoAnaliticoReal") <> -1 Then
    FrmRelExtratoContratoAnaliticoReal.SetFocus
  End If
  
End Sub

Private Sub MnuRelExtratoContratoResumoDebitos_Click()
  
  If Not FunVerifica_Permissao("FrmRelExtratoContratoResumoDebitos", "1") Then
       Exit Sub
  End If
  
  XOrigem = "Debitos"
  FrmRelExtratoContratoResumoDebitos.Show
  If FunLocalizaJanelasAtivas("FrmRelExtratoContratoResumoDebitos") <> -1 Then
    FrmRelExtratoContratoResumoDebitos.SetFocus
  End If
  
End Sub

Private Sub MnuRelExtratoContratoSintetico_Click()

  If Not FunVerifica_Permissao("FrmRelExtratoContratoSintetico", "1") Then
       Exit Sub
  End If
  
  FrmRelExtratoContratoSintetico.Show
  If FunLocalizaJanelasAtivas("FrmRelExtratoContratoSintetico") <> -1 Then
    FrmRelExtratoContratoSintetico.SetFocus
  End If

End Sub

Private Sub MnuRelFluxoGeralPagamento_Click()
  
  If Not FunVerifica_Permissao("FrmRelFluxoGeral", "1") Then
       Exit Sub
  End If
  
  XOrigem = "Pagamento"
  FrmRelFluxoGeral.Show
  If FunLocalizaJanelasAtivas("FrmRelFluxoGeral") <> -1 Then
    FrmRelFluxoGeral.SetFocus
  End If
  
End Sub

Private Sub MnuRelfluxoGeralVencimento_Click()
  
  If Not FunVerifica_Permissao("FrmRelFluxoGeral", "1") Then
       Exit Sub
  End If
  
  XOrigem = "Vencimento"
  FrmRelFluxoGeral.Show
  If FunLocalizaJanelasAtivas("FrmRelFluxoGeral") <> -1 Then
    FrmRelFluxoGeral.SetFocus
  End If
  
End Sub

Private Sub MnuRelImoveisQuitados_Click()
  
  If Not FunVerifica_Permissao("FrmRelRelacaoImoveis", "1") Then
       Exit Sub
  End If
  
  FrmRelRelacaoImoveis.Show
  If FunLocalizaJanelasAtivas("FrmRelRelacaoImoveis") <> -1 Then
    FrmRelRelacaoImoveis.SetFocus
  End If
  
End Sub

Private Sub MnuRelItiv_Click()
  FrmRelItiv.Show
  
  If FunLocalizaJanelasAtivas("FrmRelItiv") <> -1 Then
    FrmRelItiv.SetFocus
  End If
  
End Sub

Private Sub MnuRelListagemCotacao_Click()
  
  If Not FunVerifica_Permissao("FrmRelCotacoes", "1") Then
      Exit Sub
  End If
  
  FrmRelCotacoes.Show
  If FunLocalizaJanelasAtivas("FrmRelCotacoes") <> -1 Then
    FrmRelCotacoes.SetFocus
  End If
  
End Sub

Private Sub MnuRelListagemProRata_Click()
  FrmRelCotacoesProRata.Show
  If FunLocalizaJanelasAtivas("FrmRelCotacoesProRata") <> -1 Then
    FrmRelCotacoesProRata.SetFocus
  End If
  
End Sub

Private Sub MnuRelMapaGeral_Click()
    
  If Not FunVerifica_Permissao("FrmRelMapaGeral", "1") Then
       Exit Sub
  End If
  
  FrmRelMapaGeral.Show
  If FunLocalizaJanelasAtivas("FrmRelMapaGeral") <> -1 Then
    FrmRelMapaGeral.SetFocus
  End If
   
End Sub

Private Sub MnuRelPorNatureza_Click()
  
  If Not FunVerifica_Permissao("FrmRelTitulosLiquidadosNatureza", "1") Then
    Exit Sub
  End If
  
  XOrigem = "PorNatureza"
  FrmRelTitulosLiquidadosNatureza.Show
  If FunLocalizaJanelasAtivas("FrmRelTitulosLiquidadosNatureza") <> -1 Then
    FrmRelTitulosLiquidadosNatureza.SetFocus
  End If
  
End Sub

Private Sub MnuRelPromissorias_Click()
  
  If Not FunVerifica_Permissao("FrmPromissoria", "1") Then
       Exit Sub
  End If
  
  FrmPromissoria.Show
  
  If FunLocalizaJanelasAtivas("FrmPromissoria") <> -1 Then
    FrmPromissoria.SetFocus
  End If
  
End Sub

Private Sub mnuRelReapasse_Click()
  
  If Not FunVerifica_Permissao("FrmRelRepasse", "1") Then
       Exit Sub
  End If
  
  FrmRelRepasse.Show
  If FunLocalizaJanelasAtivas("FrmRelRepasse") <> -1 Then
    FrmRelRepasse.SetFocus
  End If
  
End Sub

Private Sub MnuRelSegurosAPagar_Click()
  
  If Not FunVerifica_Permissao("FrmRelSegurosAPagar", "1") Then
       Exit Sub
  End If
  
  FrmRelSegurosAPagar.Show
  If FunLocalizaJanelasAtivas("FrmRelSegurosAPagar") <> -1 Then
    FrmRelSegurosAPagar.SetFocus
  End If
  
End Sub

Private Sub MnuRelSegurosPagos_Click()
  
  If Not FunVerifica_Permissao("FrmRelSegurosPagos", "1") Then
       Exit Sub
  End If
  
  FrmRelSegurosPagos.Show
  If FunLocalizaJanelasAtivas("FrmRelSegurosPagos") <> -1 Then
    FrmRelSegurosPagos.SetFocus
  End If
  
End Sub

Private Sub MnuRelSituacaoContratos_Click()
  XOrigem = "SituacaoContratos"
  FrmRelSituacaoImoveis.Show
  If FunLocalizaJanelasAtivas("FrmRelSituacaoContratos") <> -1 Then
    FrmRelSegurosPagos.SetFocus
  End If
   
End Sub

Private Sub MnuRelTitulosLiquidadosAnalitico_Click()
  
  If Not FunVerifica_Permissao("FrmRelTitulosLiquidados", "1") Then
       Exit Sub
  End If
  
  XOrigem = "Analitico"
  FrmRelTitulosLiquidados.Show
  If FunLocalizaJanelasAtivas("FrmRelTitulosLiquidados") <> -1 Then
    FrmRelTitulosLiquidados.SetFocus
  End If
  
End Sub

Private Sub MnuRelTitulosLiquidadosPIS_Click()
    
    If Not FunVerifica_Permissao("FrmRelTitulosLiquidadosPis", "1") Then
        Exit Sub
    End If
    
    FrmRelTitulosLiquidadosPIS.Show
    If FunLocalizaJanelasAtivas("FrmRelTitulosLiquidadosPis") <> -1 Then
        FrmRelTitulosLiquidadosPIS.SetFocus
    End If
    
End Sub

Private Sub MnuRelTitulosLiquidadosSintetico_Click()
  
  If Not FunVerifica_Permissao("FrmRelTitulosLiquidados", "1") Then
       Exit Sub
  End If
  
  XOrigem = "Sintetico"
  FrmRelTitulosLiquidados.Show
  If FunLocalizaJanelasAtivas("FrmRelTitulosLiquidados") <> -1 Then
    FrmRelTitulosLiquidados.SetFocus
  End If
  
End Sub

Private Sub MnuRelTitulosVencidosAnalitico_Click()
    
  If Not FunVerifica_Permissao("FrmRelTitulosVencidos", "1") Then
       Exit Sub
  End If
  
  FrmRelTitulosVencidos.Show
  If FunLocalizaJanelasAtivas("FrmRelTitulosVencidos") <> -1 Then
    FrmRelTitulosVencidos.SetFocus
  End If
    
End Sub

Private Sub MnuRelTitulosVencidosPorCliente_Click()
  
  If Not FunVerifica_Permissao("FrmRelTitulosVencidosPorCliente", "1") Then
       Exit Sub
  End If
  
  FrmRelTitulosVencidosPorCliente.Show
  If FunLocalizaJanelasAtivas("FrmRelTitulosVencidosPorCliente") <> -1 Then
    FrmRelTitulosVencidosPorCliente.SetFocus
  End If
   
End Sub

Private Sub MnuResumoFGeral_Click()
  
  If Not FunVerifica_Permissao("FrmRelResumoGeral", "1") Then
       Exit Sub
  End If
  FrmRelResumoGeral.Show
  If FunLocalizaJanelasAtivas("FrmRelResumoGeral") <> -1 Then
    FrmRelResumoGeral.SetFocus
  End If
  
End Sub

Private Sub MnuResumoMensal_Click()
     
  If Not FunVerifica_Permissao("FrmRelResumoMensal", "1") Then
       Exit Sub
  End If
  FrmRelResumoMensal.Show
  If FunLocalizaJanelasAtivas("FrmRelResumoMensal") <> -1 Then
    FrmRelResumoMensal.SetFocus
  End If
  
End Sub

Private Sub MnuSinteticoDiario_Click()

  If Not FunVerifica_Permissao("FrmRelPrevisaoRecebimento", "1") Then
       Exit Sub
  End If
  XOrigem = "Sintetico Diario"
  FrmRelPrevisaoRecebimento.Show
  If FunLocalizaJanelasAtivas("FrmRelPrevisaoRecebimento") <> -1 Then
    FrmRelPrevisaoRecebimento.SetFocus
  End If

End Sub

Private Sub MnuSinteticoMensal_Click()

  If Not FunVerifica_Permissao("FrmRelPrevisaoRecebimento", "1") Then
      Exit Sub
  End If
    
  XOrigem = "Sintetico Mensal"
  FrmRelPrevisaoRecebimento.Show
  If FunLocalizaJanelasAtivas("FrmRelPrevisaoRecebimento") <> -1 Then
    FrmRelPrevisaoRecebimento.SetFocus
  End If

End Sub

Private Sub MnuAtiva_Click()
   FrmMultiEmpresa.Show 1
End Sub

Private Sub MDIForm_Load()
    Dim Diretorio As String 'Diretório onde os sistemas QualiPro estão instalados. Ex: C:\QualiAdmFin
    
    ReDim VetorJanelas(0) As Variant
    XGT_NomeAplic = "ADMFIN"
    XGT_IdAplic = "QUALIADMFIN-NIFMDAILAUQ"
  
    If UCase(PCliente) = "EBISA" Then
        pTipoCopia = "Multi"
        XGB_PQUALIFIN = True
        XGB_PCAPI = True
        XGB_PQUALICONT = True
    ElseIf UCase(PCliente) = "RJ" Then
        pTipoCopia = "Multi"
        XGB_PQUALIFIN = True
        XGB_PQUALIEST = True
        XGB_PCOMPRAS = True
        XGB_PCAPI = True
    Else
        subTestaProtecao XGT_NomeAplic, XGT_IdAplic, XGT_DirPadrao
        If UCase(PCliente) = "COSTA ANDRADE" Then XGB_PQUALICONT = True
    End If
    
    XGT_NomeAplic = "CAPI"
    Diretorio = Mid(App.Path, 1, (Len(App.Path) - Len("\QualiCapi")))
    
    If FunVerificaVersao = False Then
        Dim res As VbMsgBoxResult
        
        res = MsgBox("Máquina desatualizada! Deseja atualizar o Qualicapi nesta máquina?", VbMsgBoxStyle.vbYesNo + vbQuestion, "Aviso")
        
        'Encerra a conexão com o banco
        Conexao.Close
        ConexaoRelatorio.Close
        
        If (res = vbYes) Then
            Call Shell(Diretorio & "\Atualizador\Atualizador.exe QualiCapi """, vbNormalFocus)
            End
        Else
            MsgBox "A aplicação não será executada até ser atualiazada!", VbMsgBoxStyle.vbInformation, "Aviso"
            End
        End If
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
  
    MdiPrincipal.BarraStatus.Panels(2).Text = PFantasia
    MdiPrincipal.Label1.Caption = "R " & Format(App.Major, "#0") & "." & Format(App.Revision, "00")


    'Verifica as proteções para os bancos e habilita as opções de acordo com os resultados.
    'Banco Itau
    If UCase(PCliente) = "COSTA ANDRADE" Or UCase(PCliente) = "EBISA" Or UCase(PCliente) = "FRANISA" Or UCase(PCliente) = "PLENA" Then
        MnuExportacaoItau.Enabled = True
    Else
        If funTemProtecao(App.Path + "\ExpItau", "EXPITAU", "EXPITAU-UATIPXE") Then
            MnuExportacaoItau.Enabled = True
        Else
            MnuExportacaoItau.Enabled = False
        End If
    End If
 
    'Banco HSBC
    If funTemProtecao(App.Path + "\ExpHSBC", "EXPHSBC", "EXPHSBC-CBSHPXE") Then
        MnuexportacaoHSBC.Enabled = True
    Else
        MnuexportacaoHSBC.Enabled = False
    End If
  
    'Banco Real
    'If funTemProtecao(App.Path + "\ExpReal", "EXPREAL", "EXPREAL-LAERPXE") Then
        MnuExportacaoReal.Enabled = True
    'Else
    '  MnuExportacaoReal.Enabled = False
    'End If

    'Banco Sudameris
    If funTemProtecao(App.Path + "\ExpSudameris", "EXPSUDAMERIS", "EXPSUDAMERIS-SIREMADUSPXE") Then
        MnuExportacaoSudameris.Enabled = True
    Else
        MnuExportacaoSudameris.Enabled = False
    End If

    'Banco Unibanco
    If funTemProtecao(App.Path + "\ExpUnibanco", "EXPUNIBANCO", "EXPUNIBANCO-OCNABINUPXE") Then
        MnuExportacaoUnibanco.Enabled = True
    Else
        MnuExportacaoUnibanco.Enabled = False
    End If
  
    '***************************INICIO - PEGA O ID DOS TIPOS DE DESCONTOS******************************
    Dim XLO_DESCONTOS As New ADODB.Recordset
    
    SubQOpenRecordset XLO_DESCONTOS, "SELECT deac_cd_descacresc, deac_tx_Descricao FROM DescontosAcrescimos WHERE deac_tx_Descricao IN ('Multa por Atraso - Cliente', 'Juros por Atraso - Cliente', 'Deságio - Cliente', 'Descontos - Cliente', 'Seguro - Cliente', 'Outros - Cliente')", Estatico
    
    If Not XLO_DESCONTOS.EOF Then '07/05/09 - Patrícia
        XLO_DESCONTOS.MoveFirst
        While Not XLO_DESCONTOS.EOF
        
            If (XLO_DESCONTOS!deac_tx_Descricao = "Multa por Atraso - Cliente") Then
                XGI_CODMULTA = XLO_DESCONTOS!deac_cd_descacresc
            ElseIf (XLO_DESCONTOS!deac_tx_Descricao = "Juros por Atraso - Cliente") Then
                XGI_CODJUROS = XLO_DESCONTOS!deac_cd_descacresc
            ElseIf (XLO_DESCONTOS!deac_tx_Descricao = "Deságio - Cliente") Then
                XGI_CODDESAGIO = XLO_DESCONTOS!deac_cd_descacresc
            ElseIf (XLO_DESCONTOS!deac_tx_Descricao = "Descontos - Cliente") Then
                XGI_CODDESCONTO = XLO_DESCONTOS!deac_cd_descacresc
            ElseIf (XLO_DESCONTOS!deac_tx_Descricao = "Outros - Cliente") Then
                XGI_CODSEGURO = XLO_DESCONTOS!deac_cd_descacresc
            ElseIf (XLO_DESCONTOS!deac_tx_Descricao = "Seguro - Cliente") Then
                XGI_CODOUTROS = XLO_DESCONTOS!deac_cd_descacresc
            End If
        
            XLO_DESCONTOS.MoveNext
        Wend
    End If
    XLO_DESCONTOS.Close
    Set XLO_DESCONTOS = Nothing
    '***************************FIM - PEGA O ID DOS TIPOS DE DESCONTOS******************************
    
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    Finaliza = False
    FrmFinalizar.Show 1
    Cancel = Finaliza
    
End Sub

Private Sub MnuFinPagtoMultiplo_Click()
    
    If Not FunVerifica_Permissao("TelaFinPagtoMultiplos", "1") Then
        Exit Sub
    End If
    
    TelaFinPagtoMultiplos.Show
    If FunLocalizaJanelasAtivas("TelaFinPagtoMultiplos") <> -1 Then
        TelaFinPagtoMultiplos.SetFocus
    End If
    
End Sub

Private Sub MnuFinPagtoUnico_Click()
    
    If Not FunVerifica_Permissao("FrmFinPagamentos", "1") Then
        Exit Sub
    End If
    
    XFormulario = ""
    FrmFinPagamentos.Show vbModal
    If FunLocalizaJanelasAtivas("FrmFinPagamentos") <> -1 Then
        FrmFinPagamentos.SetFocus
    End If
    
End Sub

Private Sub MnuFinPlanos_Click()
    
    If Not FunVerifica_Permissao("FrmFinPlanos", "1") Then
        Exit Sub
    End If
    
    FrmFinPlanos.Show 1
    If FunLocalizaJanelasAtivas("FrmFinPlanos") <> -1 Then
        FrmFinPlanos.SetFocus
    End If
End Sub

Private Sub MnuFinTitulos_Click()
    
    If Not FunVerifica_Permissao("TelaFinTitulos", "1") Then
        Exit Sub
    End If
    
    TelaFinTitulos.Show
    If FunLocalizaJanelasAtivas("TelaFinTitulos") <> -1 Then
        TelaFinTitulos.SetFocus
    End If
    
End Sub

Private Sub MnuManCliPessoaFisica_Click()
    
    If Not FunVerifica_Permissao("TelaFornecedores", "2") Then
        Exit Sub
    End If
    
    FornClie = "C"
    Pessoa = "F"
    TelaFornecedores.Show
    If FunLocalizaJanelasAtivas("TelaFornecedores") <> -1 Then
        TelaFornecedores.SetFocus
    End If
    
    
End Sub

Private Sub MnuManCliPessoaJuridica_Click()
    
    If Not FunVerifica_Permissao("TelaFornecedores", "2") Then
        Exit Sub
    End If
    
    FornClie = "C"
    Pessoa = "J"
    TelaFornecedores.Show
    If FunLocalizaJanelasAtivas("TelaFornecedores") <> -1 Then
        TelaFornecedores.SetFocus
    End If
    
End Sub

Private Sub MnuConImpressora_Click()
    ' Inicia caixa de diálogo de configuração de Impressora
    CdbImpressora.ShowPrinter
End Sub

Private Sub MnuConSenhas_Click()
    
    If Not FunVerifica_Permissao("TelaOperador", "4") Then
        Exit Sub
    End If
    
    MdiPrincipal.MousePointer = vbHourglass
    TelaOperador.Show 1
    MdiPrincipal.MousePointer = vbDefault

End Sub

Private Sub MnuCotacoes_Click()
    
    If Not FunVerifica_Permissao("TelaCotacoes", "1") Then
        Exit Sub
    End If
    
    MdiPrincipal.MousePointer = vbHourglass
    XFormulario = "MdiPrincipal"
    TelaCotacoes.Show 1
    MdiPrincipal.MousePointer = vbDefault
    
End Sub

Private Sub MnuFinalizar_Click()
    Unload Me
End Sub

Private Sub MnuManEmpreendimentos_Click()
    
    If Not FunVerifica_Permissao("TelaManEmpreendimentos", "1") Then
        Exit Sub
    End If
    
    TelaManEmpreendimentos.Show
    If FunLocalizaJanelasAtivas("TelaManEmpreendimentos") <> -1 Then
        TelaManEmpreendimentos.SetFocus
    End If
End Sub

Private Sub MnuManFiadores_Click()
    
    If Not FunVerifica_Permissao("TelaManFiadores", "1") Then
        Exit Sub
    End If
    
    TelaManFiadores.Show
    If FunLocalizaJanelasAtivas("TelaManFiadores") <> -1 Then
        TelaManFiadores.SetFocus
    End If
End Sub

Private Sub MnuManImoveis_Click()
    
    If Not FunVerifica_Permissao("TelaManImoveis", "1") Then
       Exit Sub
    End If
    
    TelaManImoveis.Show
    If FunLocalizaJanelasAtivas("TelaManImoveis") <> -1 Then
        TelaManImoveis.SetFocus
    End If
    
End Sub

Private Sub MnuManTabCorretores_Click()

    If Not FunVerifica_Permissao("TelaManTabCorretores", "1") Then
        Exit Sub
    End If
    
    TelaManTabCorretores.Show
    If FunLocalizaJanelasAtivas("TelaManTabCorretores") <> -1 Then
        TelaManTabCorretores.SetFocus
    End If
End Sub

Private Sub MnuManTabIndexadores_Click()
    
    If Not FunVerifica_Permissao("TelaMoedas", "1") Then
        Exit Sub
    End If
    
    TelaMoedas.Show
    If FunLocalizaJanelasAtivas("TelaMoedas") <> -1 Then
        TelaMoedas.SetFocus
    End If
End Sub

Private Sub MnuConEmpresa_Click()
    
    MdiPrincipal.MousePointer = vbHourglass
    If pTipoCopia = "Mono" Then
        XInserir = "A"
        FrmPadraoEmpresa.Show
    Else
        TelaEmpresas.Show
    End If
    MdiPrincipal.MousePointer = vbDefault
    
End Sub

Private Sub MnuManTabObservacoes_Click()
    
    If Not FunVerifica_Permissao("TelaManTabObservacoes", "1") Then
        Exit Sub
    End If
    
    TelaManTabObservacoes.Show
    If FunLocalizaJanelasAtivas("TelaManTabObservacoes") <> -1 Then
        TelaManTabObservacoes.SetFocus
    End If
End Sub

Private Sub MnuManTabProfissoes_Click()
    
    If Not FunVerifica_Permissao("TelaManTabProfissoes", "1") Then
        Exit Sub
    End If
    
    TelaManTabProfissoes.Show
    If FunLocalizaJanelasAtivas("TelaManTabProfissoes") <> -1 Then
        TelaManTabProfissoes.SetFocus
    End If
End Sub

Private Sub MnuManTabTiposImoveis_Click()
    
    If Not FunVerifica_Permissao("TelaManTabTiposImoveis", "1") Then
        Exit Sub
    End If
    
    TelaManTabTiposImoveis.Show
    If FunLocalizaJanelasAtivas("TelaManTabTiposImoveis") <> -1 Then
        TelaManTabTiposImoveis.SetFocus
    End If
    
End Sub

Private Sub MnuManTabTiposPlanos_Click()
    
    If Not FunVerifica_Permissao("TelaManTabTiposPlanos", "1") Then
        Exit Sub
    End If
    
    TelaManTabTiposPlanos.Show
    If FunLocalizaJanelasAtivas("TelaManTabTiposPlanos") <> -1 Then
        TelaManTabTiposPlanos.SetFocus
    End If
    
End Sub

Private Sub MnuRelClientes_Click()
    
    If Not FunVerifica_Permissao("FrmRelRelacaoClientes", "1") Then
       Exit Sub
    End If
    
    FrmRelRelacaoClientes.Show
    If FunLocalizaJanelasAtivas("FrmRelRelacaoClientes") <> -1 Then
        FrmRelRelacaoClientes.SetFocus
    End If
    
End Sub

Private Sub MnuRelComissoes_Click()
    
    If Not FunVerifica_Permissao("FrmRelComissoes", "1") Then
        Exit Sub
    End If
    
    FrmRelComissoes.Show
    If FunLocalizaJanelasAtivas("FrmRelComissoes") <> -1 Then
        FrmRelComissoes.SetFocus
    End If
    
End Sub

Private Sub MnuRelSituacaoImoveis_Click()
    
    If Not FunVerifica_Permissao("FrmRelSituacaoImoveis", "1") Then
        Exit Sub
    End If
    
    XOrigem = "SituacaoImoveis"
    
    FrmRelSituacaoImoveis.Show
    If FunLocalizaJanelasAtivas("FrmRelSituacaoImoveis") <> -1 Then
        FrmRelSituacaoImoveis.SetFocus
    End If
    
End Sub
