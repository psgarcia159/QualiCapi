--
-- Script SQL para alteração criação de tabelas:
-- 	Alteração:
--		dbo.Corretores
--	Criação:
--		dbo.CorretoresEmpreendimento
--		dbo.CorretoresPagamento
------------------------------------------------------------------------------------------------------

USE QualiAdmFin
GO

SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

------------------------------------------------------------------------------------------------------

ALTER TABLE dbo.Corretores 
		ADD focl_cd_FornCli int NULL;

ALTER TABLE dbo.Corretores 
		ADD CONSTRAINT FK_FornClientes_Corretores FOREIGN KEY (focl_cd_FornCli) REFERENCES dbo.FornClientes (focl_cd_FornCli);
GO

------------------------------------------------------------------------------------------------------

CREATE TABLE [dbo].[CorretoresEmpreendimento](
	[empr_cd_Empresa] 				[int] 				NOT NULL,
	[empd_cd_Empreendimento] 		[varchar](4) 		NOT NULL,
	[corr_cd_Corretor] 				[int] 				NOT NULL,
	[coem_vl_PercentPadrao] 		[decimal](9, 2) 		NULL,
PRIMARY KEY CLUSTERED 
(
	[empr_cd_Empresa] ASC,
	[empd_cd_Empreendimento] ASC,
	[corr_cd_Corretor] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO

ALTER TABLE [dbo].[CorretoresEmpreendimento]  WITH CHECK ADD  CONSTRAINT [FK_CORR_COEM] FOREIGN KEY([corr_cd_Corretor])
 REFERENCES [dbo].[Corretores] ([corr_cd_Corretor])
GO

ALTER TABLE [dbo].[CorretoresEmpreendimento] CHECK CONSTRAINT [FK_CORR_COEM]
GO

ALTER TABLE [dbo].[CorretoresEmpreendimento]  WITH CHECK ADD  CONSTRAINT [FK_EMPD_COEM] FOREIGN KEY([empd_cd_Empreendimento], [empr_cd_Empresa])
 REFERENCES [dbo].[Empreendimentos] ([empd_cd_Empreendimento], [empr_cd_Empresa])
GO

ALTER TABLE [dbo].[CorretoresEmpreendimento] CHECK CONSTRAINT [FK_EMPD_COEM]
GO

------------------------------------------------------------------------------------------------------

CREATE TABLE [dbo].[CorretoresPagamento](
	[empr_cd_Empresa] 				[int] 				NOT NULL,	-- FK_CONT_COPA
	[empd_cd_Empreendimento] 		[varchar](4) 		NOT NULL,	-- FK_CONT_COPA
	[imov_cd_Imovel] 				[varchar](4) 		NOT NULL,	-- FK_CONT_COPA
	[cont_cd_Contrato] 				[varchar](2) 		NOT NULL,	-- FK_CONT_COPA
	[corr_cd_Corretor]				[int] 				NOT NULL,	-- FK_CORR_COPA
	[focl_cd_FornCli] 				[int] 					NULL,   -- FK_FOCL_COPA
	[tido_cd_TipoDoc] 				[int] 					NULL,   -- FK_TIDO_COPA
	[cere_cd_Pcr] 					[varchar](8) 			NULL,	-- FK_CERE_COPA
	[coco_cd_Codigo] 				[int] 					NULL,	-- FK_COCO_COPA
	[plco_cd_Conta] 				[varchar](7) 			NULL,	-- FK_PLCO_COPA
	[iden_cd_Projeto] 				[int] 					NULL,	-- FK_IDEN_COPA
	[copa_tx_codigo]				[varchar](22)			NULL,	--
	[copa_nr_documento] 			[varchar](14) 			NULL,   -- 
	[copa_tx_serie] 				[varchar](3) 			NULL,   -- 
	[copa_tx_statusDocumento] 		[varchar](2) 			NULL,   -- 
	[copa_dt_entrada] 				[datetime] 				NULL,   -- 
	[copa_dt_emissao] 				[datetime] 				NULL,   -- 
	[copa_dt_vencimento]			[datetime] 				NULL,   -- 
	[copa_tx_historico] 			[varchar](100) 			NULL,   -- 
	[copa_vl_valor] 				[float] 				NULL,   -- 
	[copa_nr_titulocapi] 			[varchar](22) 			NULL,   -- 
	[copa_tx_statusPagamento]		[char](1)				NULL,	-- P(processado no QualiCapi), T(transferido para o QualiFin)
 CONSTRAINT [PK_CorretoresPagamento] PRIMARY KEY CLUSTERED 
(
	[empr_cd_Empresa] 			ASC,
	[empd_cd_Empreendimento] 	ASC,
	[imov_cd_Imovel] 			ASC,
	[cont_cd_Contrato] 			ASC,
	[corr_cd_Corretor] 			ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO

ALTER TABLE [dbo].[CorretoresPagamento]  WITH NOCHECK ADD  CONSTRAINT [FK_CONT_COPA] FOREIGN KEY([cont_cd_Contrato], [imov_cd_Imovel], [empd_cd_Empreendimento], [empr_cd_Empresa])
 REFERENCES [dbo].[Contratos] ([cont_cd_Contrato], [imov_cd_Imovel], [empd_cd_Empreendimento], [empr_cd_Empresa])
GO

ALTER TABLE [dbo].[CorretoresPagamento] CHECK CONSTRAINT [FK_CONT_COPA]
GO

ALTER TABLE [dbo].[CorretoresPagamento]  WITH NOCHECK ADD CONSTRAINT [FK_CERE_COPA] FOREIGN KEY([cere_cd_Pcr])
 REFERENCES [dbo].[CentroResultados] ([cere_cd_Pcr])
GO

ALTER TABLE [dbo].[CorretoresPagamento] CHECK CONSTRAINT [FK_CERE_COPA]
GO

ALTER TABLE [dbo].[CorretoresPagamento]  WITH NOCHECK ADD  CONSTRAINT [FK_COCO_COPA] FOREIGN KEY([coco_cd_Codigo])
REFERENCES [dbo].[ContasCorrente] ([coco_cd_codigo])
GO

ALTER TABLE [dbo].[CorretoresPagamento] CHECK CONSTRAINT [FK_COCO_COPA]
GO

ALTER TABLE [dbo].[CorretoresPagamento]  WITH NOCHECK ADD  CONSTRAINT [FK_IDEN_COPA] FOREIGN KEY([iden_cd_Projeto])
REFERENCES [dbo].[Identificadores] ([iden_cd_Projeto])
GO

ALTER TABLE [dbo].[CorretoresPagamento] CHECK CONSTRAINT [FK_IDEN_COPA]
GO

ALTER TABLE [dbo].[CorretoresPagamento]  WITH NOCHECK ADD  CONSTRAINT [FK_PLCO_COPA] FOREIGN KEY([plco_cd_Conta])
REFERENCES [dbo].[PlanoContas] ([plco_cd_Conta])
GO

ALTER TABLE [dbo].[CorretoresPagamento] CHECK CONSTRAINT [FK_PLCO_COPA]
GO

ALTER TABLE [dbo].[CorretoresPagamento]  WITH NOCHECK ADD  CONSTRAINT [FK_TIDO_COPA] FOREIGN KEY([tido_cd_TipoDoc])
REFERENCES [dbo].[TipoDocumentos] ([tido_cd_TipoDoc])
GO

ALTER TABLE [dbo].[CorretoresPagamento] CHECK CONSTRAINT [FK_TIDO_COPA]
GO

ALTER TABLE [dbo].[CorretoresPagamento]  WITH NOCHECK ADD  CONSTRAINT [FK_CORR_COPA] FOREIGN KEY([corr_cd_Corretor])
REFERENCES [dbo].[Corretores] ([corr_cd_Corretor])
GO

ALTER TABLE [dbo].[CorretoresPagamento] CHECK CONSTRAINT [FK_CORR_COPA]
GO

ALTER TABLE [dbo].[CorretoresPagamento]  WITH CHECK ADD  CONSTRAINT [FK_FOCL_COPA] FOREIGN KEY([focl_cd_FornCli])
REFERENCES [dbo].[FornClientes] ([focl_cd_FornCli])
GO

ALTER TABLE [dbo].[CorretoresPagamento] CHECK CONSTRAINT [FK_FOCL_COPA]
GO

------------------------------------------------------------------------------------------------------

