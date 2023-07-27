--
-- Script para a alteração da tabela [dbo].[Comissoes] para
-- permitir múltiplas entradas para o mesmo corretor
-- Obs.: A tabela original será mantida por motivo de segurança
-- e eventual retorno ao layout atual
--
---------------------------------------------------------------
USE [QualiAdmFin]
GO

SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

--
-- Cria nova tabela de comissões, alterando a primary key
---------------------------------------------------------------
CREATE TABLE [dbo].[ComissoesNew](
	[comi_cd_Comissoes]			[int] IDENTITY(1,1) NOT NULL,
	[corr_cd_Corretor]			[int] 				NOT NULL,
	[empd_cd_Empreendimento]	[varchar](4) 		NOT NULL,
	[imov_cd_Imovel]			[varchar](4) 		NOT NULL,
	[cont_cd_Contrato]			[varchar](2) 		NOT NULL,
	[empr_cd_Empresa]			[int] 				NOT NULL,
	[comi_vl_Comissao]			[decimal](18, 2) 		NULL,
	[comi_dt_Pagamento]			[datetime] 				NULL,
	CONSTRAINT [PK_COMI01] PRIMARY KEY CLUSTERED 
	(
		[comi_cd_Comissoes] ASC
	)
) ON [PRIMARY]
GO

ALTER TABLE [dbo].[ComissoesNew]  WITH NOCHECK ADD  CONSTRAINT [FK_CONT_COMI] FOREIGN KEY([cont_cd_Contrato], [imov_cd_Imovel], [empd_cd_Empreendimento], [empr_cd_Empresa])
REFERENCES [dbo].[Contratos] ([cont_cd_Contrato], [imov_cd_Imovel], [empd_cd_Empreendimento], [empr_cd_Empresa])
ON UPDATE CASCADE
ON DELETE CASCADE
GO

ALTER TABLE [dbo].[ComissoesNew] CHECK CONSTRAINT [FK_CONT_COMI]
GO

ALTER TABLE [dbo].[ComissoesNew]  WITH NOCHECK ADD  CONSTRAINT [FK_CORR_COMI] FOREIGN KEY([corr_cd_Corretor])
REFERENCES [dbo].[Corretores] ([corr_cd_Corretor])
GO

ALTER TABLE [dbo].[ComissoesNew] CHECK CONSTRAINT [FK_CORR_COMI]
GO

--
-- Popula tabela [ComissoesNew] a partir da tabela [Comissoes]
---------------------------------------------------------------
INSERT INTO [dbo].[ComissoesNew]
	 SELECT * 
	   FROM [dbo].[Comissoes]
GO

--
-- Renomeia as 2 tabelas
---------------------------------------------------------------
EXEC sp_rename 'dbo.Comissoes', 'ComissoesOld'

EXEC sp_rename 'dbo.ComissoesNew', 'Comissoes'


--
-- Fim
---------------------------------------------------------------
