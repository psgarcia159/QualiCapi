--
-- Script para a alteração das tabelas:
--		[dbo].[Comissoes] e
--		[dbo].[CorretoresPagamento]
--
-- Objetivo: permitir múltiplas entradas para o mesmo corretor
--
-- Observações: É necessária a compilação de views e procedures
--					[dbo].[ConsCAPCorretores]
--					[dbo].[ConsCAPRelComissoes]
--
--    			Os dados das duplicatas, deverão ser obtidos por
--				views/queries que relacionem a Nota Fiscal com a
--				Duplicata pelo número da fatura.
--		
---------------------------------------------------------------
USE [QualiAdmFin]
GO

SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

--
-- Adiciona o campo Identity e substitui as constraints
-- "Primary Key"
-- Obs.: O campo id será gerado automaticamente
---------------------------------------------------------------
ALTER TABLE [dbo].[Comissoes] ADD [comi_cd_Comissoes] [int] IDENTITY(1,1) NOT NULL
GO

ALTER TABLE [dbo].[Comissoes] DROP CONSTRAINT [COMIPK01]
GO

ALTER TABLE [dbo].[Comissoes] ADD CONSTRAINT [PK_COMI01] PRIMARY KEY ([comi_cd_Comissoes] ASC)
GO

ALTER TABLE [dbo].[CorretoresPagamento] ADD [copa_cd_CorretoresPagamento] [int] IDENTITY(1,1) NOT NULL
GO

ALTER TABLE [dbo].[CorretoresPagamento] DROP CONSTRAINT [PK_CorretoresPagamento]
GO

ALTER TABLE [dbo].[CorretoresPagamento] ADD CONSTRAINT [PK_COPA01] PRIMARY KEY ([copa_cd_CorretoresPagamento] ASC)
GO

ALTER TABLE [dbo].[CorretoresPagamento] ADD [comi_cd_Comissoes] [int] NULL
GO

ALTER TABLE [dbo].[CorretoresPagamento] ADD CONSTRAINT [FK_COMI_COPA] FOREIGN KEY([comi_cd_Comissoes])
 REFERENCES [dbo].[Comissoes] ([comi_cd_Comissoes])
GO

ALTER TABLE [dbo].[CorretoresPagamento] ADD [nofi_cd_notafiscal] [int] NULL
GO

ALTER TABLE [dbo].[CorretoresPagamento] ADD CONSTRAINT [FK_NOFI_COPA] FOREIGN KEY([nofi_cd_notafiscal])
 REFERENCES [dbo].[NotasFiscais] ([nofi_cd_notafiscal])
GO

-- ALTER TABLE [dbo].[CorretoresPagamento] ADD [dupl_cd_duplicata] [int] NULL
-- GO

-- ALTER TABLE [dbo].[CorretoresPagamento] ADD CONSTRAINT [FK_DUPL_COPA] FOREIGN KEY([dupl_cd_duplicata])
--  REFERENCES [dbo].[Duplicatas] ([dupl_cd_duplicata])
-- GO

--
-- Fim
---------------------------------------------------------------

