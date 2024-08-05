-- USE [QualiAdmFin_CA_2024_05_24]
-- GO

/****** Atualiza tabela Ultimo (últimos códigos) para inserção da coluna  Nosso Numero do Boleto  ******/

SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

ALTER TABLE dbo.Ultimos ADD
	Ulti_Nr_BoletoNossoNumero int NOT NULL DEFAULT 0
GO

ALTER TABLE dbo.Ultimos SET (LOCK_ESCALATION = TABLE)
GO

