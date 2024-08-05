--USE [QualiAdmFin_CA_2024_05_24]
--GO

/****** Atualiza tabela de Rotinas ******/

INSERT INTO [dbo].[Rotinas]
           ([roti_tx_Programa]
           ,[roti_tx_Ordem]
           ,[roti_tx_Classificacao]
           ,[roti_tx_Descricao]
           ,[roti_tx_Modulo])
     VALUES
           ('TelaExpBancaria'
           ,'3'
           ,'Movimentação'
           ,'Exportação bancária: Emitir Boletos'
           ,'AdmGenerico')
GO


