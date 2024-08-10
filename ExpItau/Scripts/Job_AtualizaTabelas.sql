-- USE [QualiAdmFin_CA_2024_05_24]
-- GO

SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO


/****** Insere a rotina de emissão de boletos na tabela de Rotinas  ******/

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


/****** Atualiza tabela Ultimo (últimos códigos) para inserção da coluna NossoNumero do Boleto  ******/

ALTER TABLE [dbo].[Ultimos] ADD
	[Ulti_Nr_BoletoNossoNumero] int NOT NULL DEFAULT 0
GO

ALTER TABLE [dbo].[Ultimos] SET (LOCK_ESCALATION = TABLE)
GO


/****** Atualiza tabela Titulos para inserção das colunas referentes ao Boleto emitido  ******/

ALTER TABLE  [dbo].[Titulos] ADD
	[titu_tx_IdBoleto]			varchar(100) NULL,
	[titu_tx_NossoNumero]		varchar(100) NULL,
	[titu_tx_CodigoBarras]		varchar(100) NULL,
	[titu_tx_LinhaDigitavel]	varchar(100) NULL
GO


/****** Atualiza view para emissão de boletos no QualiCapi/ExpItau  ******/

CREATE OR ALTER VIEW [dbo].[ConsCAPExpBanco]
AS
SELECT  TOP 100 PERCENT
		dbo.Titulos.empd_cd_Empreendimento + '.' + dbo.Titulos.imov_cd_Imovel + '.' + dbo.Titulos.cont_cd_Contrato + '.' + dbo.Titulos.titu_cd_Plano + '.' + dbo.Titulos.titu_cd_Parcela + '.' + dbo.Titulos.titu_cd_Residuo AS Titulo,
		dbo.Titulos.coco_cd_Codigo, 
		dbo.Titulos.obse_cd_Observacao,
		dbo.Titulos.moed_cd_Moeda1,
		dbo.Titulos.moed_cd_Moeda2,
		dbo.Titulos.titu_vl_Parcela,
		dbo.Titulos.titu_vl_ParcelaIndex1,
		dbo.Titulos.titu_vl_ParcelaIndex2,
		dbo.Titulos.titu_vl_Desconto,
		dbo.Contratos.cont_vl_TaxaSegVida,
		dbo.Observacoes.obse_tx_Observacao,
		dbo.Titulos.titu_dt_Prorrogacao,
		dbo.Titulos.empr_cd_Empresa,
		dbo.Titulos.empd_cd_Empreendimento,
		dbo.Contratos.cont_nr_TipoSeguro,
		dbo.Contratos.cont_tx_CorrecProRata,
		Moedas_1.moed_tx_Indexacao AS Indexacao1,
		Moedas_2.moed_tx_Indexacao AS Indexacao2,
		dbo.Titulos.titu_vl_PercJurosFin,
		dbo.Contratos.cont_tx_Desagio,
		dbo.Titulos.imov_cd_Imovel,
		dbo.Titulos.cont_cd_Contrato,
		dbo.Titulos.titu_dt_Pagamento,
		dbo.FornClientes.focl_tx_Fantasia,
		dbo.FornClientes.focl_tx_RazaoSocial,
		dbo.FornClientes.focl_tx_CgcCpf,
		dbo.FornClientes.focl_tx_Tipo,
		dbo.FornClientes.focl_tx_Endereco,
		dbo.FornClientes.focl_tx_Bairro,
		dbo.FornClientes.focl_tx_Cidade,
		dbo.FornClientes.focl_tx_Estado,
		dbo.FornClientes.focl_tx_Cep,
		dbo.FornClientes.focl_tx_EMail,
		dbo.Clientes.clie_tx_EndResidencial,
		dbo.Clientes.clie_tx_BairroResidencial,
		dbo.Clientes.clie_nr_CepResidencial,
		dbo.Clientes.clie_tx_MunResidencial,
		dbo.Clientes.clie_tx_estresidencial,
		dbo.Clientes.clie_tx_EndCorresp,
		dbo.Clientes.clie_tx_BairroCorresp,
		dbo.Clientes.clie_nr_CepCorresp,
		dbo.Clientes.clie_tx_MunCorresp,
		dbo.Clientes.clie_tx_EstCorresp,
		dbo.Titulos.titu_tx_ExpBanco,
		dbo.Contratos.cont_tx_MultaProRata,
		dbo.Titulos.titu_vl_PercMultaMora,
		dbo.Titulos.titu_vl_PercJurosMora,
		dbo.Contratos.cont_tx_TipoJurosMora,
		dbo.Titulos.titu_dt_Vencimento,
		dbo.Contratos.cont_vl_PercJurosFin,
		dbo.Titulos.titu_tx_NaturezaPlano,
		dbo.Titulos.tipl_cd_TipoPlano,
		dbo.Contratos.cont_tx_Status,
		Moedas_2.moed_tx_Descricao AS Moeda2,
		Moedas_1.moed_tx_Descricao AS Moeda1,
		dbo.Titulos.titu_dt_Base,
		dbo.Titulos.titu_nr_ContratoBanco,
		dbo.Titulos.titu_dt_Criacao,
		dbo.NotasFiscais.nofi_nr_documento
FROM    dbo.Clientes 
		RIGHT OUTER JOIN
		dbo.FornClientes ON  dbo.Clientes.focl_cd_FornCli = dbo.FornClientes.focl_cd_FornCli 
		RIGHT OUTER JOIN
		dbo.Titulos 
		INNER JOIN
		dbo.Contratos ON  dbo.Titulos.cont_cd_Contrato       = dbo.Contratos.cont_cd_Contrato 
		              AND dbo.Titulos.imov_cd_Imovel         = dbo.Contratos.imov_cd_Imovel 
					  AND dbo.Titulos.empd_cd_Empreendimento = dbo.Contratos.empd_cd_Empreendimento 
					  AND dbo.Titulos.empr_cd_Empresa        = dbo.Contratos.empr_cd_Empresa 
					  ON  dbo.FornClientes.focl_cd_FornCli   = dbo.Contratos.focl_cd_FornCli 
		LEFT OUTER JOIN
		dbo.Moedas Moedas_2 ON  dbo.Titulos.moed_cd_Moeda2   = Moedas_2.moed_cd_CodMoeda 
		LEFT OUTER JOIN
		dbo.Moedas Moedas_1 ON  dbo.Titulos.moed_cd_Moeda1   = Moedas_1.moed_cd_CodMoeda 
		LEFT OUTER JOIN
		dbo.Observacoes ON dbo.Titulos.obse_cd_Observacao    = dbo.Observacoes.obse_cd_Observacao 
		LEFT OUTER JOIN
		dbo.NotasFiscais ON dbo.Titulos.nofi_cd_NotaFiscal   = dbo.NotasFiscais.nofi_cd_notafiscal
ORDER BY titulo

GO

