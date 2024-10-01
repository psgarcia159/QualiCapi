-- USE [QualiAdmFin_CA_2024_05_24]
-- GO

/****** View para emissão de boletos no QualiCapi/ExpItau (solicitação Costa Andrade) ******/
DROP VIEW IF EXISTS [dbo].[ConsCAPEmissaoBoleto]
GO

SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE VIEW [dbo].[ConsCAPEmissaoBoleto]
AS
SELECT  dbo.Titulos.empd_cd_Empreendimento + '.' + dbo.Titulos.imov_cd_Imovel + '.' + dbo.Titulos.cont_cd_Contrato + '.' + dbo.Titulos.titu_cd_Plano + '.' + dbo.Titulos.titu_cd_Parcela + '.' + dbo.Titulos.titu_cd_Residuo AS Titulo,
		dbo.Titulos.coco_cd_Codigo, 
		dbo.Titulos.cont_cd_Contrato,
		dbo.Titulos.empd_cd_Empreendimento,
		dbo.Titulos.empr_cd_Empresa,
		dbo.Titulos.imov_cd_Imovel,
		dbo.Titulos.moed_cd_Moeda1,
		dbo.Titulos.moed_cd_Moeda2,
		dbo.Titulos.obse_cd_Observacao,
		dbo.Titulos.tipl_cd_TipoPlano,
		dbo.Titulos.titu_dt_Base,
		dbo.Titulos.titu_dt_Criacao,
		dbo.Titulos.titu_dt_Pagamento,
		dbo.Titulos.titu_dt_Prorrogacao,
		dbo.Titulos.titu_dt_Vencimento,
		dbo.Titulos.titu_nr_ContratoBanco,
		dbo.Titulos.titu_tx_ExpBanco,
		dbo.Titulos.titu_tx_NaturezaPlano,
		dbo.Titulos.titu_vl_Desconto,
		dbo.Titulos.titu_vl_Parcela,
		dbo.Titulos.titu_vl_ParcelaIndex1,
		dbo.Titulos.titu_vl_ParcelaIndex2,
		dbo.Titulos.titu_vl_PercJurosFin,
		dbo.Titulos.titu_vl_PercJurosMora,
		dbo.Titulos.titu_vl_PercMultaMora,
		dbo.Empreendimentos.empd_tx_Nome,
		dbo.Contratos.cont_nr_TipoSeguro,
		dbo.Contratos.cont_tx_CorrecProRata,
		dbo.Contratos.cont_tx_Desagio,
		dbo.Contratos.cont_tx_MultaProRata,
		dbo.Contratos.cont_tx_Status,
		dbo.Contratos.cont_tx_TipoJurosMora,
		dbo.Contratos.cont_vl_PercJurosFin,
		dbo.Contratos.cont_vl_TaxaSegVida,
		dbo.FornClientes.focl_tx_Bairro,
		dbo.FornClientes.focl_tx_Cep,
		dbo.FornClientes.focl_tx_CgcCpf,
		dbo.FornClientes.focl_tx_Cidade,
		dbo.FornClientes.focl_tx_EMail,
		dbo.FornClientes.focl_tx_Endereco,
		dbo.FornClientes.focl_tx_Estado,
		dbo.FornClientes.focl_tx_Fantasia,
		dbo.FornClientes.focl_tx_RazaoSocial,
		dbo.FornClientes.focl_tx_Tipo,
		dbo.Clientes.clie_nr_CepCorresp,
		dbo.Clientes.clie_nr_CepResidencial,
		dbo.Clientes.clie_tx_BairroCorresp,
		dbo.Clientes.clie_tx_BairroResidencial,
		dbo.Clientes.clie_tx_EndCorresp,
		dbo.Clientes.clie_tx_EndResidencial,
		dbo.Clientes.clie_tx_EstCorresp,
		dbo.Clientes.clie_tx_MunCorresp,
		dbo.Clientes.clie_tx_MunResidencial,
		dbo.Clientes.clie_tx_estresidencial,
		Moedas_1.moed_tx_Descricao 				AS Moeda1,
		Moedas_1.moed_tx_Indexacao 				AS Indexacao1,
		Moedas_2.moed_tx_Descricao 				AS Moeda2,
		Moedas_2.moed_tx_Indexacao 				AS Indexacao2,
		dbo.Observacoes.obse_tx_Observacao,
		dbo.NotasFiscais.nofi_nr_documento
FROM    dbo.Titulos 
		INNER JOIN
		dbo.Empreendimentos ON dbo.Titulos.empd_cd_Empreendimento = dbo.Empreendimentos.empd_cd_Empreendimento 
					       AND dbo.Titulos.empr_cd_Empresa        = dbo.Empreendimentos.empr_cd_Empresa 
		INNER JOIN
		dbo.Contratos ON dbo.Titulos.cont_cd_Contrato        = dbo.Contratos.cont_cd_Contrato 
		             AND dbo.Titulos.imov_cd_Imovel          = dbo.Contratos.imov_cd_Imovel 
					 AND dbo.Titulos.empd_cd_Empreendimento  = dbo.Contratos.empd_cd_Empreendimento 
					 AND dbo.Titulos.empr_cd_Empresa         = dbo.Contratos.empr_cd_Empresa 
		INNER JOIN
		dbo.FornClientes ON dbo.FornClientes.focl_cd_FornCli = dbo.Contratos.focl_cd_FornCli 
		LEFT OUTER JOIN
		dbo.Clientes  ON  dbo.Clientes.focl_cd_FornCli       = dbo.FornClientes.focl_cd_FornCli 
		LEFT OUTER JOIN
		dbo.Moedas AS Moedas_2 ON dbo.Titulos.moed_cd_Moeda2 = Moedas_2.moed_cd_CodMoeda 
		LEFT OUTER JOIN
		dbo.Moedas AS Moedas_1 ON dbo.Titulos.moed_cd_Moeda1 = Moedas_1.moed_cd_CodMoeda 
		LEFT OUTER JOIN
		dbo.Observacoes ON dbo.Titulos.obse_cd_Observacao    = dbo.Observacoes.obse_cd_Observacao 
		LEFT OUTER JOIN
		dbo.NotasFiscais ON dbo.Titulos.nofi_cd_NotaFiscal   = dbo.NotasFiscais.nofi_cd_notafiscal

GO
