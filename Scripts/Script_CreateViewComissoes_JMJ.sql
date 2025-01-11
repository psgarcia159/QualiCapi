--
-- Script para a criação da view usada na leitura das comissões
--      
---------------------------------------------------------------
USE [QualiAdmFin]
GO

SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE OR ALTER VIEW [dbo].[ConsCAPRelComissoesPorUnidade]
AS
SELECT  
        dbo.Comissoes.empr_cd_Empresa,
        dbo.Comissoes.empd_cd_Empreendimento,
        dbo.Empreendimentos.empd_tx_Nome,
        dbo.Comissoes.imov_cd_Imovel,
        dbo.Imoveis.imov_vl_AVista,
        dbo.Comissoes.cont_cd_Contrato,
        dbo.Comissoes.imov_cd_Imovel + ' - ' + dbo.FornClientes.focl_tx_RazaoSocial AS Unidade,
        dbo.Contratos.cont_dt_Venda,
        dbo.Contratos.cont_dt_RegistroContrato,
        dbo.Comissoes.corr_cd_Corretor,
        dbo.Comissoes.comi_vl_Comissao,
        dbo.Comissoes.comi_dt_Pagamento,
        dbo.Corretores.corr_tx_Nome,
        dbo.CorretoresPagamento.copa_nr_titulocapi,
        dbo.NotasFiscais.nofi_cd_notafiscal,
        dbo.NotasFiscais.nofi_nr_documento,
        dbo.NotasFiscais.nofi_tx_serie,
        dbo.NotasFiscais.focl_cd_FornCli,
        dbo.NotasFiscais.nofi_nr_fatura,
        dbo.Duplicatas.dupl_cd_duplicata,
        dbo.Duplicatas.dupl_dt_vencimento,
        dbo.Duplicatas.dupl_dt_PagRec   
  FROM  dbo.Comissoes
        INNER JOIN
        dbo.Contratos            ON dbo.Comissoes.cont_cd_Contrato             = dbo.Contratos.cont_cd_Contrato 
                                AND dbo.Comissoes.imov_cd_Imovel               = dbo.Contratos.imov_cd_Imovel 
                                AND dbo.Comissoes.empd_cd_Empreendimento       = dbo.Contratos.empd_cd_Empreendimento 
                                AND dbo.Comissoes.empr_cd_Empresa              = dbo.Contratos.empr_cd_Empresa 
        INNER JOIN
        dbo.FornClientes         ON dbo.Contratos.focl_cd_FornCli              = dbo.FornClientes.focl_cd_FornCli
        INNER JOIN                                                         
        dbo.Imoveis              ON dbo.Imoveis.imov_cd_Imovel                 = dbo.Contratos.imov_cd_Imovel 
                                AND dbo.Imoveis.empd_cd_Empreendimento         = dbo.Contratos.empd_cd_Empreendimento 
                                AND dbo.Imoveis.empr_cd_Empresa                = dbo.Contratos.empr_cd_Empresa 
        INNER JOIN
        dbo.Empreendimentos      ON dbo.Empreendimentos.empr_cd_Empresa        = dbo.Imoveis.empr_cd_Empresa 
                                AND dbo.Empreendimentos.empd_cd_Empreendimento = dbo.Imoveis.empd_cd_Empreendimento
                                AND dbo.Empreendimentos.empd_cd_Empreendimento = dbo.Imoveis.empd_cd_Empreendimento 
        INNER JOIN
        dbo.Corretores           ON dbo.Comissoes.corr_cd_Corretor             = dbo.Corretores.corr_cd_Corretor
        LEFT OUTER JOIN
        dbo.CorretoresPagamento  ON dbo.Comissoes.comi_cd_Comissoes            = dbo.CorretoresPagamento.comi_cd_Comissoes
        LEFT OUTER JOIN
        dbo.NotasFiscais         ON dbo.CorretoresPagamento.nofi_cd_notafiscal = dbo.NotasFiscais.nofi_cd_notafiscal
        LEFT OUTER JOIN
        dbo.Duplicatas           ON dbo.NotasFiscais.nofi_nr_fatura            = dbo.Duplicatas.dupl_nr_fatura

GO


