12/07/2024
-----------

Formato/schema JSON das interações entre a aplicação ExpItau (VB6), o arquivo 
de configuração, a API do Itaú e a biblioteca Boleto2Net (COM).

Obs.: Os dados/valores constantes nos Jsons abaixo, são meramente elucidativos.
O que importa são as chaves e o tipo de valor.


- Arquivo de configuração da aplicação, para a emissão dos boletos:
-------------------------------------------------------------------------------
{
    "PathToFiles": 			"C:/Temp/Boletos",
    "UseSandbox": 			false,
    "Production": {
		"EtapaProcesso": 	"validacao",
		"TipoBoleto": 		"a vista",
		"Carteira":			"109",
        "OAuth2": {
            "UrlToken": 	"https://sts.itau.com.br/api/oauth/token",
            "GrantType": 	"client_credentials",
            "ClientId": 	"7f85ee8a-8929-4e59-b578-a43a8fafdba5",
            "ClientSecret": "17159de0-9ff7-4c64-a6cb-0055888c78a0",
			"Certificate": 	"LOCAL_MACHINE\\My\\7f85ee8a-8929-4e59-b578-a43a8fafdba5"
        },
        "Endpoints": {
            "UrlBase": 		"https://api.itau.com.br/cash_management/v2",
            "Boletos": {
				"ApiKey":	"c570d9e6-1465-40d2-b8fd-be56ca5cca4a",
                "UrlPost": 	"/boletos",
                "UrlGet": 	"/boletos/{{id_boleto}}"
            }
        },
		"EmailConfig": {
			"ServerName": 	"smtp.gmail.com",
			"ServerPort": 	465,
			"UseSSL": 		true,
			"UserName": 	"psgarcia159@gmail.com",
			"UserPassword": "raqs yopx twjv vsas",
			"MailFrom": 	"Costa Andrade <no-reply@gmail.com>",
			"MailCC": 		"Costa Andrade <psgarcia159@gmail.com>",
			"MailSubject": 	"Boleto disponí­vel",
			"IsHtmlBody": 	true
		}
    },
    "Sandbox": {
		"EtapaProcesso": 	"validacao",
		"TipoBoleto": 		"a vista",
		"Carteira":			"109",
        "OAuth2": {
            "UrlToken": 	"https://sandbox.devportal.itau.com.br/api/oauth/jwt",
            "GrantType": 	"client_credentials",
            "ClientId": 	"d80d56bc-1513-3cb7-9b09-85ca65b50ff8",
            "ClientSecret": "47e60927-cab0-46e5-9629-52477bdf9bf8",
			"Certificate": 	""
        },
        "Endpoints": {
            "UrlBase": 		"https://devportal.itau.com.br/sandboxapi/cash_management_ext_v2/v2",
            "Boletos": {
				"ApiKey":	"c570d9e6-1465-40d2-b8fd-be56ca5cca4a",
                "UrlPost": 	"/boletos",
                "UrlGet": 	"/boletos/{{id_boleto}}"
            }
        },
		"EmailConfig": {
			"ServerName": 	"smtp.gmail.com",
			"ServerPort": 	465,
			"UseSSL": 		true,
			"UserName": 	"psgarcia159@gmail.com",
			"UserPassword": "raqs yopx twjv vsas",
			"MailFrom": 	"Costa Andrade <no-reply@gmail.com>",
			"MailCC": 		"Costa Andrade <psgarcia159@gmail.com>",
			"MailSubject": 	"Boleto disponí­vel",
			"IsHtmlBody": 	true
		}
    }
}


- Formato de envio dos boletos para registro na API (com indicação dos usados):
-------------------------------------------------------------------------------
    {
        "data": {
            "etapa_processo_boleto": "efetivacao",                                          # - Simulação de Emissão - 'validacao' Emissão - 'efetivacao'
            "codigo_canal_operacao": "API",                                                 # - Fixo 'API'
            "beneficiario": {   
                "id_beneficiario": "id_beneficiario"                                        # - Agência (4 dígitos) + Conta (7 dígitos) + DAC (1 dígito) =
                                                                                                Format(Left(DatContaCorrente.Recordset.Fields!coco_cd_Agencia, 4), "0000") +
                                                                                                Format(Left(DatContaCorrente.Recordset.Fields!coco_tx_Conta, 7), "0000000") +
                                                                                                Format(Right(FunNuloVal(DatContaCorrente.Recordset.Fields!coco_nr_Dac), 1), "0")
            },  
            "dado_boleto": {    
                "descricao_instrumento_cobranca": "boleto",                                 # - Fixo 'boleto'
                "forma_envio": "email",                                                     # - Deve ser indicado um dos tipos 'impressao' ou 'email'
                "texto_endereco_email": "exemplo@itau.com.br",                              # - Caso informado 'email' no campo forma_envio, é obrigatório informar um e-mail válido =
                                                                                                XFO_EXPORTACAO!focl_tx_EMail
                "assunto_email": "Assunto do e-mail",                                       # - O campo pode ser preenchido de forma personalizada. Máximo caracteres: 50
                "mensagem_email": "Texto para enviar ao cliente",                           # - O campo pode ser preenchido de forma personalizada. Máximo caracteres: 200
                "tipo_boleto": "a vista",                                                   # - Deve ser indicado um dos tipos do boleto: 'a vista' ou 'proposta'. Informar 'a vista'
                "codigo_carteira": "109",                                                   # - Fixo '109'
                "valor_titulo": "00000000000001000",                                        # - Formato do campo: 15 dígitos inteiros e 2 casas decimais =
                                                                                                Format(XLF_VALOR * 100, "00000000000000000")
                "codigo_especie": "02",                                                     # - Fixo '02' - NP - Nota Promissoria (Ver tabela de Espécies)
                "valor_abatimento": "00000000000000000",                                    # - Formato do campo: 15 dígitos inteiros e 2 casas decimais
                "data_emissao": "2022-12-21",                                               # - Formato: AAAA-MM-DD =
                                                                                                Format(DtpExportacao, "yyyy-mm-dd")
                "pagamento_parcial": true,                                                  # - Se aceita pagamento_parcial (true ou false) padrão é false. Omitir este parâmetro
                "quantidade_maximo_parcial": 2,                                             # - Se pagamento_parcial = true deve ser entre 1 e 99. Omitir este parâmetro
                "pagador": {                                                                
                    "pessoa": {                                                             
                        "nome_pessoa": "Pessoa teste",                                      # - Nome/Razão social do pagador. Máximo caracteres: 50 =
                                                                                                XLT_NOME
                        "nome_fantasia": "Pessoa teste",                                    # - Nome Fantasia do pagador. Exemplo: Empresa A. Máximo: 50 caracteres =
                                                                                                XLT_FANTASIA
                        "tipo_pessoa": {
                            "codigo_tipo_pessoa": "F",                                      # - Tipo de pessoa do pagador Pessoa Física - 'F' Pessoa Jurídica - 'J' =
                                                                                                XFO_EXPORTACAO!focl_tx_Tipo
                            "numero_cadastro_pessoa_fisica": "cpf_pagador"                  # - CPF do pagador para tipo_pessoa = F com 11 numeros (sem pontos, traços ou barras); não informar o campo caso tipo_pessoa J  = XLT_CGCCPF
                            "numero_cadastro_nacional_pessoa_juridica": "cnpj_pagador"      # - CNPJ do pagador para tipo_pessoa = J com 14 números (sem pontos, traços ou barras); não informar o campo caso tipo_pessoa F = XLT_CGCCPF
                        }
                    },
                    "endereco": {
                        "nome_logradouro": "Rua endereço,71",                               # - Nome do logradouro, número, complemento. Obrigatório caso o sacador_avalista tenha sido indicado. Máximo caracteres: 45* =
                                                                                                XLT_ENDERECO
                        "nome_bairro": "Bairro",                                            # - Nome do bairro. Obrigatório caso informe sacador avalista. Obrigatório caso o sacador_avalista tenha sido indicado. Máximo: 15 caracteres* =
                                                                                                XLT_BAIRRO
                        "nome_cidade": "Cidade",                                            # - Nome da cidade. Obrigatório caso informe sacador avalista. Obrigatório caso o sacador_avalista tenha sido indicado. Máximo: 20 caracteres* =
                                                                                                XLT_CIDADE
                        "sigla_UF": "PE",                                                   # - Sigla da UF. Obrigatório caso informe sacador avalista. Obrigatório caso o sacador_avalista tenha sido indicado. Máximo: 2 caracteres =
                                                                                                XLT_ESTADO
                        "numero_CEP": "51340540"                                            # - CEP. Obrigatório caso informe sacador avalista.Formato: 8 números, sem pontos e traços =
                                                                                                XLT_CEP
                    }                                                                       
                },                                                                          
                "sacador_avalista ": {                                                      # - Se tiver sacador/avalista Me.cboSacadorAvalista.Text = ""
                    "pessoa": {                                                             
                        "nome_pessoa": "Pessoa teste",                                      # - Nome/Razão social do avalista. Máximo caracteres: 50 =
                                                                                                DatEmpresa.Recordset.Fields!empr_tx_RazaoSocial
                        "tipo_pessoa": {
                           "codigo_tipo_pessoa": "F",                                       # - Tipo de pessoa do avalista Pessoa Física - 'F' Pessoa Jurídica - 'J'. Informar 'J'
                           "numero_cadastro_pessoa_fisica": "cpf_avalista"                  # - CPF do avalista para tipo_pessoa = F com 11 numeros (sem pontos, traços ou barras); não informar
                           "numero_cadastro_nacional_pessoa_juridica": "cnpj_avalista"      # - CNPJ do avalista para tipo_pessoa = J com 14 números (sem pontos, traços ou barras) =
                                                                                                Replace(Replace(Replace(DatEmpresa.Recordset.Fields!empr_tx_CGC, ".", ""), "/", ""), "-", "")
                        }
                    },
                    "endereco": {
                        "nome_logradouro": "Rua endereço,71",                               # - Nome do logradouro, número, complemento. Obrigatório caso o sacador_avalista tenha sido indicado. Máximo caracteres: 45* =
                                                                                                DatEmpresa.Recordset.Fields!empr_tx_Endereco
                        "nome_bairro": "Bairro",                                            # - Nome do bairro. Obrigatório caso informe sacador avalista. Obrigatório caso o sacador_avalista tenha sido indicado. Máximo: 15 caracteres* =
                                                                                                DatEmpresa.Recordset.Fields!empr_tx_Bairro
                        "nome_cidade": "Cidade",                                            # - Nome da cidade. Obrigatório caso informe sacador avalista. Obrigatório caso o sacador_avalista tenha sido indicado. Máximo: 20 caracteres*
                                                                                                DatEmpresa.Recordset.Fields!empr_tx_Cidade
                        "sigla_UF": "PE",                                                   # - Sigla da UF. Obrigatório caso informe sacador avalista. Obrigatório caso o sacador_avalista tenha sido indicado. Máximo: 2 caracteres
                                                                                                DatEmpresa.Recordset.Fields!empr_tx_Estado
                        "numero_CEP": "51340540"                                            # - CEP. Obrigatório caso informe sacador avalista.Formato: 8 números, sem pontos e traços
                                                                                                Replace(Replace(DatEmpresa.Recordset.Fields!empr_tx_Cep, "-", ""), ".", "")
                    }
                },
                "dados_individuais_boleto": [
                    {
                        "numero_nosso_numero": "20000000",                                  # - Máximo: 08 caracteres. Ver a exportação para a CEF???
                        "data_vencimento": "2023-01-14",                                    # - Formato: AAAA-MM-DD =
                                                                                                Format(XFO_EXPORTACAO!titu_dt_Vencimento, "yyyy-mm-dd")
                        "valor_titulo": "00000000000119900",                                # - Formato do campo: 15 dígitos inteiros e 2 casas decimais =
                                                                                                Format(XLF_VALOR * 100, "00000000000000000")
                        "texto_uso_beneficiario": "2",                                      # - Campo de 25 caracteres, utilizado na API legado como "identificador_titulo_empresa". Deve ser utilizado apenas letras e números.
                        "texto_seu_numero": "2"                                             # - Seu número é a identificação do boleto que poderá ter letras e números e facilitará a consulta e acompanhamento do status do boleto. Este campo é para controle do cliente e obrigatório em caso de serviço de protesto. Máximo: 10 caracteres*
                    }
                ],
                "multa": {
                    "codigo_tipo_multa": "02",                                              # - Código da multa '01' - Quando se deseja cobrar um valor fixo de multa após o vencimento. '02' - Quando se deseja cobrar um percentual do valor do título de multa após o vencimento. '03' - Quando não se deseja cobrar multa caso o pagamento seja feito após o vencimento (isento)
                    "data_multa": "2024-09-21",                                             # - Data de início de cobrança de multa. Caso o campo esteja vazio, será automaticamente assumido que a cobrança de multa se inicia logo após o vencimento. Formato: AAAA-MM-DD
                    "percentual_multa": "000000100000"                                      # - Percentual da multa cobrada. Obrigatório para tipo_multa 02. Valor calculado deve ser superior a R$0,01. Formato do campo: 7 dígitos inteiros e 5 casas decimais. Exemplo: 999999900000.
                },                                                                          
                "juros": {                                                                  
                    "codigo_tipo_juros": "90",                                              # - Tipo da cobrança dos juros no cálculo da cobrança. Para cada um dos valores informados, será impresso no boleto uma anotação referente.
                                                                                                '05' - Quando não se deseja cobrar juros caso o pagamento seja feito após o vencimento (isento)
                                                                                                '90' - Percentual mensal (utilizando parâmetros do cadastro de beneficiário para dias úteis ou corridos)
                                                                                                '91' - Percentual diário (utilizando parâmetros do cadastro de beneficiário para dias úteis ou corridos)
                                                                                                '92' - Percentual anual (utilizando parâmetros do cadastro de beneficiário para dias úteis ou corridos)
                                                                                                '93' - Valor diário (utilizando parâmetros do cadastro de beneficiário para dias úteis ou corridos)
                    "data_juros": "2024-09-21",                                             # - Data de início de cobrança de juros. Caso o campo esteja vazio, será automaticamente assumido que a cobrança de juros se inicia logo após o vencimento. Formato: AAAA-MM-DD
                    "percentual_juros": "000000100000"                                      # - Percentual dos juros a ser cobrado. Valor calculado deve ser superior a R$0,01. Obrigatório para tipo_juros ‘90’, ‘91’ e ‘92’. Formato do campo: 7 dígitos inteiros e 5 casas decimais. Exemplo: 999999900000
                },
                "recebimento_divergente": {
                    "codigo_tipo_autorizacao": "03",                                        # - Tipo de autorização de recebimento divergente da cobrança
                                                                                                '01' - Quando o título aceita qualquer valor divergente ao da cobrança.
                                                                                                '02' - Quando o título contém uma faixa de valores aceitos para recebimentos divergentes
                                                                                                '03' - Quando o título não deve aceitar pagamentos de valores divergentes ao da cobrança.
                                                                                                '04' - Quando o título aceitar pagamentos de valores superiores ao mínimo definido
                    "codigo_tipo_recebimento": "V",                                         # - Tipo de autorização de recebimento divergente da cobrança. Obrigatório para codigo_tipo_autorizacao diferente de 01 e 03. V - Recebimento divergente for informado por valores P - Recebimento divergente for informado por percentuais
                    "valor_minimo": "00000000000000100",                                    # - Valor mínimo permitido para pagamento. Obrigatório para codigo_aceite_pagamento_divergente 2 ou 4. Formato do campo: 15 dígitos inteiros e 2 casas decimais
                    "percentual_minimo": "000000100000",                                    # - Percentual mínimo permitido para pagamento. Obrigatório para codigo_tipo_autorizacao 2 ou 4. Formato do campo: 7 dígitos inteiros e 5 casas decimais
                    "valor_maximo": "00000099999999900",                                    # - Valor máximo permitido para pagamento. Obrigatório para codigo_tipo_autorizacao 2. Formato do campo: 15 dígitos inteiros e 2 casas decimais
                    "percentual_maximo": "000001000000"                                     # - Percentual máximo permitido. Obrigatório para codigo_tipo_autorizacao 2. Formato do campo: 7 dígitos inteiros e 5 casas decimais
                },
                "instrucao_cobranca": [
                    {
                        "codigo_instrucao_cobranca": "2",                                   # - Códigos das instruções de protesto. Podem ser enviadas até 3 instruções. Se houverem mais comandos, iremos descartar uma instrução aleatoriamente. Ver "Tabela de Instruções".
                        "quantidade_dias_apos_vencimento": 10,                              # - Quantidade de dias após vencimento do boleto, o prazo deve ser entre 01 e 99 dias. Para baixa também é possível 365 dias de vencido.
                        "dia_util": false                                                   # - Caso a quantidade de dias após o vencimento tenha que ser contabilizada em dia útil, informar true. Caso tenha que ser contabilizada em dias corridos, informar false. Atenção: O CEP do pagador e sacador avalista (se houver) precisa estar correto conforme Correio.
                    }
                ],
                "protesto": {
                    "protesto": true,                                                       # - Em caso de protesto informar "true". Em caso de não protesto não enviar o bloco de protesto OU enviar "false".
                    "quantidade_dias_protesto": 10                                          # - Em caso de protesto "true" (campo acima) enviar a quantidade de dias, mínimo 1 e máximo 99. Em caso de protesto "false" não enviar quantidade de dias OU enviar valor zero.
                },
                "desconto_expresso": false
            }
        }
    }


- Formato de retorno/recepção dos boletos registrados (ou não?) na API:
-------------------------------------------------------------------------------
	{
		"codigo_canal_operacao": "BKL",
		"codigo_operador": "889911348",
		"etapa_processo_boleto": "efetivacao",
		"beneficiario": {
			"id_beneficiario": "150000052061",
			"nome_cobranca": "MUNDI EMPRRENDIMENTOS E L ME",
			"tipo_pessoa": {
				"codigo_tipo_pessoa": "J",
				"numero_cadastro_nacional_pessoa_juridica": "08867659000151"
			},
			"endereco": {
				"nome_logradouro": "R PORTUGAL, 13, EDF T NOVO 1 AN",
				"nome_bairro": "COMERCIO",
				"nome_cidade": "SALVADOR",
				"sigla_UF": "BA",
				"numero_CEP": "40015000"
			}
		},
		"dado_boleto": {
			"descricao_instrumento_cobranca": "boleto",
			"forma_envio": "impressao",
			"tipo_boleto": "a vista",
			"pagador": {
				"pessoa": {
					"nome_pessoa": "Nubibat",
					"tipo_pessoa": {
						"codigo_tipo_pessoa": "F",
						"numero_cadastro_pessoa_fisica": "05201005225"
					}
				},
				"endereco": {
					"nome_logradouro": "Av Hilario Pereira de Souza, 492",
					"nome_bairro": " ",
					"nome_cidade": "Osasco",
					"sigla_UF": "SP",
					"numero_CEP": "04131020"
				},
				"pagador_eletronico_DDA": false,
				"praca_protesto": true
			},
			"sacador_avalista": {
				"pessoa": {
					"nome_pessoa": "Sacador Teste",
					"nome_fantasia": "Empresa A",
					"tipo_pessoa": {
						"codigo_tipo_pessoa": "F",
						"numero_cadastro_pessoa_fisica": "38365972840"
					}
				},
				"endereco": {
					"nome_logradouro": "Av do Estado, 55343",
					"nome_bairro": "Ipiranga",
					"nome_cidade": "São Paulo",
					"sigla_UF": "SP",
					"numero_CEP": "06120100"
				}
			},
			"codigo_carteira": "157",
			"codigo_tipo_vencimento": 3,
			"valor_total_titulo": "00000000000010001",
			"dados_individuais_boleto": [{
					"id_boleto_individual": "8835353e-ecb5-43f8-adeb-4cbf796f6be4",
					"numero_nosso_numero": "00001056",
					"dac_titulo": "8",
					"data_vencimento": "2021-06-01",
					"valor_titulo": "00000000000010001",
					"codigo_barras": "34192863800000100011570000105681500052061000",
					"numero_linha_digitavel": "34191570070010568150600520610007286380000010001",
					"data_limite_pagamento": "2031-06-01",
					"lista_mensagens_cobranca": []
				}
			],
			"codigo_especie": "01",
			"data_emissao": "2021-05-25",
			"pagamento_parcial": false,
			"quantidade_maximo_parcial": "0",
			"lista_mensagem_cobranca": [{
					"mensagem": "jaime3 desconto fixo percentual"
				}, {
					"mensagem": "teste2"
				}
			],
			"recebimento_divergente": {
				"codigo_tipo_autorizacao": "03",
				"codigo_tipo_recebimento": "P",
				"percentual_minimo": "00000000000000000",
				"percentual_maximo": "00000000000000000"
			},
			"desconto_expresso": true
		}
	}


- Formato de envio para a rotina de geração/emissão do boleto (Boleto2Net/COM):
-------------------------------------------------------------------------------
	{
		"Boleto": {
			"Aceite": "A",
			"AgenciaCobradoraRecebedora": "",
			"AvisoDebitoAutomatico": "",
			"BancoCobradorRecebedor": "",
			"Carteira": "",
			"CarteiraImpressaoBoleto": "",
			"CodigoBaixaDevolucao": 0,
			"CodigoInstrucao1": "",
			"CodigoInstrucao2": "",
			"CodigoInstrucao3": "",
			"CodigoMoeda": 9,
			"CodigoOcorrencia": "",
			"CodigoOcorrenciaAuxiliar": "",
			"CodigoProtesto": 0,
			"ComplementoInstrucao1": "",
			"ComplementoInstrucao2": "",
			"ComplementoInstrucao3": "",
			"DataCredito": "2024-07-09",
			"DataDesconto": "2024-07-09",
			"DataEmissao": "2024-07-09",
			"DataJuros": "2024-07-09",
			"DataMulta": "2024-07-09",
			"DataProcessamento": "2024-07-09",
			"DataVencimento": "2024-07-15",
			"Demonstrativos": "",
			"DescricaoOcorrencia": "",
			"DiasBaixaDevolucao": 30,
			"DiasProtesto": 3,
			"EspecieDocumento": 12,
			"EspecieMoeda": "R$",
			"IdentificadorDebitoAutomatico": "",
			"ImprimirValoresAuxiliares": true,
			"MensagemArquivoRemessa": "MensagemArquivoRemessa",
			"MensagemInstrucoesCaixa": "MensagemInstrucoesCaixa",
			"NossoNumero": "223350",
			"NossoNumeroDV": "2",
			"NossoNumeroFormatado": "109/00223350-2",
			"NumeroControleParticipante": "1234567890",
			"NumeroDocumento": "98765",
			"PercentualJurosDia": 1.5,
			"PercentualMulta": 3.0,
			"QRCode": "",
			"QuantidadeMoeda": 0,
			"RegistroArquivoRetorno": "",
			"TipoCarteira": 1,
			"UsoBanco": "",
			"ValorAbatimento": 0.00,
			"ValorDesconto": 0.00,
			"ValorIOF": 20.00,
			"ValorJurosDia": 10.00,
			"ValorMoeda": "",
			"ValorMulta": 0.00,
			"ValorOutrasDespesas": 0.00,
			"ValorOutrosCreditos": 0.00,
			"ValorPago": 0.00,
			"ValorPagoCredito": 0.00,
			"ValorTarifas": 0.00,
			"ValorTitulo": 1000.00,
			"VariacaoCarteira": ""
		},
		"Banco": {
			"Codigo": 341,
			"Digito": "7",
			"Nome": "Itaú",
			"RemoveAcentosArquivoRemessa": true
		},
		"Cedente": {
			"CPFCNPJ": "86.875.666/0001-09",
			"Codigo": "987654321",
			"CodigoDV": "0",
			"CodigoFormatado": "987654321-0",
			"CodigoTransmissao": "",
			"MostrarCNPJnoBoleto": true,
			"Nome": "Cedente Teste",
			"Observacoes": "",
			"TipoCPFCNPJ": "J",
			"ContaBancaria": {
				"Agencia": "1234",
				"CarteiraPadrao": "109",
				"CodigoBancoCorrespondente": 0,
				"Conta": "56789",
				"DigitoAgencia": "",
				"DigitoConta": "0",
				"LocalPagamento": "",
				"MensagemFixaSacado": "Mensagem fixa sacado",
				"MensagemFixaTopoBoleto": "Mensagem fixa topo boleto",
				"NossoNumeroBancoCorrespondente": "12345678",
				"OperacaoConta": "",
				"TipoCarteiraPadrao": 1,
				"TipoDistribuicao": 2,
				"TipoDocumento": 1,
				"TipoFormaCadastramento": 1,
				"TipoImpressaoBoleto": 2,
				"VariacaoCarteiraPadrao": ""
			},
			"Endereco": {
				"LogradouroEndereco": "Rua Teste do Banco",
				"LogradouroNumero": "1234",
				"LogradouroComplemento": "Conunto 341",
				"Bairro": "Brotas",
				"Cidade": "Salvador",
				"UF ": "BA",
				"CEP": "4000000"
			}
		},
		"Sacado": {
			"CPFCNPJ": "71.738.978/0001-01",
			"Nome": "Sacado Teste PJ",
			"Observacoes": "Matricula 123/4",
			"Endereco": {
				"LogradouroEndereco": "Rua Teste do Sacado",
				"LogradouroNumero": "567",
				"LogradouroComplemento": "Edf. Mandarim",
				"Bairro": "Caminho das Árvores",
				"Cidade": "Salvador",
				"UF ": "BA",
				"CEP": "41820774"
			}
		},
		"Avalista": {
			"CPFCNPJ": "86.875.666/0001-09",
			"Nome": "Avalista Sem Nome",
			"Observacoes": "Observacoes inválidas",
			"Endereco": {
				"LogradouroEndereco": "Rua Teste do Avalista",
				"LogradouroNumero": "370",
				"LogradouroComplemento": "Apto. 1705",
				"Bairro": "Caminho das Árvores",
				"Cidade": "Salvador",
				"UF ": "BA",
				"CEP": "41820123"
			}
		},
		"CodigoBarra": {
			"CampoLivre": "",
			"CodigoBanco": "CodigoBanco",
			"CodigoDeBarras": "34191719500000600001090022335021234567890000",
			"DigitoVerificador": "1",
			"FatorVencimento": 0.00,
			"LinhaDigitavel": "34191.09008 22335.021238 45678.900007 1 71950000100000",
			"Moeda": 9,
			"ValorDocumento": "100000"
		},
		"PathToFiles": "C:/Temp/Boletos"
	}

- Formato de retorno/recepção da geração dos boletos:
-------------------------------------------------------------------------------
	{
		"status": "OK",
		"message": "Boleto emitido!",
		"arquivoPDF": "C:\Temp\Boletos\Boleto_109_98765.PDF",
		"arquivoHTML": "C:\Temp\Boletos\Boleto_109_98765.html"
	}
