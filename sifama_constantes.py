"""Constantes e mapeamentos usados no fluxo de assinatura de lotes do SIFAMA."""

URL_LOGIN = "https://appweb1.antt.gov.br/sca/Site/Login.aspx"
URL_LOTES = "https://appweb1.antt.gov.br/spm/Site/PublicacaoDOU/ConsultarAssinarLotePublicacaoSimplesDOU.aspx"

_CP = "ContentPlaceHolderCorpo_" * 4

ID_LOGIN_USUARIO = f"{_CP}TextBoxUsuario"
ID_LOGIN_SENHA = f"{_CP}TextBoxSenha"
ID_LOGIN_BTN = f"{_CP}ButtonOk"

ID_DATA_INICIAL = f"{_CP}txbDataInicial"
ID_DATA_FINAL = f"{_CP}txbDataFinal"
ID_TIPO_PUBLICACAO = f"{_CP}ddlTipoPublicacao"
ID_FORMA_FISCALIZACAO = f"{_CP}ddlFormaFiscalizacao"
ID_TIPO_FISCALIZACAO = f"{_CP}tipoSubTipoFiscalizacao_ddlTipoFiscalizacao"
ID_SUBTIPO_FISCALIZACAO = f"{_CP}tipoSubTipoFiscalizacao_ddlSubTipoFiscalizacao"
ID_BTN_PESQUISAR = f"{_CP}btnPesquisar"
ID_TABELA_LOTES = f"{_CP}gdvLotePublicacao"
CSS_BTN_ASSINAR = f"[id^='{_CP}gdvLotePublicacao_btnAssinarAuto_']"
ID_BTN_PROX_PAG = f"{_CP}ucPaginador_ucPaginador_lbNextPage"

ID_SENHA_CERT = "ContentPlaceHolderCorpo_" * 3 + "ucSenhaCertificadoDigital_txbSenhaCertificadoDigital"
ID_BTN_SALVAR = "ContentPlaceHolderCorpo_" * 3 + "btnSalvar"
ID_MESSAGEBOX_OK = "MessageBox_ButtonOk"

TIPOS_PUBLICACAO = {
    "Notificação de Autuação": "1",
    "Notificação de Multa": "2",
    "Cancelamento": "4",
    "Notificação de Segunda Multa": "10",
    "Notificação de Penalidade": "16",
    "Notificação Final de Multa": "17",
}

FORMAS_FISCALIZACAO = {
    "": "--Selecione--",
    "1": "Eletrônica",
    "2": "Manual",
    "3": "Remota",
}
LISTA_FORMAS = [("", "--Selecione--"), ("1", "Eletrônica"), ("2", "Manual"), ("3", "Remota")]

TIPOS_FISCALIZACAO = {
    "": "--Selecione--",
    "2": "Excesso de Peso",
    "3": "Cargas",
    "4": "Passageiros",
    "5": "Cargas Internacional",
    "7": "Passageiros Internacional",
    "8": "Infraestrutura Rodoviária",
    "9": "Evasão de Pedágio",
}
LISTA_TIPOS_FISC = [
    ("", "--Selecione--"),
    ("2", "Excesso de Peso"),
    ("3", "Cargas"),
    ("4", "Passageiros"),
    ("5", "Cargas Internacional"),
    ("7", "Passageiros Internacional"),
    ("8", "Infraestrutura Rodoviária"),
    ("9", "Evasão de Pedágio"),
]

SUBTIPOS_POR_TIPO = {
    "2": [("", "--Selecione--"), ("5", "Excesso de Peso"), ("7", "CMT - Capacidade Máxima de Tração"), ("16", "Evasão de Balança")],
    "3": [("", "--Selecione--"), ("8", "RNTRC - Registro Nacional de Transportadores Rodoviários de Cargas"), ("9", "PEF - Pagamento Eletrônico de Frete"), ("10", "Vale Pedágio"), ("17", "Produtos Perigosos"), ("24", "Piso Mínimo de Frete")],
    "4": [("", "--Selecione--"), ("11", "Longa Distância"), ("12", "Semiurbano"), ("13", "Fretamento"), ("14", "Não Autorizado"), ("19", "Passageiro Econômico Financeiro"), ("20", "Fretamento Contínuo"), ("23", "Ferroviário de Passageiros")],
    "5": [("", "--Selecione--"), ("21", "Cargas Internacional"), ("22", "Produtos Perigosos Internacional")],
    "7": [("", "--Selecione--"), ("25", "Longa Distância"), ("26", "Semiurbano"), ("27", "Fretamento"), ("28", "Não Autorizado"), ("29", "Fretamento Contínuo")],
    "8": [("", "--Selecione--"), ("30", "Infraestrutura Rodoviária")],
    "9": [("", "--Selecione--"), ("31", "Evasão de Pedágio")],
}

DELAYS = {
    "apos_clicar_assinar": 1.2,
    "carregar_guia_assinatura": 2.0,
    "apos_preencher_senha": 0.8,
    "aguardar_messagebox": 600,
    "log_aguardando_ok_cada": 15,
    "apos_progresso_sumir": 1.5,
    "apos_clicar_ok": 1.0,
    "apos_voltar_aba": 1.5,
    "entre_lotes": 1.5,
    "carregar_proxima_pagina": 2.0,
    "apos_refresh": 4.0,
    "apos_fechar_navegador": 2.0,
}

DELAYS_SENHA_CLIQUES = frozenset(
    {
        "apos_clicar_assinar",
        "carregar_guia_assinatura",
        "apos_preencher_senha",
        "apos_progresso_sumir",
        "apos_clicar_ok",
        "apos_voltar_aba",
        "entre_lotes",
    }
)
