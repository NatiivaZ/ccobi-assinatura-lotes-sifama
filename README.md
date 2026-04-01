# Automação — Assinatura de Lotes DOU (SIFAMA / ANTT)

Ferramenta desktop para **automatizar a assinatura digital de lotes de publicação no Diário Oficial da União (DOU)** no ecossistema **SIFAMA**, acessando os portais oficiais da **ANTT** (Agência Nacional de Transportes Terrestres).  
Desenvolvida no contexto do projeto **CCOBI – SERASA**.

---

## Objetivo

Reduzir trabalho manual repetitivo ao:

1. Autenticar no sistema SCA da ANTT.
2. Abrir a tela de **consulta e assinatura de lotes de publicação simples no DOU**.
3. Aplicar filtros (período, tipo de publicação, forma e tipo/subtipo de fiscalização).
4. Percorrer a grade de lotes, clicar em **Assinar** para cada item desejado.
5. Preencher a **senha do certificado digital** na janela de assinatura, confirmar e voltar à lista até concluir a fila (incluindo **paginação**).

A interface gráfica (**Tkinter**) permite ajustar tempos de espera, acompanhar logs e exportar relatório em Excel quando aplicável.

---

## Requisitos

| Item | Detalhe |
|------|---------|
| **Python** | 3.8 ou superior (recomendado 3.10+) |
| **Google Chrome** | Instalado e atualizado |
| **ChromeDriver** | Pode ser empacotado localmente (`chromedriver.exe` na pasta do projeto) ou gerenciado conforme sua configuração do Selenium |
| **Certificado digital** | Válido para assinatura nos sistemas da ANTT |
| **Credenciais** | Usuário e senha do portal SCA/ANTT |

### Dependências Python

```bash
pip install -r requirements.txt
```

- `selenium` — automação do navegador  
- `openpyxl` — geração de planilhas de log/relatório (quando habilitado)

---

## Instalação rápida

1. Clone ou copie esta pasta do projeto.  
2. Crie um ambiente virtual (opcional, recomendado):

   ```bash
   python -m venv venv
   venv\Scripts\activate
   ```

3. Instale dependências: `pip install -r requirements.txt`  
4. Execute o script principal (ou use `iniciar.bat` / empacotamento PyInstaller conforme `automacao_assinatura_lotes.spec`).

---

## Como usar (visão geral)

1. **Inicie** `automacao_assinatura_lotes.py` (ou o executável gerado).  
2. Informe **usuário**, **senha do portal** e **senha do certificado digital** nos campos indicados.  
3. Defina **data inicial e final** da consulta de lotes.  
4. Escolha **tipo de publicação** (ex.: Notificação de Autuação, Multa, Cancelamento, etc.).  
5. Configure **forma de fiscalização** (Eletrônica, Manual, Remota) e **tipo/subtipo** conforme as regras do seu caso.  
6. Ajuste, se necessário, os **fatores de delay** (rede lenta ou sistema instável → aumente; ambiente rápido → pode reduzir com cautela).  
7. Inicie a automação e **acompanhe o log**; em caso de falha pontual, o sistema registra o motivo conforme a lógica implementada (ex.: guia de assinatura não abriu, timeout na barra de progresso, exceção no lote).

> **Importante:** mudanças no HTML dos portais da ANTT podem exigir atualização dos seletores (`id`, CSS). Este projeto mapeia IDs com prefixo `ContentPlaceHolderCorpo_` típicos de ASP.NET WebForms.

---

## Comportamento técnico (resumo)

- **URLs principais:** login SCA e tela de lotes DOU em `appweb1.antt.gov.br`.  
- **Fluxo:** login → filtros → pesquisa → tabela de lotes → para cada linha: abrir assinatura (nova guia/aba) → senha certificado → salvar → aguardar confirmação → retornar à lista → eventualmente avançar página.  
- **Delays:** constantes configuráveis (`DELAYS`) com fatores multiplicadores na GUI para “senha/cliques” e demais etapas.  
- **Erros tratados como falha de lote:** entre outros, nova guia não abre no tempo esperado, barra de progresso não conclui no timeout, exceções não previstas durante o processamento.

---

## Estrutura de arquivos

| Arquivo | Função |
|---------|--------|
| `automacao_assinatura_lotes.py` | Código principal (GUI + automação Selenium) |
| `requirements.txt` | Dependências |
| `automacao_assinatura_lotes.spec` | Configuração PyInstaller para `.exe` |
| `build_exe.bat` | Script auxiliar de build |
| `chromedriver.exe` | Driver local (se presente; versão deve combinar com o Chrome) |
| `iniciar.bat` | Atalho para executar o script no Windows |

---

## Segurança e compliance

- **Não commite** credenciais, senhas de certificado ou logs com dados pessoais em repositórios públicos.  
- O uso desta automação deve estar **alinhado às políticas da ANTT** e ao **uso aceitável** dos sistemas institucionais.  
- Trate planilhas e logs como **dados sensíveis** quando contiverem informações de infrações ou identificação.

---

## Solução de problemas

| Sintoma | O que verificar |
|---------|-----------------|
| Elemento não encontrado | Atualização do site; conferir IDs no DevTools |
| Timeout frequente | Aumentar fatores de delay; checar rede e carga do portal |
| ChromeDriver | Versão compatível com a versão do Chrome instalada |
| Certificado | Validade, driver do token, senha correta |

---

## Licença e contexto

Uso interno / projeto **CCOBI – SERASA**. Ajuste a licença conforme a política da sua organização.
