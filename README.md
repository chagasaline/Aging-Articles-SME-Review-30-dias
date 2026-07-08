# KB Mailer — Aging Articles

Sistema web que substitui a macro VBA `MandaEmail` da planilha "Cobrança + 30".
Faz upload da planilha, agrupa os artigos por dono e **envia os e-mails direto por
SMTP autenticado, sem precisar abrir o Outlook.

## O que ele reproduz da macro original

- Mesmo assunto: `KB Mgmt. - Aging Articles (SME Review more than 30 days)`
- Mesmo texto PT + tabela + EN
- Mesma tabela HTML (DocId / Type / Article Title / Article Owner Name), com as
  mesmas cores
- Mesmo Cc fixo: ` emails das pessoas separados por ponto e vírgulas; `
- Remetente: ` email do remetente`

**Melhoria em relação à macro:** o agrupamento por dono não depende mais das
linhas estarem "coladas" na planilha (a macro original quebrava o agrupamento se
uma linha do mesmo dono não estivesse logo em seguida da anterior). Agora agrupa
por e-mail em toda a planilha.

## 1. Pré-requisitos no Office 365 / Exchange Online

Como o envio é autenticado (usuário + senha), você precisa de **uma** destas duas opções:

**Opção A — usar a própria caixa `IT.Knowledge.Management@vale.com`**
- A caixa precisa ter uma senha própria (se for uma caixa compartilhada, o TI
  precisa habilitar login direto nela, ou criar uma senha).
- SMTP AUTH precisa estar habilitado para essa caixa no Exchange Admin Center
  (a Microsoft desabilita por padrão em tenants novos — peça ao time de
  infraestrutura de e-mail para habilitar `Authenticated SMTP` para essa conta).
- Se a conta tiver MFA, será necessário gerar uma **senha de aplicativo** (App
  Password) para usar no SMTP, já que SMTP AUTH básico não suporta MFA
  interativo.

**Opção B — usar sua própria conta autenticada, enviando "como" a caixa**
- Sua conta faz login no SMTP normalmente.
- O TI concede a você a permissão **"Send As"** (ou "Send on Behalf") na caixa
  `email` (via Exchange Admin Center ou PowerShell:
  `Add-RecipientPermission`).
- Sem essa permissão, o servidor rejeita o e-mail ou troca o remetente
  automaticamente.

Fale com o time responsável pelo Exchange Online da Vale para validar qual das
duas opções está disponível.

## 2. Instalação

```bash
cd kb-mailer
python3 -m venv venv
source venv/bin/activate      # Windows: venv\Scripts\activate
pip install -r requirements.txt
cp .env.example .env
```

Edite o `.env` com as credenciais reais (host/porta SMTP, usuário, senha,
remetente, lista de Cc).

## 3. Rodar

```bash
python app.py
```

Acesse **http://localhost:5000** no navegador.

## 4. Usar

1. Suba a planilha `.xlsx`/`.xlsm` (mesmas colunas de hoje: Doc ID, Type,
   Article Title, Article Owner Name, Owner E mail).
2. Confira a pré-visualização — quantos destinatários únicos, quantos artigos
   cada um tem.
3. Marque/desmarque quem deve receber.
4. **Sempre teste primeiro:**
   - Deixe "Modo simulação (dry-run)" marcado para ver o resultado sem
     enviar nada de verdade, **ou**
   - Preencha "E-mail de teste" com o seu próprio e-mail para receber uma
     cópia real de como ficaria, sem mandar para os donos de verdade.
5. Quando estiver confiante, desmarque o dry-run, apague o e-mail de teste e
   envie para valer. O sistema pede uma confirmação antes de disparar.

## 5. Rodando isso "para sempre" (sem precisar abrir terminal toda vez)

Para virar um sistema web de verdade acessível pela equipe (não só localhost),
seria necessário publicá-lo em um servidor interno da Vale (IIS/Windows Server,
uma VM Linux, um App Service do Azure, etc.) com HTTPS e, idealmente, um login
simples antes da tela de envio. Posso ajudar a preparar isso (Dockerfile,
deploy no Azure App Service, etc.) — é só pedir.

## Segurança

- O `.env` contém a senha do e-mail: nunca o coloque em um repositório Git
  público nem o compartilhe.
- Considere restringir o acesso a este sistema (rede interna / VPN da Vale,
  login).
