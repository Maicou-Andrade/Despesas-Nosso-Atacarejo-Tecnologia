# Configuração OAuth2 para Google Sheets

Para usar este extrator de despesas com OAuth2, você precisa configurar as credenciais do Google Cloud Console.

## Passo a Passo:

### 1. Acesse o Google Cloud Console
- Vá para: https://console.cloud.google.com/

### 2. Crie ou Selecione um Projeto
- Se não tiver um projeto, clique em "Criar Projeto"
- Dê um nome ao projeto (ex: "Extrator de Despesas")

### 3. Ative a Google Sheets API
- No menu lateral, vá em "APIs e Serviços" > "Biblioteca"
- Procure por "Google Sheets API"
- Clique em "Ativar"

### 4. Configure a Tela de Consentimento OAuth
- Vá em "APIs e Serviços" > "Tela de consentimento OAuth"
- Escolha "Externo" (para uso pessoal)
- Preencha as informações obrigatórias:
  - Nome do aplicativo: "Extrator de Despesas"
  - Email de suporte: seu email
  - Email do desenvolvedor: seu email

### 5. Crie Credenciais OAuth 2.0
- Vá em "APIs e Serviços" > "Credenciais"
- Clique em "Criar Credenciais" > "ID do cliente OAuth 2.0"
- Tipo de aplicativo: "Aplicativo para computador"
- Nome: "Extrator de Despesas Desktop"

### 6. Baixe o Arquivo de Credenciais
- Após criar, clique no ícone de download
- Renomeie o arquivo baixado para `credentials.json`
- Coloque o arquivo na pasta do projeto (mesmo local do app.py)

### 7. Execute o Aplicativo
- Execute: `streamlit run app.py`
- Na primeira vez, será aberto um navegador para autorizar o acesso
- Faça login com sua conta Google e autorize o aplicativo
- As credenciais serão salvas automaticamente em `token.pickle`

## Arquivos Importantes:
- `credentials.json` - Credenciais OAuth2 (não compartilhe!)
- `token.pickle` - Token de acesso salvo (criado automaticamente)

## Segurança:
- Nunca compartilhe o arquivo `credentials.json`
- Adicione estes arquivos ao `.gitignore` se usar controle de versão
- O token é renovado automaticamente quando expira

## Problemas Comuns:
1. **Erro 403**: Verifique se a Google Sheets API está ativada
2. **Arquivo não encontrado**: Certifique-se que `credentials.json` está na pasta correta
3. **Token expirado**: Delete `token.pickle` e execute novamente para reautorizar