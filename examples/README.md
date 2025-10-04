# Exemplos de arquivos

Este diretório reúne um conjunto mínimo de arquivos para testar o Emaileria via linha de comando, assistente (`emaileria_wizard.py`) ou interface gráfica (`gui.py`).

## Conteúdo

- `leads_exemplo.xlsx`: planilha com três contatos fictícios e colunas extras (`produto`, `validade`, `cidade`) para demonstrar placeholders Jinja2.
- `assunto_exemplo.txt`: template de assunto usando o nome e a cidade do lead.
- `corpo_exemplo.html`: template HTML com placeholders para todas as colunas da planilha.

## Como usar

1. Abra o aplicativo desejado e selecione a planilha `leads_exemplo.xlsx` quando for solicitado.
2. Informe o remetente e as credenciais SMTP (pode ser um envio em modo `--dry-run` para testes).
3. Escolha os arquivos de template conforme necessário:
   - Assunto: `assunto_exemplo.txt`
   - Corpo: `corpo_exemplo.html`

Os placeholders serão substituídos automaticamente pelos valores de cada linha, permitindo validar o fluxo de ponta a ponta antes de usar dados reais.
