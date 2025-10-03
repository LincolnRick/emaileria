# Emaileria

Ferramenta de linha de comando para enviar e-mails personalizados via Gmail a partir de uma planilha Excel.

## Pré-requisitos

1. **Python 3.10+** instalado.
2. Uma conta Gmail com [verificação em duas etapas](https://myaccount.google.com/security) habilitada e uma [senha de app](https://support.google.com/accounts/answer/185833) gerada especificamente para o envio automatizado.
3. Instalar as dependências:

   ```bash
   python -m venv .venv
   source .venv/bin/activate
   pip install -r requirements.txt
   ```

## Como configurar .env

1. Copie o arquivo `.env.example` para `.env` na raiz do projeto.
2. Preencha as variáveis com as credenciais e remetente desejado.
3. Garanta que `RATE_LIMIT_PER_MINUTE` reflita o limite máximo de envios por minuto aceito pelo seu provedor SMTP.

As variáveis padrão estão configuradas para uso com servidores SMTP do Gmail via SSL (`smtp.gmail.com:465`).

## Como formatar

Para aplicar as ferramentas de lint e formatação configuradas no projeto, execute:

```bash
ruff check .
black .
```

## Preparando a planilha

A planilha deve conter, pelo menos, as seguintes colunas (sem distinção entre maiúsculas/minúsculas):

- `email`: endereço de destino.
- `tratamento`: forma de tratamento para o lead (ex.: "Sr.", "Sra.").
- `nome`: nome do destinatário.

Qualquer outra coluna presente na planilha poderá ser utilizada dentro dos templates de assunto e corpo.

## Templates de e-mail

Tanto o assunto quanto o corpo aceitam placeholders [Jinja2](https://jinja.palletsprojects.com/) com os nomes das colunas da planilha. Exemplo de corpo HTML em um arquivo `template.html`:

```html
<p>Olá {{ tratamento }} {{ nome }},</p>
<p>Gostaríamos de apresentar o nosso plano funerário com cobertura nacional...</p>
<p>Atenciosamente,<br>Equipe Exemplo</p>
```

Assunto de exemplo: `Plano funerário especial para {{ nome }}`.

## Execução

```bash
python email_sender.py leads.xlsx \
  --sender "seu-email@gmail.com" \
  --subject-template-file template_assunto.txt \
  --body-template-file template.html
```

Também é possível fornecer os conteúdos diretamente em linha, como no exemplo abaixo:

```bash
python email_sender.py leads.xlsx \
  --sender "seu-email@gmail.com" \
  --smtp-user "seu-email@gmail.com" \
  --smtp-password "sua-senha-de-app" \
  --subject-template "Plano funerário especial para {{ nome }}" \
  --body-template "$(cat template.html)"
```

Ao executar o comando, o script solicitará a senha SMTP (recomenda-se usar a senha de app). Você também pode informar as credenciais via linha de comando, como mostrado acima. As mensagens são enviadas utilizando `smtplib.SMTP_SSL` diretamente contra `smtp.gmail.com:465`, portanto não é necessário configurar nenhum projeto no Google Cloud.

### Opções adicionais

- `--sheet`: define o nome da aba da planilha (caso não seja a primeira).
- `--smtp-user`: usuário SMTP a ser autenticado (por padrão é o mesmo informado em `--sender`).
- `--smtp-password`: senha ou senha de app a ser utilizada (se omitido, será solicitado via prompt seguro).
- `--subject-template`: template do assunto do e-mail.
- `--subject-template-file`: caminho para um arquivo contendo o template do assunto (sobrescreve `--subject-template`).
- `--body-template`: template do corpo do e-mail (HTML).
- `--body-template-file`: caminho para um arquivo contendo o template do corpo (sobrescreve `--body-template`).
- `--dry-run`: apenas renderiza as mensagens sem enviá-las.
- `--log-level`: ajusta o nível de log (padrão: `INFO`).

## Interface gráfica (GUI)

Instale dependências (incluindo PySimpleGUI):

```bash
pip install -r requirements.txt
```

Rode a interface:

```bash
python gui.py
```

Preencha os campos e clique Enviar.
O log mostrará exatamente o que email_sender.py está fazendo (suporta --dry-run).

### Gerar executável com PyInstaller

**Windows**

```bash
pyinstaller --onefile --noconsole gui.py
```

Saída: `dist/gui.exe`

Ícone opcional: `--icon assets/emaileria.ico`

**macOS (Intel/ARM)**

```bash
pyinstaller --onefile --windowed gui.py
```

Saída: `dist/gui`

Dica: templates HTML podem ficar numa pasta `templates/` ao lado do executável. Como o `gui.py` lê o arquivo escolhido pelo usuário, não é obrigatório empacotar assets.

## Segurança

- Prefira utilizar senhas de app em vez da senha principal da conta Gmail.
- Guarde a planilha com os dados sensíveis em local seguro.

## Gerando executável

1. Crie (ou reutilize) um ambiente virtual e instale as dependências de execução e de build:

   ```bash
   python -m venv .venv
   source .venv/bin/activate
   pip install -r requirements.txt
   pip install -r requirements-build.txt
   ```

2. Execute o script de build:

   ```bash
   ./scripts/build_executable.sh
   ```

O binário gerado ficará disponível em `dist/emaileria` (ou `dist/emaileria.exe` no Windows).

## Licença

Distribuído sob a licença MIT. Veja `LICENSE` para mais detalhes.
