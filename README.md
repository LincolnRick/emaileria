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
python email_sender.py leads.xlsx --sender "seu-email@gmail.com" \
  --smtp-user "seu-email@gmail.com" --smtp-password "sua-senha-de-app" \
  --subject-template "Plano funerário especial para {{ nome }}" \
  --body-template-file template.html --dry-run
```

O exemplo acima funciona igualmente em shells Windows (PowerShell ou CMD),
bastando remover as barras invertidas de continuação de linha caso prefira
uma única linha:

```powershell
python email_sender.py leads.xlsx --sender "seu-email@gmail.com" --smtp-user "seu-email@gmail.com" `
  --smtp-password "sua-senha-de-app" --subject-template "Plano funerário especial para {{ nome }}" `
  --body-template-file template.html --dry-run
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

## Assistente interativo (wizard)

Usuários que preferem uma experiência guiada podem executar o assistente interativo em modo texto:

```bash
python emaileria_wizard.py
```

O assistente faz perguntas passo a passo sobre os arquivos envolvidos no envio, renderiza prévias e executa o disparo das mensagens ao final. Também é possível gerar um executável standalone (`emaileria-wizard` ou `emaileria-wizard.exe`) seguindo as instruções da seção [Gerando executável](#gerando-executável).

## Interface gráfica (GUI)

A interface gráfica dispensa qualquer interação com a linha de comando. Depois de instalar as dependências (ver seção [Pré-requisitos](#pré-requisitos)), siga estes passos:

1. (Opcional) Ative o ambiente virtual que você criou anteriormente.
2. Na raiz do projeto, execute:

   ```bash
   python gui.py
   ```

   No Windows, é possível dar dois cliques em `gui.py` caso a associação com Python esteja configurada. O terminal será aberto automaticamente e mostrará os logs da aplicação.

### Preenchendo os campos

A janela apresenta campos simples para serem preenchidos na ordem sugerida abaixo:

1. **Planilha (XLSX/CSV)** – escolha a planilha com os contatos. Você pode usar o exemplo `examples/readme/leads.csv` para testar o fluxo.
2. **Aba (sheet)** – informe o nome da aba caso não esteja usando a primeira aba da planilha (deixe em branco para usar a padrão).
3. **Remetente (From)** – endereço de e-mail que aparecerá como remetente.
4. **SMTP User** – usuário para autenticação SMTP. Normalmente é o mesmo do remetente.
5. **SMTP Password** – senha de app do Gmail. O campo esconde os caracteres digitados. Deixe em branco para que o envio use a variável de ambiente `SMTP_PASSWORD` (caso definida) ou seja solicitado durante o processo.
6. **Assunto (Jinja2)** – template do assunto. Use `examples/readme/template_assunto.txt` como referência.
7. **Template HTML** – selecione o arquivo HTML com o corpo do e-mail (por exemplo, `examples/readme/template_corpo.html`). O conteúdo é lido na hora do envio, portanto você pode editar o arquivo e reenviar sem reiniciar a GUI.
8. **Dry-run (não enviar, apenas pré-visualizar)** – marcado por padrão para simular o envio sem disparar mensagens reais. Desmarque quando estiver seguro.
9. **Log level** – ajuste o detalhamento dos logs exibidos na área inferior.

Clique em **Enviar** para iniciar o processamento. A parte inferior da janela funciona como um terminal de log, mostrando o andamento do envio (ou da pré-visualização). Use o botão **Sair** para encerrar a aplicação com segurança.

### Fluxo sem linha de comando

1. Abra a GUI (`python gui.py`).
2. Selecione a planilha e os templates desejados.
3. Revise os campos de autenticação SMTP.
4. Escolha se deseja executar um dry-run ou enviar de fato.
5. Clique em **Enviar** e acompanhe os logs até a conclusão. Qualquer erro será destacado em vermelho.

## Gerando executável da GUI

### Empacotando a GUI com PyInstaller

Execute os comandos abaixo a partir da raiz do repositório (com o ambiente virtual ativo e as dependências instaladas):

- **Windows**

  ```bash
  pyinstaller --onefile --noconsole gui.py
  ```

  O executável ficará em `dist/gui.exe`.

- **macOS (Intel/ARM)**

  ```bash
  pyinstaller --onefile --windowed gui.py
  ```

  O aplicativo ficará em `dist/gui`.

Após a build, copie a planilha e os templates (por exemplo, os de `examples/readme/`) para a mesma pasta do executável ou para qualquer local acessível pelo seletor de arquivos da GUI.

## Segurança

- Prefira utilizar senhas de app em vez da senha principal da conta Gmail.
- Guarde a planilha com os dados sensíveis em local seguro.
- Nunca informe a senha SMTP diretamente na linha de comando. Utilize o campo **SMTP Password** da GUI ou defina a variável de ambiente `SMTP_PASSWORD` antes de iniciar os scripts.

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

Os binários gerados ficarão disponíveis em `dist/emaileria` e `dist/emaileria-wizard` (ou com extensão `.exe` no Windows).
O primeiro corresponde ao envio automático via linha de comando, enquanto o segundo empacota o assistente interativo `emaileria_wizard.py`.

## Licença

Distribuído sob a licença MIT. Veja `LICENSE` para mais detalhes.
