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

A GUI permite configurar envios completos sem recorrer ao terminal. Depois de instalar as dependências (veja [Pré-requisitos](#pré-requisitos)), execute `python gui.py` na raiz do projeto. No Windows, também é possível abrir o arquivo com um duplo clique – o console exibirá os logs em tempo real.

### Passo a passo rápido

1. Selecione a planilha (**Planilha (XLSX/CSV)**). O seletor lista automaticamente os exemplos `examples/readme/leads_exemplo.xlsx` e outros arquivos recentes.
2. Caso esteja usando um Excel, escolha a aba desejada (**Aba (sheet)**). Para CSV o campo permanece desabilitado.
3. Preencha **Remetente (From)**, **SMTP User** e, se necessário, o campo **SMTP Password** (recomenda-se uma senha de app; o conteúdo não é exibido em lugar algum).
4. Informe o **Assunto (Jinja2)** e escolha o **Template HTML**. Os arquivos `examples/readme/assunto_exemplo.txt` e `examples/readme/corpo_exemplo.html` aparecem na lista de sugestões.
5. Defina os campos opcionais **CC**, **BCC** e **Reply-To** com listas separadas por vírgula.
6. Ajuste o comportamento do envio: mantenha **Dry-run** marcado para apenas renderizar os e-mails, escolha o **Log level** e utilize o **Intervalo entre envios (s)** para aplicar rate limit (0 a 2 segundos, padrão 0,75s).
7. Clique em **Validar & Prévia** para carregar a planilha, verificar placeholders obrigatórios e visualizar três amostras renderizadas (assunto + trecho do corpo). Somente após uma validação bem-sucedida o botão **Enviar** é habilitado.
8. Inicie o envio real (ou dry-run) com **Enviar**. A barra de progresso e o contador exibem a evolução, enquanto o botão **Cancelar** permite interromper com segurança.

### Recursos da janela

* **Validar & Prévia**: garante que as colunas `email`, `tratamento` e `nome` existam, sinaliza placeholders ausentes (considerando os globais `now`, `hoje`, `data_envio`, `hora_envio`) e abre um resumo scrollável com três mensagens de exemplo.
* **Campos CC/BCC/Reply-To**: aceitam listas separadas por vírgula com validação básica, repassadas diretamente para o envio SMTP.
* **Slider de intervalo**: define o tempo mínimo entre mensagens (0 a 2 segundos) e é respeitado inclusive nos reenvios via GUI.
* **Barra de progresso + botão Cancelar**: exibem contagem de enviados/total e permitem abortar o envio via flag thread-safe, sem travamentos.
* **Área de log integrada**: recebe todos os registros da aplicação (inclusive do módulo core) por meio de um `logging.Handler` dedicado, preservando a senha fora de qualquer resumo ou log.

### Preferências e dicas

* As escolhas mais recentes ficam salvas em `~/.emaileria_gui.json` (planilha, sheet, template HTML, remetente, usuário SMTP, CC/BCC/Reply-To, nível de log, estado do dry-run e valor do intervalo).
* Use os exemplos do diretório `examples/readme/` para testar rapidamente todo o fluxo.
* Senhas nunca são exibidas na tela nem persistidas em disco. Para maior segurança, utilize uma senha de app e a variável de ambiente `SMTP_PASSWORD` sempre que possível.

## Gerando executável (.exe/.app)

O projeto inclui uma especificação `emaileria.spec` pronta para o PyInstaller, bem como alvos no `Makefile` que simplificam o processo:

```bash
# Executa a GUI diretamente
make gui

# Gera binário Windows (console oculto)
make build-win

# Gera aplicativo macOS (modo windowed)
make build-mac
```

Os comandos acima utilizam `pyinstaller --clean --onefile` com base na configuração oficial. Os artefatos resultantes são gravados em `dist/` (por exemplo, `dist/emaileria.exe` no Windows).

Antes de distribuir o executável:

1. Crie/ative um ambiente virtual e instale as dependências de runtime e build (`pip install -r requirements.txt` e, se necessário, `pip install pyinstaller`).
2. Gere a build desejada com o alvo apropriado.
3. Empacote junto aos exemplos (`examples/readme/leads_exemplo.xlsx`, `examples/readme/assunto_exemplo.txt`, `examples/readme/corpo_exemplo.html`) ou disponibilize-os em uma pasta acessível ao usuário final.

### Boas práticas de segurança

* Utilize sempre senhas de app em vez da senha principal da conta.
* Não compartilhe planilhas contendo dados sensíveis fora de ambientes confiáveis.
* Evite fornecer a senha SMTP via linha de comando; prefira o campo dedicado na GUI ou a variável de ambiente `SMTP_PASSWORD`.

## Licença

Distribuído sob a licença MIT. Veja `LICENSE` para mais detalhes.
