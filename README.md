# Emaileria

Ferramenta de linha de comando para enviar e-mails personalizados via Gmail a partir de uma planilha Excel.

## Pré-requisitos

1. **Python 3.10+** instalado.
2. Criar um projeto no [Google Cloud Console](https://console.cloud.google.com/) e habilitar a API do Gmail.
3. Configurar uma credencial do tipo *Desktop* e baixar o arquivo `credentials.json` para a raiz do projeto.
4. Instalar as dependências:

   ```bash
   python -m venv .venv
   source .venv/bin/activate
   pip install -r requirements.txt
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
  --subject-template "Plano funerário especial para {{ nome }}" \
  --body-template "$(cat template.html)"
```

Na primeira execução será aberta uma janela do navegador solicitando que autorize o acesso à sua conta Gmail. O token de acesso será salvo no arquivo `token.json`.

### Opções adicionais

- `--sheet`: define o nome da aba da planilha (caso não seja a primeira).
- `--credentials`: caminho alternativo para o `credentials.json`.
- `--token`: caminho alternativo para o `token.json`.
- `--dry-run`: apenas renderiza as mensagens sem enviá-las.
- `--log-level`: ajusta o nível de log (padrão: `INFO`).

## Segurança

- Nunca compartilhe o conteúdo de `credentials.json` ou `token.json`.
- Guarde a planilha com os dados sensíveis em local seguro.

## Licença

Distribuído sob a licença MIT. Veja `LICENSE` para mais detalhes.
