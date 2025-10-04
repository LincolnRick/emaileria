# Emaileria Examples

This directory contains small, self-contained examples that demonstrate how to
work with **Emaileria**.

## Dry-run example

The [`send_messages_dry_run.py`](./send_messages_dry_run.py) script shows how to
render messages for a list of contacts without actually sending any email. It
uses the sample CSV file and templates stored in `examples/readme/`.

Run it from the project root with:

```bash
python examples/send_messages_dry_run.py
```

The script will render each message and print the generated subject so that you
can validate the template placeholders. No messages are delivered because the
`dry_run` mode is enabled.
