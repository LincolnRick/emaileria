from __future__ import annotations

from datetime import datetime
from pathlib import Path
import html
import re


def _strip_scripts(html_text: str) -> str:
    """Remove scripts e event handlers potencialmente inseguros."""
    html_text = re.sub(r"(?is)<script.*?>.*?</script>", "", html_text)
    html_text = re.sub(r" on\w+\s*=\s*\".*?\"", "", html_text)
    html_text = re.sub(r" on\w+\s*=\s*\'.*?\'", "", html_text)
    return html_text


def build_preview_page(previews: list[dict], out_dir: Path | None = None) -> Path:
    """Cria uma página HTML com cartões e retorna o caminho do arquivo gerado."""
    if out_dir is None:
        out_dir = Path("previews") / datetime.now().strftime("%Y%m%d_%H%M%S")
    out_dir.mkdir(parents=True, exist_ok=True)
    index = out_dir / "index.html"

    cards: list[str] = []
    for preview in previews:
        subj = html.escape(preview.get("subject", ""))
        addr = html.escape(preview.get("email", ""))
        body = _strip_scripts(preview.get("body_html", "") or "")
        card = f"""
        <section class=\"card\">
          <div class=\"meta\">
            <div><strong>#{preview.get('idx', '')}</strong> — {addr}</div>
            <div class=\"subject\">{subj}</div>
          </div>
          <iframe srcdoc='{body.replace("'", "&apos;")}' sandbox=""></iframe>
        </section>
        """
        cards.append(card)

    html_doc = f"""<!doctype html>
<html lang=\"pt-br\">
<head>
  <meta charset=\"utf-8\">
  <title>Emaileria — Prévia</title>
  <meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">
  <style>
    :root {{
      --bg:#f6f7fb; --fg:#0f172a; --muted:#475569; --card:#fff; --line:#e2e8f0;
      --accent:#2563eb;
    }}
    body {{ margin:0; background:var(--bg); color:var(--fg); font:16px/1.4 system-ui, -apple-system, Segoe UI, Roboto, Arial, sans-serif; }}
    header {{ padding:16px 24px; background:#fff; border-bottom:1px solid var(--line); position:sticky; top:0; z-index:10; }}
    header h1 {{ margin:0; font-size:18px; }}
    .wrap {{ max-width:1080px; margin:24px auto; padding:0 24px; }}
    .grid {{ display:grid; grid-template-columns:1fr; gap:16px; }}
    .card {{ background:var(--card); border:1px solid var(--line); border-radius:12px; padding:12px; }}
    .card .meta {{ display:flex; justify-content:space-between; align-items:center; gap:12px; margin:4px 4px 10px; color:var(--muted); }}
    .card .meta .subject {{ color:var(--fg); font-weight:600; }}
    .card iframe {{
      width:100%; max-width:600px; height:520px; border:1px solid var(--line);
      border-radius:8px; display:block; background:#fff; margin:8px auto;
    }}
    .legend {{ color:var(--muted); margin-top:8px; font-size:14px; }}
  </style>
</head>
<body>
  <header><h1>Prévia — {len(previews)} mensagem(ns)</h1></header>
  <main class=\"wrap\">
    <p class=\"legend\">Visualização isolada: cada e-mail é renderizado em um <em>iframe</em> (largura de 600px), simulando o layout do cliente de e-mail.</p>
    <div class=\"grid\">
      {''.join(cards)}
    </div>
  </main>
</body>
</html>"""
    index.write_text(html_doc, encoding="utf-8")
    return index
