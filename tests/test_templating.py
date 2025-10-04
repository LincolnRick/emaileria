from datetime import date, datetime

from emaileria.templating import render


def test_render_injects_default_dates() -> None:
    subject_template = "{{ data_envio }}"
    body_template = "{{ hora_envio }}"

    subject, body = render(subject_template, body_template, {})

    assert subject == date.today().strftime("%Y-%m-%d")
    datetime.strptime(body, "%H:%M")


def test_render_exposes_now_with_datefmt_filter() -> None:
    subject_template = ""
    body_template = "{{ now|datefmt('%Y-%m-%d') }}"

    _, body = render(subject_template, body_template, {})

    assert body == date.today().strftime("%Y-%m-%d")


def test_render_allows_context_to_override_globals() -> None:
    subject_template = "{{ data_envio }}"
    body_template = "{{ hora_envio }}"

    subject, body = render(
        subject_template,
        body_template,
        {"data_envio": "custom", "hora_envio": "12:34"},
    )

    assert subject == "custom"
    assert body == "12:34"
