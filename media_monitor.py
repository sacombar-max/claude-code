#!/usr/bin/env python3
"""
Daily media monitoring routine.
Fetches news from configured RSS feeds, retrieves economic indicators,
and generates an executive summary via Claude API.
"""

import os
import sys
import datetime
import smtplib
import xml.etree.ElementTree as ET
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import requests
import yfinance as yf
import anthropic
from dotenv import load_dotenv

load_dotenv()

# ──────────────────────────────────────────────
# Feed sources grouped by category
# ──────────────────────────────────────────────
FEEDS = {
    "MACRO GLOBAL": [
        ("Reuters Business",    "https://feeds.reuters.com/reuters/businessNews"),
        ("Reuters World",       "https://feeds.reuters.com/reuters/worldNews"),
        ("IMF News",            "https://www.imf.org/en/News/rss?language=eng"),
    ],
    "MERCADOS": [
        ("WSJ Markets",         "https://feeds.a.dj.com/rss/RSSMarketsMain.xml"),
        ("Reuters Markets",     "https://feeds.reuters.com/reuters/marketsNews"),
    ],
    "LATINOAMÉRICA": [
        ("Bloomberg Línea",     "https://www.bloomberglinea.com/arc/outboundfeeds/rss/"),
        ("El País América",     "https://feeds.elpais.com/mrss-s/pages/ep/site/elpais.com/section/america/portada"),
    ],
    "COLOMBIA": [
        ("Portafolio",          "https://www.portafolio.co/rss_feed.xml"),
        ("La República",        "https://www.larepublica.co/rss.xml"),
        ("BanRep Noticias",     "https://www.banrep.gov.co/es/rss-noticias"),
        ("DANE Boletines",      "https://www.dane.gov.co/index.php/noticias/rss"),
    ],
    "GEOPOLÍTICA": [
        ("BBC World",           "https://feeds.bbci.co.uk/news/world/rss.xml"),
        ("AP News Top",         "https://feeds.apnews.com/rss/apf-topnews"),
        ("The Economist",       "https://www.economist.com/rss"),
    ],
}

# ──────────────────────────────────────────────
# Economic indicators (Yahoo Finance tickers)
# ──────────────────────────────────────────────
INDICATORS = {
    "Dólar (DXY)":   "DX-Y.NYB",
    "Euro/USD":      "EURUSD=X",
    "WTI Petróleo":  "CL=F",
    "Brent Petróleo":"BZ=F",
    "S&P 500":       "^GSPC",
    "TRM COP/USD":   "COP=X",
}


_NS = {
    "atom":    "http://www.w3.org/2005/Atom",
    "content": "http://purl.org/rss/1.0/modules/content/",
    "dc":      "http://purl.org/dc/elements/1.1/",
}

def _text(el):
    return (el.text or "").strip() if el is not None else ""


def fetch_feed(name, url, max_items=5):
    """Fetch and parse an RSS/Atom feed without external dependencies."""
    headers = {"User-Agent": "MediaMonitor/1.0 (daily news aggregator)"}
    try:
        resp = requests.get(url, headers=headers, timeout=10)
        resp.raise_for_status()
        root = ET.fromstring(resp.content)
    except Exception as exc:
        print(f"  [warn] {name}: {exc}", file=sys.stderr)
        return []

    items = []

    # Atom feeds
    atom_entries = root.findall("atom:entry", _NS) or root.findall("{http://www.w3.org/2005/Atom}entry")
    if atom_entries:
        for entry in atom_entries[:max_items]:
            ns = "http://www.w3.org/2005/Atom"
            title   = _text(entry.find(f"{{{ns}}}title"))
            summary = _text(entry.find(f"{{{ns}}}summary")) or _text(entry.find(f"{{{ns}}}content"))
            link_el = entry.find(f"{{{ns}}}link")
            link    = link_el.get("href", "") if link_el is not None else ""
            pub     = _text(entry.find(f"{{{ns}}}updated")) or _text(entry.find(f"{{{ns}}}published"))
            items.append({"source": name, "title": title, "summary": summary[:500], "link": link, "published": pub})
        return items

    # RSS 2.0 / RSS 1.0
    channel = root.find("channel") or root
    for item in list(channel.findall("item"))[:max_items]:
        title   = _text(item.find("title"))
        summary = _text(item.find("description")) or _text(item.find("{http://purl.org/rss/1.0/modules/content/}encoded"))
        link    = _text(item.find("link"))
        pub     = _text(item.find("pubDate")) or _text(item.find("{http://purl.org/dc/elements/1.1/}date"))
        items.append({"source": name, "title": title, "summary": summary[:500], "link": link, "published": pub})

    return items


def fetch_all_news():
    """Fetch all configured feeds grouped by category."""
    all_news = {}
    for category, feeds in FEEDS.items():
        items = []
        for name, url in feeds:
            items.extend(fetch_feed(name, url))
        all_news[category] = items
        print(f"  {category}: {len(items)} artículos")
    return all_news


def fetch_indicators():
    """Retrieve current values for economic indicators via yfinance."""
    results = {}
    for label, ticker in INDICATORS.items():
        try:
            data  = yf.Ticker(ticker)
            info  = data.fast_info
            price = getattr(info, "last_price", None) or getattr(info, "regularMarketPrice", None)
            prev  = getattr(info, "previous_close", None)
            change_pct = ((price - prev) / prev * 100) if price and prev else None
            results[label] = {
                "price":      round(price, 4) if price else "N/D",
                "change_pct": round(change_pct, 2) if change_pct else "N/D",
            }
        except Exception as exc:
            results[label] = {"price": "N/D", "change_pct": "N/D"}
            print(f"  [warn] {label}: {exc}", file=sys.stderr)
    return results


def build_news_text(all_news):
    lines = []
    for category, items in all_news.items():
        lines.append(f"\n### {category}")
        if not items:
            lines.append("  (sin artículos disponibles)")
            continue
        for item in items:
            lines.append(f"- [{item['source']}] {item['title']}")
            if item["summary"]:
                lines.append(f"  {item['summary'][:300]}")
    return "\n".join(lines)


def build_indicators_text(indicators):
    lines = []
    for label, data in indicators.items():
        price   = data["price"]
        chg     = data["change_pct"]
        arrow   = "▲" if isinstance(chg, float) and chg > 0 else ("▼" if isinstance(chg, float) and chg < 0 else "–")
        chg_str = f"{arrow} {abs(chg)}%" if isinstance(chg, float) else "N/D"
        lines.append(f"  {label:<20} {str(price):>12}   {chg_str}")
    return "\n".join(lines)


def generate_report(all_news, indicators):
    """Call Claude API to produce the executive summary."""
    api_key = os.getenv("ANTHROPIC_API_KEY")
    if not api_key:
        raise EnvironmentError("ANTHROPIC_API_KEY no está definida en el entorno.")

    client = anthropic.Anthropic(api_key=api_key)
    today  = datetime.date.today().strftime("%d de %B de %Y")

    news_text       = build_news_text(all_news)
    indicators_text = build_indicators_text(indicators)

    system_prompt = (
        "Eres un analista económico senior. "
        "Tu tarea es redactar un informe ejecutivo diario en español, "
        "claro, preciso y profesional, dirigido a ejecutivos y directivos empresariales."
    )

    user_prompt = f"""Con base en los siguientes titulares y resúmenes de prensa del {today},
genera un INFORME EJECUTIVO DIARIO con el siguiente formato exacto:

═══════════════════════════════════════════════════════
INFORME EJECUTIVO DIARIO – {today}
═══════════════════════════════════════════════════════

## NOTICIAS DESTACADAS

Selecciona las **3 noticias más relevantes** de toda la revisión de medios.
Para cada una escribe:
- **Título breve en negrita** (Fuente: nombre del medio)
  Un párrafo de máximo 4 líneas que explique el hecho, su contexto y su posible impacto.

## INDICADORES ECONÓMICOS CLAVE

Incluye la siguiente tabla con los valores proporcionados (no inventes cifras):

{indicators_text}

Agrega una línea de análisis de 2-3 oraciones sobre la lectura conjunta de los mercados.

═══════════════════════════════════════════════════════
Fin del informe
═══════════════════════════════════════════════════════

---
TITULARES RECOPILADOS (uso interno para el análisis):
{news_text}
"""

    message = client.messages.create(
        model="claude-opus-4-7",
        max_tokens=1500,
        system=system_prompt,
        messages=[{"role": "user", "content": user_prompt}],
    )
    return message.content[0].text


def save_report(report, output_dir="reports"):
    os.makedirs(output_dir, exist_ok=True)
    date_str  = datetime.date.today().strftime("%Y-%m-%d")
    file_path = os.path.join(output_dir, f"informe_{date_str}.txt")
    with open(file_path, "w", encoding="utf-8") as f:
        f.write(report)
    return file_path


def _report_to_html(report):
    lines      = report.splitlines()
    html_lines = ["<html><body style='font-family:Arial,sans-serif;max-width:700px;margin:auto;color:#222;'>"]
    for line in lines:
        s = line.strip()
        if s.startswith("═"):
            html_lines.append("<hr style='border:2px solid #1a3c6e;'>")
        elif s.startswith("## "):
            html_lines.append(f"<h2 style='color:#1a3c6e;border-bottom:1px solid #ccc;padding-bottom:4px;'>{s[3:]}</h2>")
        elif s.startswith("**") and s.endswith("**"):
            html_lines.append(f"<strong>{s[2:-2]}</strong><br>")
        elif s.startswith("**"):
            formatted = s.replace("**", "<strong>", 1).replace("**", "</strong>", 1)
            html_lines.append(f"<p style='margin:12px 0 4px 0;'>{formatted}</p>")
        elif s.startswith("Lectura de mercados:"):
            html_lines.append(f"<p style='background:#f0f4ff;padding:10px;border-left:4px solid #1a3c6e;'>{s}</p>")
        elif s.startswith("Fuentes consultadas:") or s.startswith("Generado automáticamente"):
            html_lines.append(f"<p style='font-size:11px;color:#888;'>{s}</p>")
        elif s == "":
            html_lines.append("<br>")
        else:
            html_lines.append(f"<p style='margin:4px 0;'>{s}</p>")
    html_lines.append("</body></html>")
    return "\n".join(html_lines)


def send_email(report, date_str):
    gmail_user     = os.getenv("GMAIL_USER")
    gmail_password = os.getenv("GMAIL_APP_PASSWORD")
    recipient      = os.getenv("GMAIL_RECIPIENT", gmail_user)

    if not gmail_user or not gmail_password:
        print("  [info] GMAIL_USER o GMAIL_APP_PASSWORD no definidos — se omite el envío por correo.")
        return

    subject = f"Informe Ejecutivo Diario – {date_str}"
    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"]    = f"Monitor de Medios <{gmail_user}>"
    msg["To"]      = recipient

    msg.attach(MIMEText(report, "plain", "utf-8"))
    msg.attach(MIMEText(_report_to_html(report), "html", "utf-8"))

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(gmail_user, gmail_password)
        server.sendmail(gmail_user, recipient, msg.as_string())

    print(f"  ✉️  Correo enviado a {recipient}")


def main():
    today    = datetime.date.today().strftime("%d/%m/%Y")
    date_str = datetime.date.today().strftime("%d de %B de %Y")
    print(f"\n{'═'*55}")
    print(f"  MONITOR DE MEDIOS  –  {today}")
    print(f"{'═'*55}\n")

    print("📡 Obteniendo noticias...")
    all_news = fetch_all_news()

    print("\n📈 Obteniendo indicadores económicos...")
    indicators = fetch_indicators()
    for label, data in indicators.items():
        print(f"  {label}: {data['price']}")

    print("\n🤖 Generando informe con Claude...")
    report = generate_report(all_news, indicators)

    path = save_report(report)
    print(f"\n✅ Informe guardado en: {path}\n")

    print("📧 Enviando por correo...")
    send_email(report, date_str)

    print(report)


if __name__ == "__main__":
    main()
