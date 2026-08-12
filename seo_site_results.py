"""
SEO Site Results Analyzer
-------------------------
Para cada linha do input (site + keyword):
  1. Busca o termo no Google e pega o top 5 orgânico
  2. Para a empresa do input e cada concorrente do top 5, conta:
     a) site:dominio keyword  → páginas sobre o tema (especialista)
     b) site:dominio          → total indexado do site (generalista)
  3. Calcula % de especialização = keyword / total
  4. Busca no Semrush: tráfego de Marca (branded) + tráfego recente (total)
  5. Exporta XLSX com os dados

Usa ScrapingBee Google Search API + Semrush DPA (config1.json).
"""

from __future__ import annotations

import argparse
import json
import os
import random
import re
import time
from datetime import datetime
from string import ascii_lowercase, digits
from typing import Any
from urllib.parse import urlparse

import pandas as pd
import requests

# ============================================================================
# CONFIG
# ============================================================================

SCRAPINGBEE_API_KEY = os.getenv("SCRAPINGBEE_API_KEY", "").strip()
SCRAPINGBEE_GOOGLE_URL = "https://app.scrapingbee.com/api/v1/store/google"
SEMRUSH_RPC_URL = "https://www.semrush.com/dpa/rpc"

# Brasil / português por padrão
COUNTRY_CODE = "br"
LANGUAGE = "pt"
SEMRUSH_DATABASE = "br"

# Delay entre requests (segundos) para não estourar rate limit
REQUEST_DELAY = 1.0
MAX_RETRIES = 3

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_INPUT = os.path.join(SCRIPT_DIR, "input_exemplo.csv")
DEFAULT_OUTPUT_DIR = os.path.join(SCRIPT_DIR, "output")
SEMRUSH_CONFIG_PATH = os.path.join(SCRIPT_DIR, "config1.json")


# ============================================================================
# HELPERS
# ============================================================================

def normalize_domain(value: str) -> str:
    """Extrai domínio limpo de URL ou string site:..."""
    if not value:
        return ""
    value = value.strip().lower()
    value = re.sub(r"^site:", "", value, flags=re.I)
    if "://" not in value:
        value = "https://" + value
    parsed = urlparse(value)
    host = (parsed.netloc or parsed.path).split("/")[0]
    host = host.removeprefix("www.")
    return host.strip(".")


def load_input(filepath: str) -> list[dict[str, str]]:
    """
    Aceita CSV com colunas site,keyword
    ou TXT com linhas: site|keyword  /  site;keyword  /  site,keyword
    """
    rows: list[dict[str, str]] = []
    ext = os.path.splitext(filepath)[1].lower()

    if ext in {".csv", ".xlsx", ".xls"}:
        df = pd.read_csv(filepath) if ext == ".csv" else pd.read_excel(filepath)
        cols = {c.lower().strip(): c for c in df.columns}
        site_col = cols.get("site") or cols.get("dominio") or cols.get("domain") or cols.get("url")
        kw_col = cols.get("keyword") or cols.get("keywords") or cols.get("palavra") or cols.get("palavra_chave")
        if not site_col or not kw_col:
            raise ValueError(
                "Input precisa ter colunas 'site' e 'keyword' "
                f"(encontradas: {list(df.columns)})"
            )
        for _, row in df.iterrows():
            site = normalize_domain(str(row[site_col]))
            keyword = str(row[kw_col]).strip()
            if site and keyword and keyword.lower() != "nan":
                rows.append({"site": site, "keyword": keyword})
        return rows

    with open(filepath, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            if "|" in line:
                parts = line.split("|", 1)
            elif ";" in line:
                parts = line.split(";", 1)
            else:
                parts = line.split(",", 1)
            if len(parts) != 2:
                print(f"[WARN] Linha ignorada (formato inválido): {line}")
                continue
            site = normalize_domain(parts[0])
            keyword = parts[1].strip()
            if site and keyword:
                rows.append({"site": site, "keyword": keyword})
    return rows


def parse_result_count_from_html(html: str) -> int | None:
    """
    Fallback: Google esconde o total em #result-stats / texto
    tipo 'Aproximadamente 1.234 resultados' / 'About 1,234 results'.
    """
    if not html:
        return None

    # resultStats id (desktop clássico)
    m = re.search(
        r'id=["\']result-stats["\'][^>]*>([^<]+)',
        html,
        flags=re.I,
    )
    text = m.group(1) if m else ""

    if not text:
        m2 = re.search(
            r"(?:Aproximadamente|About)\s+([\d\.\,]+)\s+result",
            html,
            flags=re.I,
        )
        if m2:
            text = m2.group(0)

    if not text:
        return None

    nums = re.findall(r"[\d\.\,]+", text)
    if not nums:
        return None

    raw = nums[0]
    # BR: 1.234.567  |  EN: 1,234,567
    if "." in raw and "," in raw:
        # ex: 1.234,5 -> remove milhar ., decimal ,
        raw = raw.replace(".", "").replace(",", ".")
    elif raw.count(".") > 1 or (raw.count(".") == 1 and len(raw.split(".")[-1]) == 3):
        raw = raw.replace(".", "")
    elif raw.count(",") > 1 or (raw.count(",") == 1 and len(raw.split(",")[-1]) == 3):
        raw = raw.replace(",", "")
    else:
        raw = raw.replace(",", "")

    try:
        return int(float(raw))
    except ValueError:
        return None


# ============================================================================
# SCRAPINGBEE GOOGLE API
# ============================================================================

def google_search(
    query: str,
    *,
    add_html: bool = False,
    light_request: bool = True,
) -> dict[str, Any]:
    """Chama ScrapingBee Google Search API e retorna JSON."""
    params = {
        "api_key": SCRAPINGBEE_API_KEY,
        "search": query,
        "country_code": COUNTRY_CODE,
        "language": LANGUAGE,
        "light_request": str(light_request).lower(),
        "nb_results": "10",
    }
    if add_html:
        params["add_html"] = "true"

    last_err: Exception | None = None
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            r = requests.get(SCRAPINGBEE_GOOGLE_URL, params=params, timeout=90)
            if r.status_code == 200:
                return r.json()
            # créditos / bloqueio temporário
            print(f"[API] HTTP {r.status_code} tentativa {attempt}/{MAX_RETRIES}: {r.text[:200]}")
            if r.status_code in {429, 500, 503}:
                time.sleep(2 * attempt)
                continue
            r.raise_for_status()
        except Exception as e:
            last_err = e
            print(f"[API] Erro tentativa {attempt}/{MAX_RETRIES}: {e}")
            time.sleep(2 * attempt)

    raise RuntimeError(f"Falha na busca Google: {query!r} ({last_err})")


def get_number_of_results(query: str) -> int:
    """
    Retorna o total estimado de resultados para a query.

    Fonte principal: meta_data.number_of_results (ScrapingBee já parseia
    o texto que o Google mostra / esconde no HTML).

    Fallback: add_html=true + parse de #result-stats.
    """
    data = google_search(query, add_html=False, light_request=True)
    meta = data.get("meta_data") or {}
    count = meta.get("number_of_results")

    if isinstance(count, int) and count >= 0:
        return count
    if isinstance(count, str) and count.isdigit():
        return int(count)

    # Fallback com HTML completo
    print(f"[FALLBACK] number_of_results ausente para: {query} — tentando HTML")
    data = google_search(query, add_html=True, light_request=False)
    meta = data.get("meta_data") or {}
    count = meta.get("number_of_results")
    if isinstance(count, int) and count >= 0:
        return count

    html = data.get("html") or data.get("body") or ""
    parsed = parse_result_count_from_html(html)
    if parsed is not None:
        return parsed

    # Zero resultados reais vs falha de parse
    organic = data.get("organic_results") or []
    return len(organic) if organic else 0


def get_top_organic_domains(keyword: str, top_n: int = 5, exclude: str | None = None) -> list[dict[str, Any]]:
    """Busca o keyword no Google e retorna até top_n domínios orgânicos únicos."""
    data = google_search(keyword, add_html=False, light_request=True)
    organic = data.get("organic_results") or []

    seen: set[str] = set()
    results: list[dict[str, Any]] = []
    exclude_norm = normalize_domain(exclude) if exclude else ""

    for item in organic:
        url = item.get("url") or item.get("link") or ""
        domain = normalize_domain(item.get("domain") or url)
        if not domain:
            continue
        if domain == exclude_norm:
            continue
        if domain in seen:
            continue
        seen.add(domain)
        results.append(
            {
                "domain": domain,
                "position": item.get("position") or item.get("rank") or len(results) + 1,
                "title": item.get("title") or "",
                "url": url,
            }
        )
        if len(results) >= top_n:
            break

    return results


# ============================================================================
# SEMRUSH — Marca + Tráfego recente
# ============================================================================

_SEMRUSH_CONFIG: dict[str, Any] | None = None
_SEMRUSH_USER_IDX = 0
_SEMRUSH_CACHE: dict[str, dict[str, Any]] = {}


def load_semrush_config() -> dict[str, Any] | None:
    global _SEMRUSH_CONFIG
    if _SEMRUSH_CONFIG is not None:
        return _SEMRUSH_CONFIG
    try:
        with open(SEMRUSH_CONFIG_PATH, encoding="utf-8") as f:
            _SEMRUSH_CONFIG = json.load(f)
            return _SEMRUSH_CONFIG
    except FileNotFoundError:
        print(f"[SEMRUSH] config1.json não encontrado em: {SEMRUSH_CONFIG_PATH}")
        return None
    except Exception as e:
        print(f"[SEMRUSH] Erro ao carregar config: {e}")
        return None


def _next_semrush_user(config: dict[str, Any]) -> dict[str, Any]:
    global _SEMRUSH_USER_IDX
    users = config.get("users") or []
    if not users:
        raise RuntimeError("Nenhum user Semrush em config1.json")
    user = users[_SEMRUSH_USER_IDX % len(users)]
    _SEMRUSH_USER_IDX += 1
    return user


def get_semrush_traffic(domain: str) -> dict[str, Any]:
    """
    Retorna tráfego Semrush (database BR):
      - marca: trafficBranded (último daily)
      - trafego_recente: traffic total (último daily)

    A API DPA exige o batch MonthlyTrend + DailyTrend (igual ao scraper existente).
    """
    domain = normalize_domain(domain)
    if domain in _SEMRUSH_CACHE:
        cached = _SEMRUSH_CACHE[domain]
        print(f"  [cache Semrush] {domain}: marca={cached.get('marca')} trafego={cached.get('trafego_recente')}")
        return cached

    empty = {"marca": None, "trafego_recente": None, "trafego_non_branded": None}
    config = load_semrush_config()
    if not config:
        _SEMRUSH_CACHE[domain] = empty
        return empty

    try:
        user = _next_semrush_user(config)
        req_id = "".join(random.choices(ascii_lowercase + digits, k=33))
        req_id = f"{req_id[:8]}-{req_id[8:12]}-{req_id[12:16]}-{req_id[16:]}"

        common_params = {
            "request_id": req_id,
            "report": "organic.overview",
            "args": {
                "database": SEMRUSH_DATABASE,
                "searchItem": domain,
                "searchType": "domain",
                "filter": {},
            },
            "userId": user["userId"],
            "apiKey": user["apiKey"],
        }

        # Batch obrigatório: Monthly + Daily (Daily sozinho retorna Invalid Request)
        payload = [
            {
                "id": 7,
                "jsonrpc": "2.0",
                "method": "organic.MonthlyTrend",
                "params": common_params,
            },
            {
                "id": 8,
                "jsonrpc": "2.0",
                "method": "organic.DailyTrend",
                "params": common_params,
            },
        ]

        res = requests.post(SEMRUSH_RPC_URL, json=payload, timeout=30)
        res.raise_for_status()
        data = res.json()

        if not isinstance(data, list) or len(data) < 2:
            print(f"  [SEMRUSH] resposta inválida para {domain}: {type(data)}")
            _SEMRUSH_CACHE[domain] = empty
            return empty

        if data[0].get("error"):
            err = data[0]["error"].get("message", "Unknown")
            print(f"  [SEMRUSH] erro {domain}: {err}")
            _SEMRUSH_CACHE[domain] = empty
            return empty

        daily = data[1].get("result") or []
        if not daily:
            # fallback: último ponto mensal
            monthly = data[0].get("result") or []
            if not monthly:
                print(f"  [SEMRUSH] sem dados para {domain}")
                _SEMRUSH_CACHE[domain] = empty
                return empty
            latest = max(monthly, key=lambda x: x.get("date", 0))
        else:
            latest = max(daily, key=lambda x: x.get("date", 0))

        out = {
            "marca": latest.get("trafficBranded"),
            "trafego_recente": latest.get("traffic"),
            "trafego_non_branded": latest.get("trafficNonBranded"),
        }
        print(
            f"  [SEMRUSH] {domain}: marca={out['marca']} "
            f"trafego_recente={out['trafego_recente']}"
        )
        _SEMRUSH_CACHE[domain] = out
        return out
    except Exception as e:
        print(f"  [SEMRUSH] falha {domain}: {e}")
        _SEMRUSH_CACHE[domain] = empty
        return empty


# ============================================================================
# PIPELINE
# ============================================================================

# Cache de site:dominio (total indexado não muda por keyword)
_SITE_TOTAL_CACHE: dict[str, int | None] = {}


def specialization_pct(keyword_count: int | None, total_count: int | None) -> float | None:
    """% de páginas do site que batem com a keyword (especialização)."""
    if keyword_count is None or total_count is None:
        return None
    if total_count <= 0:
        return 0.0
    return round((keyword_count / total_count) * 100, 2)


def get_site_total(domain: str) -> int | None:
    """Conta site:dominio (total indexado). Usa cache entre keywords."""
    if domain in _SITE_TOTAL_CACHE:
        print(f"  [cache] site:{domain} = {_SITE_TOTAL_CACHE[domain]}")
        return _SITE_TOTAL_CACHE[domain]
    query = f"site:{domain}"
    try:
        count = get_number_of_results(query)
    except Exception as e:
        print(f"  [ERRO] {query}: {e}")
        count = None
    _SITE_TOTAL_CACHE[domain] = count
    return count


def analyze_row(site: str, keyword: str, top_n: int = 5) -> list[dict[str, Any]]:
    """
    Processa um par site+keyword e retorna linhas de output:
    - empresa input + top 5 Google
    - resultados_keyword / resultados_total / pct_especializacao
    - marca (Semrush branded) + trafego_recente (Semrush total)
    """
    print("\n" + "=" * 70)
    print(f"Keyword: {keyword}")
    print(f"Empresa input: {site}")
    print("=" * 70)

    print(f"[1/4] Buscando top {top_n} orgânico para: {keyword}")
    competitors = get_top_organic_domains(keyword, top_n=top_n, exclude=site)
    time.sleep(REQUEST_DELAY)

    companies: list[dict[str, Any]] = [
        {"empresa": site, "tipo": "input", "posicao_google": None, "titulo": "", "url": ""}
    ]
    for c in competitors:
        companies.append(
            {
                "empresa": c["domain"],
                "tipo": "google_top",
                "posicao_google": c["position"],
                "titulo": c["title"],
                "url": c["url"],
            }
        )

    print(f"[2/4] Contando site:dominio + keyword...")
    print(f"[3/4] Contando site:dominio total...")
    print(f"[4/4] Semrush: Marca + Tráfego recente...")
    rows: list[dict[str, Any]] = []

    for i, company in enumerate(companies, start=1):
        domain = company["empresa"]
        query_keyword = f"site:{domain} {keyword}"
        query_total = f"site:{domain}"

        print(f"  ({i}/{len(companies)}) {query_keyword}")
        try:
            count_keyword = get_number_of_results(query_keyword)
        except Exception as e:
            print(f"  [ERRO] {e}")
            count_keyword = None
        time.sleep(REQUEST_DELAY)

        print(f"  ({i}/{len(companies)}) {query_total}")
        count_total = get_site_total(domain)
        time.sleep(REQUEST_DELAY)

        print(f"  ({i}/{len(companies)}) Semrush {domain}")
        semrush = get_semrush_traffic(domain)
        time.sleep(random.uniform(1.0, 2.0))

        pct = specialization_pct(count_keyword, count_total)
        marca = semrush.get("marca")
        trafego = semrush.get("trafego_recente")

        # Coluna combinada pedida: Marca + Tráfego recente
        if marca is None and trafego is None:
            marca_e_trafego = None
        else:
            marca_fmt = f"{int(marca):,}".replace(",", ".") if marca is not None else "—"
            trafego_fmt = f"{int(trafego):,}".replace(",", ".") if trafego is not None else "—"
            marca_e_trafego = f"Marca: {marca_fmt} | Tráfego: {trafego_fmt}"

        rows.append(
            {
                "keyword": keyword,
                "empresa_input": site,
                "empresa": domain,
                "tipo": company["tipo"],
                "posicao_google": company["posicao_google"],
                "titulo_serp": company["titulo"],
                "url_serp": company["url"],
                "query_keyword": query_keyword,
                "resultados_keyword": count_keyword,
                "query_total": query_total,
                "resultados_total": count_total,
                "pct_especializacao": pct,
                "marca": marca,
                "trafego_recente": trafego,
                "marca_e_trafego": marca_e_trafego,
            }
        )

    return rows


def run(input_path: str, output_dir: str, top_n: int = 5) -> str:
    if not SCRAPINGBEE_API_KEY:
        raise SystemExit(
            "Defina a variável de ambiente SCRAPINGBEE_API_KEY antes de rodar o script."
        )
    os.makedirs(output_dir, exist_ok=True)
    items = load_input(input_path)
    if not items:
        raise SystemExit(f"Nenhuma linha válida em: {input_path}")

    print(f"Input: {input_path}")
    print(f"Linhas: {len(items)}")
    print(f"Top N: {top_n}")
    print(f"Country/Lang: {COUNTRY_CODE}/{LANGUAGE}")

    _SITE_TOTAL_CACHE.clear()
    _SEMRUSH_CACHE.clear()
    all_rows: list[dict[str, Any]] = []
    for item in items:
        all_rows.extend(analyze_row(item["site"], item["keyword"], top_n=top_n))

    df = pd.DataFrame(all_rows)

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_xlsx = os.path.join(output_dir, f"seo_site_results_{stamp}.xlsx")

    with pd.ExcelWriter(out_xlsx, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="detalhado", index=False)

        if not df.empty:
            df_labeled = df.assign(
                label=df.apply(
                    lambda r: (
                        f"{r['empresa']} (input)"
                        if r["tipo"] == "input"
                        else f"{r['empresa']} (#{r['posicao_google']})"
                    ),
                    axis=1,
                )
            )

            # Resumo keyword: site:dominio keyword
            resumo_kw = (
                df_labeled.pivot_table(
                    index=["keyword", "empresa_input"],
                    columns="label",
                    values="resultados_keyword",
                    aggfunc="first",
                )
                .reset_index()
            )
            resumo_kw.to_excel(writer, sheet_name="resumo_keyword", index=False)

            # Resumo total: site:dominio
            resumo_total = (
                df_labeled.pivot_table(
                    index=["keyword", "empresa_input"],
                    columns="label",
                    values="resultados_total",
                    aggfunc="first",
                )
                .reset_index()
            )
            resumo_total.to_excel(writer, sheet_name="resumo_total", index=False)

            # Resumo % especialização
            resumo_pct = (
                df_labeled.pivot_table(
                    index=["keyword", "empresa_input"],
                    columns="label",
                    values="pct_especializacao",
                    aggfunc="first",
                )
                .reset_index()
            )
            resumo_pct.to_excel(writer, sheet_name="resumo_especializacao", index=False)

            if "marca" in df.columns:
                resumo_marca = (
                    df_labeled.pivot_table(
                        index=["keyword", "empresa_input"],
                        columns="label",
                        values="marca",
                        aggfunc="first",
                    )
                    .reset_index()
                )
                resumo_marca.to_excel(writer, sheet_name="resumo_marca", index=False)

            if "trafego_recente" in df.columns:
                resumo_trafego = (
                    df_labeled.pivot_table(
                        index=["keyword", "empresa_input"],
                        columns="label",
                        values="trafego_recente",
                        aggfunc="first",
                    )
                    .reset_index()
                )
                resumo_trafego.to_excel(writer, sheet_name="resumo_trafego", index=False)

    print("\n" + "=" * 70)
    print(f"Concluído! Arquivo: {out_xlsx}")
    print("=" * 70)
    return out_xlsx


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Conta site:keyword e site:total (especialista vs generalista) via ScrapingBee"
    )
    parser.add_argument(
        "-i",
        "--input",
        default=DEFAULT_INPUT,
        help="CSV/XLSX/TXT com site e keyword",
    )
    parser.add_argument(
        "-o",
        "--output-dir",
        default=DEFAULT_OUTPUT_DIR,
        help="Pasta de saída",
    )
    parser.add_argument(
        "-n",
        "--top",
        type=int,
        default=5,
        help="Quantidade de concorrentes do Google (default: 5)",
    )
    args = parser.parse_args()
    run(args.input, args.output_dir, top_n=args.top)


if __name__ == "__main__":
    main()
