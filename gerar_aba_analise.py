# -*- coding: utf-8 -*-
"""
Gera aba 'Analise' no XLSX da Mamba:
- 1 keyword por Topico unico (maior trafego)
- Top 5 Google organico + Mamba Digital sempre na posicao 6
- site:dominio | site:dominio keyword | densidade (%)
"""
from __future__ import annotations

import os
import sys
import time
from pathlib import Path
from typing import Any
from urllib.parse import urlparse

import pandas as pd

# Reutiliza funcoes do seo_site_results
SCRIPT_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(SCRIPT_DIR))

# Carrega API key do .env da clinica (mesmo padrao do projeto)
def _load_scrapingbee_key() -> str:
    env_key = os.getenv("SCRAPINGBEE_API_KEY", "").strip()
    if env_key:
        return env_key
    candidates = [
        Path(r"C:\Users\Usuario\Desktop\clinica cidade\Seo\.env"),
        SCRIPT_DIR / ".env",
    ]
    for p in candidates:
        if not p.exists():
            continue
        for line in p.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if line.startswith("SCRAPINGBEE_API_KEY="):
                return line.split("=", 1)[1].strip().strip('"').strip("'")
    return ""


os.environ["SCRAPINGBEE_API_KEY"] = _load_scrapingbee_key()

import seo_site_results as seo  # noqa: E402

seo.SCRAPINGBEE_API_KEY = os.environ["SCRAPINGBEE_API_KEY"]
seo.REQUEST_DELAY = 1.2

MAMBA_DOMAIN = "mambadigital.com.br"
MAMBA_URL = "https://mambadigital.com.br/"

XLSX_PATHS = [
    SCRIPT_DIR / "mamba-digital-seo-topicos.xlsx",
    SCRIPT_DIR
    / "mambadigital.com.br-organic.Positions-br-20260811-2026-08-12T17_22_42Z.xlsx",
]

# Topicos que nao fazem sentido para analise competitiva de conteudo
SKIP_TOPICOS = {"Ruido SERP", "Conteudo Geral"}


def pick_keywords(df: pd.DataFrame) -> pd.DataFrame:
    base = df[~df["Topico"].isin(SKIP_TOPICOS)].copy()
    idx = base.groupby("Topico")["Traffic"].idxmax()
    sel = base.loc[idx, ["Topico", "Keyword", "Traffic", "Position", "URL"]].copy()
    return sel.sort_values("Traffic", ascending=False).reset_index(drop=True)


def analyze_keyword(topico: str, keyword: str, traffic: int) -> list[dict[str, Any]]:
    print("\n" + "=" * 70)
    print(f"Topico: {topico}")
    print(f"Keyword: {keyword} (trafego={traffic})")
    print("=" * 70)

    print(f"[1/3] Top 5 organico Google (excluindo {MAMBA_DOMAIN})")
    competitors = seo.get_top_organic_domains(keyword, top_n=5, exclude=MAMBA_DOMAIN)
    time.sleep(seo.REQUEST_DELAY)

    companies: list[dict[str, Any]] = []
    for i, c in enumerate(competitors, start=1):
        companies.append(
            {
                "posicao": i,
                "empresa": c["domain"],
                "tipo": "google_top",
                "titulo_serp": c.get("title") or "",
                "url_serp": c.get("url") or "",
                "posicao_google": c.get("position"),
            }
        )

    # Completa ate 5 slots se SERP trouxe menos
    while len(companies) < 5:
        companies.append(
            {
                "posicao": len(companies) + 1,
                "empresa": "",
                "tipo": "google_top",
                "titulo_serp": "",
                "url_serp": "",
                "posicao_google": None,
            }
        )

    # Sempre Top 6 = Mamba Digital
    companies.append(
        {
            "posicao": 6,
            "empresa": MAMBA_DOMAIN,
            "tipo": "mamba",
            "titulo_serp": "Mamba Digital | A maior assessoria de marketplace da América Latina",
            "url_serp": MAMBA_URL,
            "posicao_google": None,
        }
    )

    print("[2/3] Contando site:dominio + keyword")
    print("[3/3] Contando site:dominio (total indexado)")

    rows: list[dict[str, Any]] = []
    for company in companies:
        domain = company["empresa"]
        if not domain:
            rows.append(
                {
                    "Topico": topico,
                    "Keyword": keyword,
                    "Trafego_keyword": traffic,
                    "Posicao": company["posicao"],
                    "Empresa": "",
                    "Tipo": company["tipo"],
                    "Posicao_Google": company["posicao_google"],
                    "Titulo_SERP": "",
                    "URL_SERP": "",
                    "Query_keyword": "",
                    "Resultados_keyword": None,
                    "Query_total": "",
                    "Resultados_total": None,
                    "Densidade_pct": None,
                }
            )
            continue

        query_kw = f"site:{domain} {keyword}"
        query_total = f"site:{domain}"

        print(f"  ({company['posicao']}/6) {query_kw}")
        try:
            count_kw = seo.get_number_of_results(query_kw)
        except Exception as e:
            print(f"  [ERRO] {e}")
            count_kw = None
        time.sleep(seo.REQUEST_DELAY)

        print(f"  ({company['posicao']}/6) {query_total}")
        count_total = seo.get_site_total(domain)
        time.sleep(seo.REQUEST_DELAY)

        dens = seo.specialization_pct(count_kw, count_total)
        rows.append(
            {
                "Topico": topico,
                "Keyword": keyword,
                "Trafego_keyword": traffic,
                "Posicao": company["posicao"],
                "Empresa": domain,
                "Tipo": company["tipo"],
                "Posicao_Google": company["posicao_google"],
                "Titulo_SERP": company["titulo_serp"],
                "URL_SERP": company["url_serp"],
                "Query_keyword": query_kw,
                "Resultados_keyword": count_kw,
                "Query_total": query_total,
                "Resultados_total": count_total,
                "Densidade_pct": dens,
            }
        )

    return rows


def write_analise_sheet(path: Path, df_analise: pd.DataFrame) -> None:
    # Preserva abas existentes
    existing = pd.read_excel(path, sheet_name=None)
    existing["Analise"] = df_analise
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        for name, sheet_df in existing.items():
            sheet_df.to_excel(writer, sheet_name=name, index=False)


def main() -> None:
    if not seo.SCRAPINGBEE_API_KEY:
        raise SystemExit("SCRAPINGBEE_API_KEY nao encontrada.")

    src = XLSX_PATHS[0]
    if not src.exists():
        src = XLSX_PATHS[1]
    print(f"Lendo: {src}")

    kw_df = pd.read_excel(src, sheet_name="Keywords")
    selected = pick_keywords(kw_df)
    print("\nKeywords selecionadas (1 por topico):")
    print(selected.to_string(index=False))

    seo._SITE_TOTAL_CACHE.clear()
    all_rows: list[dict[str, Any]] = []
    for _, row in selected.iterrows():
        all_rows.extend(
            analyze_keyword(
                str(row["Topico"]),
                str(row["Keyword"]),
                int(row["Traffic"] or 0),
            )
        )

    df_analise = pd.DataFrame(all_rows)

    for path in XLSX_PATHS:
        if not path.exists():
            continue
        try:
            write_analise_sheet(path, df_analise)
            print(f"Aba Analise salva em: {path}")
        except PermissionError:
            alt = path.with_name(path.stem + "_com_analise.xlsx")
            write_analise_sheet(alt, df_analise) if False else None
            # salva copia se original estiver aberto
            with pd.ExcelWriter(alt, engine="openpyxl") as writer:
                for name, sheet_df in pd.read_excel(path, sheet_name=None).items():
                    if name == "Analise":
                        continue
                    sheet_df.to_excel(writer, sheet_name=name, index=False)
                df_analise.to_excel(writer, sheet_name="Analise", index=False)
            print(f"[AVISO] Arquivo aberto. Salvo em: {alt}")

    print("\n=== Resumo Analise ===")
    pd.set_option("display.width", 200)
    pd.set_option("display.max_columns", 15)
    print(
        df_analise[
            ["Topico", "Keyword", "Posicao", "Empresa", "Resultados_keyword", "Resultados_total", "Densidade_pct"]
        ].to_string(index=False)
    )
    print(f"\nLinhas: {len(df_analise)}")


if __name__ == "__main__":
    main()
