# -*- coding: utf-8 -*-
"""
Analise de concorrencia para keywords no Top 5 (exceto Marca Mamba).
Query: \"keyword\" -instagram -youtube -tiktok
"""
from __future__ import annotations

import os
import sys
import time
import unicodedata
from pathlib import Path

import pandas as pd

SCRIPT_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(SCRIPT_DIR))


def _load_scrapingbee_key() -> str:
    env_key = os.getenv("SCRAPINGBEE_API_KEY", "").strip()
    if env_key:
        return env_key
    for p in [
        Path(r"C:\Users\Usuario\Desktop\clinica cidade\Seo\.env"),
        SCRIPT_DIR / ".env",
    ]:
        if not p.exists():
            continue
        for line in p.read_text(encoding="utf-8").splitlines():
            if line.strip().startswith("SCRAPINGBEE_API_KEY="):
                return line.split("=", 1)[1].strip().strip('"').strip("'")
    return ""


os.environ["SCRAPINGBEE_API_KEY"] = _load_scrapingbee_key()
import seo_site_results as seo  # noqa: E402

seo.SCRAPINGBEE_API_KEY = os.environ["SCRAPINGBEE_API_KEY"]
seo.REQUEST_DELAY = 1.1

XLSX_CANDIDATES = [
    SCRIPT_DIR / "mamba-digital-seo-topicos.xlsx",
    SCRIPT_DIR / "mamba-digital-seo-topicos-com-blog.xlsx",
    SCRIPT_DIR
    / "mambadigital.com.br-organic.Positions-br-20260811-2026-08-12T17_22_42Z.xlsx",
]
CACHE_CSV = SCRIPT_DIR / "_concorrencia_top5_cache.csv"
SHEET_NAME = "Analise Concorrencia"


def norm_kw(s: str) -> str:
    s = str(s).lower().strip()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(c for c in s if not unicodedata.combining(c))
    return " ".join(s.split())


def competition_level(n: int | None) -> str:
    if n is None:
        return "Indefinido"
    if n <= 1000:
        return "Baixa"
    if n <= 50000:
        return "Media"
    if n <= 500000:
        return "Alta"
    return "Muito Alta"


def search_level(volume: int | float | None) -> str:
    """Nivel de buscas com base no Search Volume (Semrush)."""
    try:
        n = float(volume)
    except (TypeError, ValueError):
        return "Indefinido"
    if n <= 500:
        return "Baixa"
    if n <= 2000:
        return "Media"
    if n <= 10000:
        return "Alta"
    return "Muito Alta"


def build_query(keyword: str) -> str:
    # aspas na keyword + exclusao de redes sociais
    kw = str(keyword).strip().replace('"', "")
    return f'"{kw}" -instagram -youtube -tiktok'


def pick_keywords(df: pd.DataFrame, limit: int = 20) -> pd.DataFrame:
    base = df[
        (df["Position"] <= 5)
        & (~df["Topico"].isin(["Marca Mamba", "Ruido SERP"]))
    ].copy()
    base["_nk"] = base["Keyword"].map(norm_kw)
    # 1 linha por keyword normalizada (maior trafego, depois melhor posicao)
    base = base.sort_values(["Traffic", "Position"], ascending=[False, True])
    uniq = base.drop_duplicates(subset=["_nk"], keep="first").copy()
    uniq = uniq.drop(columns=["_nk"]).reset_index(drop=True)
    if limit and limit > 0:
        uniq = uniq.head(limit).copy()
    return uniq.reset_index(drop=True)


def load_cache() -> dict[str, int | None]:
    if not CACHE_CSV.exists():
        return {}
    c = pd.read_csv(CACHE_CSV)
    out = {}
    for _, r in c.iterrows():
        out[str(r["Keyword"])] = None if pd.isna(r["Resultados"]) else int(r["Resultados"])
    return out


def save_cache(rows: list[dict]) -> None:
    pd.DataFrame(rows)[["Keyword", "Query", "Resultados"]].drop_duplicates(
        subset=["Keyword"], keep="last"
    ).to_csv(CACHE_CSV, index=False)


def analyze(selected: pd.DataFrame) -> pd.DataFrame:
    cache = load_cache()
    rows: list[dict] = []
    total = len(selected)
    for i, (_, r) in enumerate(selected.iterrows(), start=1):
        kw = str(r["Keyword"])
        query = build_query(kw)
        print(f"[{i}/{total}] {query}")

        if kw in cache:
            count = cache[kw]
            print(f"  [cache] {count}")
        else:
            try:
                count = seo.get_number_of_results(query)
            except Exception as e:
                print(f"  [ERRO] {e}")
                count = None
            cache[kw] = count
            time.sleep(seo.REQUEST_DELAY)

        sv = None
        if "Search Volume" in r.index and pd.notna(r.get("Search Volume")):
            sv = int(r["Search Volume"])
        rows.append(
            {
                "Keyword": kw,
                "Topico": r["Topico"],
                "Subtopico": r["Subtopico"],
                "Position": int(r["Position"]),
                "Traffic": int(r["Traffic"]),
                "Search Volume": sv,
                "Nivel_buscas": search_level(sv if sv is not None else r["Traffic"]),
                "URL": r["URL"],
                "Query": query,
                "Resultados": count,
                "Nivel_concorrencia": competition_level(count),
            }
        )

        # checkpoint a cada 15
        if i % 15 == 0:
            save_cache(
                [{"Keyword": k, "Query": build_query(k), "Resultados": v} for k, v in cache.items()]
            )

    save_cache([{"Keyword": k, "Query": build_query(k), "Resultados": v} for k, v in cache.items()])

    out = pd.DataFrame(rows)
    # ordena: concorrencia crescente (oportunidade), depois trafego
    out["_sort"] = out["Resultados"].fillna(10**12)
    out = out.sort_values(["_sort", "Traffic"], ascending=[True, False]).drop(columns=["_sort"])
    cols = [
        "Keyword",
        "Topico",
        "Subtopico",
        "Position",
        "Traffic",
        "Search Volume",
        "Nivel_buscas",
        "Resultados",
        "Nivel_concorrencia",
        "URL",
        "Query",
    ]
    cols = [c for c in cols if c in out.columns] + [c for c in out.columns if c not in cols]
    return out[cols].reset_index(drop=True)


def write_sheet(path: Path, df: pd.DataFrame) -> None:
    existing = pd.read_excel(path, sheet_name=None)
    existing[SHEET_NAME] = df
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        for name, sheet in existing.items():
            sheet.to_excel(writer, sheet_name=name, index=False)


def main() -> None:
    if not seo.SCRAPINGBEE_API_KEY:
        raise SystemExit("SCRAPINGBEE_API_KEY nao encontrada")

    src = next(p for p in XLSX_CANDIDATES if p.exists())
    print(f"Lendo: {src}")
    kw = pd.read_excel(src, sheet_name="Keywords")
    selected = pick_keywords(kw, limit=20)
    print(f"Keywords Top5 (sem Marca/Ruido) — amostra: {len(selected)}")

    result = analyze(selected)
    result.to_csv(SCRIPT_DIR / "_analise_concorrencia.csv", index=False)

    print("\n=== Distribuicao Nivel ===")
    print(result["Nivel_concorrencia"].value_counts().to_string())
    print("\n=== Top oportunidades (menos resultados) ===")
    print(
        result.head(15)[
            ["Keyword", "Topico", "Position", "Traffic", "Resultados", "Nivel_concorrencia"]
        ].to_string(index=False)
    )

    for p in XLSX_CANDIDATES:
        if not p.exists():
            continue
        try:
            write_sheet(p, result)
            print(f"Aba salva: {p.name}")
        except PermissionError:
            alt = p.with_name(p.stem + "_concorrencia.xlsx")
            # tenta montar a partir do src
            try:
                base = pd.read_excel(src, sheet_name=None)
            except Exception:
                base = {"Keywords": kw}
            base[SHEET_NAME] = result
            with pd.ExcelWriter(alt, engine="openpyxl") as writer:
                for name, sheet in base.items():
                    sheet.to_excel(writer, sheet_name=name, index=False)
            print(f"[AVISO] Aberto. Salvo em: {alt}")


if __name__ == "__main__":
    main()
