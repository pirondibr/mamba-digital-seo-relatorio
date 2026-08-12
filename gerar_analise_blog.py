# -*- coding: utf-8 -*-
"""Cria aba Analise Blog: total de paginas + % por topico."""
from __future__ import annotations

from collections import Counter
from pathlib import Path
from urllib.parse import urlparse

import pandas as pd

PATHS = [
    Path(r"C:\Users\Usuario\Desktop\Seo Conteudo\mamba-digital-seo-topicos.xlsx"),
    Path(
        r"C:\Users\Usuario\Desktop\Seo Conteudo\mambadigital.com.br-organic.Positions-br-20260811-2026-08-12T17_22_42Z.xlsx"
    ),
]


def url_path(u: str) -> str:
    p = urlparse(str(u)).path.rstrip("/").lower()
    return p or "/"


def assign_topic(path: str, topics: list[str]) -> str:
    # Comparativos antes de ML/Shopee generico
    if any(
        x in path
        for x in [
            "mercado-livre-shopee-ou-amazon",
            "shopee-ou-mercado-livre",
            "amazon-ou-mercado-livre",
            "shopee-mercado-livre",
            "maiores-marketplaces",
        ]
    ):
        return "Comparativo de Marketplaces"
    if "/category/noticias" in path:
        return "Conteudo Geral"
    if "shopee" in path:
        return "Shopee"
    if "mercado-livre" in path or "mercadolivre" in path or "mercadolider" in path:
        return "Mercado Livre"
    if any(
        x in path
        for x in [
            "full-funnel",
            "erp",
            "concorrente",
            "conversao",
            "social-commerce",
            "cliente-comprar",
        ]
    ):
        return "Marketing Digital"
    if any(x in path for x in ["capital-de-giro", "precificacao"]):
        return "Gestao e Financas"
    if any(x in path for x in ["black-friday", "cyber", "dia-das-criancas"]):
        return "Promocoes e Datas Comerciais"
    if any(
        x in path
        for x in ["codigo-ean", "fulfillment", "produtos-mais-vendidos", "roi-marketplace", "como-vender-em-marketplace"]
    ):
        return "Operacao Marketplace"
    if path == "/blog" or "/category/" in path:
        return "Marca Mamba"
    c = Counter([t for t in topics if t and t != "Ruido SERP"])
    return c.most_common(1)[0][0] if c else "Outros"


def build_analise_blog(kw: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    blog = kw[
        kw["URL"].astype(str).str.contains("mambadigital.com.br/blog", case=False, na=False)
    ].copy()
    blog["path"] = blog["URL"].map(url_path)

    pages = []
    for path, g in blog.groupby("path"):
        topico = assign_topic(path, g["Topico"].tolist())
        best = g.sort_values("Traffic", ascending=False).iloc[0]
        pages.append(
            {
                "URL": str(best["URL"]),
                "Path": path,
                "Topico": topico,
                "Keywords_ranking": int(len(g)),
                "Trafego_total": int(g["Traffic"].sum()),
                "Tipo_pagina": "Categoria" if ("/category/" in path or path == "/blog") else "Artigo",
            }
        )

    pages_df = (
        pd.DataFrame(pages)
        .sort_values(["Topico", "Trafego_total"], ascending=[True, False])
        .reset_index(drop=True)
    )
    total = len(pages_df)

    dist = (
        pages_df.groupby("Topico", as_index=False)
        .agg(
            Paginas=("Path", "count"),
            Keywords_ranking=("Keywords_ranking", "sum"),
            Trafego_total=("Trafego_total", "sum"),
        )
        .sort_values("Paginas", ascending=False)
        .reset_index(drop=True)
    )
    dist["Pct_paginas"] = (dist["Paginas"] / total * 100).round(2)
    dist = dist[["Topico", "Paginas", "Pct_paginas", "Keywords_ranking", "Trafego_total"]]

    resumo = pd.DataFrame(
        [
            {"Metrica": "Total de paginas do blog identificadas", "Valor": total},
            {"Metrica": "Artigos", "Valor": int((pages_df["Tipo_pagina"] == "Artigo").sum())},
            {"Metrica": "Categorias / index", "Valor": int((pages_df["Tipo_pagina"] == "Categoria").sum())},
            {"Metrica": "Topicos cobertos", "Valor": int(dist["Topico"].nunique())},
        ]
    )
    return resumo, dist, pages_df


def compose_sheet(resumo: pd.DataFrame, dist: pd.DataFrame, pages: pd.DataFrame) -> pd.DataFrame:
    """Uma unica tabela plana para HTML/Excel legivel."""
    rows: list[dict] = []
    rows.append({"Secao": "RESUMO", "Topico": "Metrica", "Paginas": "Valor", "Pct_paginas": "", "Keywords_ranking": "", "Trafego_total": "", "URL": "", "Tipo_pagina": ""})
    for _, r in resumo.iterrows():
        rows.append(
            {
                "Secao": "RESUMO",
                "Topico": r["Metrica"],
                "Paginas": r["Valor"],
                "Pct_paginas": "",
                "Keywords_ranking": "",
                "Trafego_total": "",
                "URL": "",
                "Tipo_pagina": "",
            }
        )
    rows.append({"Secao": "", "Topico": "", "Paginas": "", "Pct_paginas": "", "Keywords_ranking": "", "Trafego_total": "", "URL": "", "Tipo_pagina": ""})
    rows.append(
        {
            "Secao": "DISTRIBUICAO",
            "Topico": "Topico",
            "Paginas": "Paginas",
            "Pct_paginas": "Pct_paginas",
            "Keywords_ranking": "Keywords_ranking",
            "Trafego_total": "Trafego_total",
            "URL": "",
            "Tipo_pagina": "",
        }
    )
    for _, r in dist.iterrows():
        rows.append(
            {
                "Secao": "DISTRIBUICAO",
                "Topico": r["Topico"],
                "Paginas": r["Paginas"],
                "Pct_paginas": r["Pct_paginas"],
                "Keywords_ranking": r["Keywords_ranking"],
                "Trafego_total": r["Trafego_total"],
                "URL": "",
                "Tipo_pagina": "",
            }
        )
    rows.append({"Secao": "", "Topico": "", "Paginas": "", "Pct_paginas": "", "Keywords_ranking": "", "Trafego_total": "", "URL": "", "Tipo_pagina": ""})
    rows.append(
        {
            "Secao": "PAGINAS",
            "Topico": "Topico",
            "Paginas": "",
            "Pct_paginas": "",
            "Keywords_ranking": "Keywords_ranking",
            "Trafego_total": "Trafego_total",
            "URL": "URL",
            "Tipo_pagina": "Tipo_pagina",
        }
    )
    for _, r in pages.iterrows():
        rows.append(
            {
                "Secao": "PAGINAS",
                "Topico": r["Topico"],
                "Paginas": "",
                "Pct_paginas": "",
                "Keywords_ranking": r["Keywords_ranking"],
                "Trafego_total": r["Trafego_total"],
                "URL": r["URL"],
                "Tipo_pagina": r["Tipo_pagina"],
            }
        )
    return pd.DataFrame(rows)


def save_workbook(path: Path, sheet_df: pd.DataFrame, dist: pd.DataFrame, pages: pd.DataFrame, resumo: pd.DataFrame) -> Path:
    try:
        existing = pd.read_excel(path, sheet_name=None)
    except PermissionError:
        alt = path.with_name(path.stem + "_atualizado.xlsx")
        existing = pd.read_excel(path, sheet_name=None) if False else {}
        # try read via copy name - if locked, read from other file
        raise

    existing["Analise Blog"] = sheet_df
    # tambem abas limpas auxiliares para uso analitico
    existing["Analise Blog Dist"] = dist
    existing["Analise Blog Paginas"] = pages

    try:
        with pd.ExcelWriter(path, engine="openpyxl") as writer:
            for name, df in existing.items():
                # nao duplicar auxiliares demais — user pediu uma aba.
                if name in {"Analise Blog Dist", "Analise Blog Paginas"}:
                    continue
                df.to_excel(writer, sheet_name=name, index=False)
            # garantir Analise Blog
            sheet_df.to_excel(writer, sheet_name="Analise Blog", index=False)
        return path
    except PermissionError:
        alt = path.with_name("mamba-digital-seo-topicos-com-blog.xlsx")
        # merge from readable sibling if needed
        other = PATHS[1] if path == PATHS[0] else PATHS[0]
        base = pd.read_excel(other if other.exists() else path, sheet_name=None)
        base["Analise Blog"] = sheet_df
        with pd.ExcelWriter(alt, engine="openpyxl") as writer:
            for name, df in base.items():
                df.to_excel(writer, sheet_name=name, index=False)
        print(f"[AVISO] Arquivo aberto. Salvo em: {alt}")
        return alt


def main() -> None:
    src = next(p for p in PATHS if p.exists())
    # Prefer unlocked read
    kw = pd.read_excel(src, sheet_name="Keywords")
    resumo, dist, pages = build_analise_blog(kw)
    sheet_df = compose_sheet(resumo, dist, pages)

    print("=== RESUMO ===")
    print(resumo.to_string(index=False))
    print("\n=== DISTRIBUICAO ===")
    print(dist.to_string(index=False))

    saved = []
    for p in PATHS:
        if not p.exists():
            continue
        try:
            existing = pd.read_excel(p, sheet_name=None)
            existing["Analise Blog"] = sheet_df
            with pd.ExcelWriter(p, engine="openpyxl") as writer:
                for name, df in existing.items():
                    df.to_excel(writer, sheet_name=name, index=False)
            print("OK:", p)
            saved.append(p)
        except PermissionError:
            alt = Path(r"C:\Users\Usuario\Desktop\Seo Conteudo\mamba-digital-seo-topicos-com-blog.xlsx")
            # rebuild from last successful or src
            base_path = saved[0] if saved else src
            try:
                base = pd.read_excel(base_path, sheet_name=None)
            except Exception:
                base = {"Keywords": kw}
            base["Analise Blog"] = sheet_df
            # try keep other sheets from locked file via already-read if any
            with pd.ExcelWriter(alt, engine="openpyxl") as writer:
                for name, df in base.items():
                    df.to_excel(writer, sheet_name=name, index=False)
            print("[AVISO] Salvo copia:", alt)
            saved.append(alt)

    # export json-friendly tables for HTML update
    out_dir = Path(r"C:\Users\Usuario\Desktop\Seo Conteudo")
    resumo.to_csv(out_dir / "_blog_resumo.csv", index=False)
    dist.to_csv(out_dir / "_blog_dist.csv", index=False)
    pages.to_csv(out_dir / "_blog_pages.csv", index=False)
    print("CSV auxiliares salvos.")


if __name__ == "__main__":
    main()
