# -*- coding: utf-8 -*-
"""Gera aba Topicos (resumo visual por topico)."""
from __future__ import annotations

from pathlib import Path
from urllib.parse import urlparse

import pandas as pd

SCRIPT_DIR = Path(__file__).resolve().parent
PATHS = [
    SCRIPT_DIR / "mamba-digital-seo-topicos.xlsx",
    SCRIPT_DIR / "mamba-digital-seo-topicos-com-blog.xlsx",
    SCRIPT_DIR / "mamba-digital-seo-topicos-com-blog_concorrencia.xlsx",
    SCRIPT_DIR
    / "mambadigital.com.br-organic.Positions-br-20260811-2026-08-12T17_22_42Z.xlsx",
]


def url_path(u: str) -> str:
    return (urlparse(str(u)).path or "/").rstrip("/").lower() or "/"


def build_topicos(kw: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for topico, g in kw.groupby("Topico"):
        if topico == "Ruido SERP":
            continue
        blog = g[g["URL"].astype(str).str.contains("/blog/", case=False, na=False)]
        pages = blog["URL"].map(url_path).nunique()
        t1 = g[g["Position"] <= 1]
        t5 = g[g["Position"] <= 5]
        t10 = g[g["Position"] <= 10]
        rows.append(
            {
                "Topico": topico,
                "Keywords": int(len(g)),
                "Subtopicos": int(g["Subtopico"].nunique()),
                "Paginas_blog": int(pages),
                "Search_Volume_total": int(g["Search Volume"].fillna(0).sum())
                if "Search Volume" in g.columns
                else 0,
                "Trafego_total": int(g["Traffic"].fillna(0).sum()),
                "Keywords_Top1": int(len(t1)),
                "Trafego_Top1": int(t1["Traffic"].fillna(0).sum()),
                "Keywords_Top5": int(len(t5)),
                "Trafego_Top5": int(t5["Traffic"].fillna(0).sum()),
                "Keywords_Top10": int(len(t10)),
                "Trafego_Top10": int(t10["Traffic"].fillna(0).sum()),
            }
        )
    out = pd.DataFrame(rows).sort_values("Trafego_total", ascending=False).reset_index(drop=True)
    total_kw = out["Keywords"].sum() or 1
    total_pages = out["Paginas_blog"].sum() or 1
    out["Pct_keywords"] = (out["Keywords"] / total_kw * 100).round(1)
    out["Pct_paginas_blog"] = (out["Paginas_blog"] / total_pages * 100).round(1)
    return out


def main() -> None:
    src = next(p for p in PATHS if p.exists())
    kw = pd.read_excel(src, sheet_name="Keywords")
    topicos = build_topicos(kw)
    print(topicos.to_string(index=False))
    topicos.to_csv(SCRIPT_DIR / "_topicos_resumo.csv", index=False)

    for p in PATHS:
        if not p.exists():
            continue
        try:
            existing = pd.read_excel(p, sheet_name=None)
            # monta ordem: Keywords, Topicos, resto
            ordered = {}
            if "Keywords" in existing:
                ordered["Keywords"] = existing["Keywords"]
            ordered["Topicos"] = topicos
            for name, df in existing.items():
                if name in {"Keywords", "Topicos"}:
                    continue
                ordered[name] = df
            with pd.ExcelWriter(p, engine="openpyxl") as writer:
                for name, df in ordered.items():
                    df.to_excel(writer, sheet_name=name, index=False)
            print("OK", p.name)
        except PermissionError:
            print("LOCKED", p.name)


if __name__ == "__main__":
    main()
