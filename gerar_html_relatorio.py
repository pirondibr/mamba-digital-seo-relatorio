# -*- coding: utf-8 -*-
"""Gera HTML com abas a partir do XLSX Mamba SEO."""
from __future__ import annotations

import html
import json
from pathlib import Path

import pandas as pd

SRC_CANDIDATES = [
    Path(r"C:\Users\Usuario\Desktop\Seo Conteudo\mamba-digital-seo-topicos.xlsx"),
    Path(r"C:\Users\Usuario\Desktop\Seo Conteudo\mamba-digital-seo-topicos-com-blog.xlsx"),
    Path(r"C:\Users\Usuario\Desktop\Seo Conteudo\mamba-digital-seo-topicos-com-blog_concorrencia.xlsx"),
    Path(r"C:\Users\Usuario\Desktop\Seo Conteudo\mambadigital.com.br-organic.Positions-br-20260811-2026-08-12T17_22_42Z.xlsx"),
]
OUT = Path(r"C:\Users\Usuario\Desktop\Seo Conteudo\mamba-digital-seo-relatorio.html")

SHEET_ORDER = [
    "Keywords",
    "Topicos",
    "Links",
    "Keywords por Topico",
    "Trafego por Topico",
    "Contagem Unicos",
    "Analise",
    "Analise Concorrencia",
    "Novos termos",
    "Analise Blog",
    "Sitemap URLs",
]
SHEET_LABELS = {
    "Keywords": "Keywords",
    "Topicos": "Tópicos",
    "Links": "Links",
    "Keywords por Topico": "Keywords por Tópico",
    "Trafego por Topico": "Tráfego por Tópico",
    "Contagem Unicos": "Contagem Únicos",
    "Analise": "Análise",
    "Analise Concorrencia": "Análise Concorrência",
    "Novos termos": "Novos termos",
    "Analise Blog": "Análise Blog",
    "Sitemap URLs": "Sitemap URLs",
}


def fmt_cell(val, col: str) -> str:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return "—"
    if isinstance(val, str) and val.strip() == "":
        return "—"
    # URLs
    if col.upper().endswith("URL") or col in {"URL", "URL_SERP"}:
        s = str(val)
        safe = html.escape(s, quote=True)
        short = s if len(s) <= 60 else s[:57] + "…"
        return f'<a href="{safe}" target="_blank" rel="noopener">{html.escape(short)}</a>'
    # Densidade
    if "Densidade" in col or "pct" in col.lower() or "Pct" in col:
        try:
            return f"{float(val):.2f}%"
        except (TypeError, ValueError):
            return html.escape(str(val))
    # Inteiros / trafego / resultados / posicao / keywords counts
    num_cols = {
        "Traffic",
        "Position",
        "Trafego_Total",
        "Trafego_Top1",
        "Trafego_Top5",
        "Trafego_Top10",
        "Trafego_Top50",
        "Keywords",
        "Keywords_Top1",
        "Keywords_Top5",
        "Keywords_Top10",
        "Keywords_Top50",
        "Trafego_keyword",
        "Posicao",
        "Posicao_Google",
        "Resultados_keyword",
        "Resultados_total",
        "Contagem",
        "Subtopicos_unicos",
    }
    if col in num_cols:
        try:
            n = float(val)
            if pd.isna(n):
                return "—"
            if abs(n - int(n)) < 1e-9:
                return f"{int(n):,}".replace(",", ".")
            return f"{n:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        except (TypeError, ValueError):
            pass
    return html.escape(str(val))


def df_to_records(df: pd.DataFrame) -> list[dict]:
    records = []
    for _, row in df.iterrows():
        rec = {}
        for c in df.columns:
            v = row[c]
            if pd.isna(v):
                rec[c] = None
            elif hasattr(v, "item"):
                try:
                    rec[c] = v.item()
                except Exception:
                    rec[c] = str(v)
            else:
                rec[c] = v
        records.append(rec)
    return records


def build_kpis(sheets: dict[str, pd.DataFrame]) -> str:
    kw = sheets["Keywords"]
    n_kw = len(kw)
    n_top = int(kw["Topico"].nunique())
    n_sub = int(kw["Subtopico"].nunique())
    traf = int(kw["Traffic"].sum())
    analise = sheets.get("Analise")
    n_concorrentes = 0
    if analise is not None and not analise.empty:
        n_concorrentes = int(
            analise[analise["Tipo"] == "google_top"]["Empresa"].nunique()
        )
    cards = [
        ("Keywords", f"{n_kw:,}".replace(",", ".")),
        ("Tópicos", str(n_top)),
        ("Subtópicos", str(n_sub)),
        ("Tráfego total", f"{traf:,}".replace(",", ".")),
        ("Concorrentes (Análise)", str(n_concorrentes)),
    ]
    html_cards = "".join(
        f'<div class="kpi"><span class="kpi-val">{v}</span><span class="kpi-label">{l}</span></div>'
        for l, v in cards
    )
    return f'<div class="kpis">{html_cards}</div>'


def main() -> None:
    src = next(p for p in SRC_CANDIDATES if p.exists())
    # Prefer file that already has Analise Blog / Concorrencia
    for p in SRC_CANDIDATES:
        if not p.exists():
            continue
        try:
            names = pd.ExcelFile(p).sheet_names
        except Exception:
            continue
        if "Analise Concorrencia" in names:
            src = p
            break
        if "Analise Blog" in names:
            src = p

    print(f"Lendo: {src}")
    sheets_raw = pd.read_excel(src, sheet_name=None)

    # Gera aba Keywords por Topico + garante Topicos
    if "Keywords" in sheets_raw:
        kw_by_topic = (
            sheets_raw["Keywords"]
            .sort_values(["Topico", "Subtopico", "Traffic"], ascending=[True, True, False])
            .reset_index(drop=True)
        )
        sheets_raw["Keywords por Topico"] = kw_by_topic
        # rebuild Topicos if missing or stale
        try:
            from gerar_aba_topicos import build_topicos

            sheets_raw["Topicos"] = build_topicos(sheets_raw["Keywords"])
        except Exception:
            pass

    sheets = {}
    for k in SHEET_ORDER:
        if k in sheets_raw:
            sheets[k] = sheets_raw[k]

    # Persiste abas derivadas nos XLSX disponiveis
    for p in SRC_CANDIDATES:
        if not p.exists() or "Keywords" not in sheets:
            continue
        try:
            existing = pd.read_excel(p, sheet_name=None)
            ordered = {}
            if "Keywords" in existing:
                ordered["Keywords"] = existing["Keywords"]
                if "Topicos" in sheets:
                    ordered["Topicos"] = sheets["Topicos"]
                elif "Topicos" in existing:
                    ordered["Topicos"] = existing["Topicos"]
                if "Links" in existing:
                    ordered["Links"] = existing["Links"]
                ordered["Keywords por Topico"] = (
                    existing["Keywords"]
                    .sort_values(["Topico", "Subtopico", "Traffic"], ascending=[True, True, False])
                    .reset_index(drop=True)
                )
            for name, df in existing.items():
                if name in {"Keywords", "Topicos", "Links", "Keywords por Topico"}:
                    continue
                ordered[name] = df
            with pd.ExcelWriter(p, engine="openpyxl") as writer:
                for name, df in ordered.items():
                    df.to_excel(writer, sheet_name=name, index=False)
            print(f"XLSX atualizado: {p.name}")
        except PermissionError:
            print(f"[AVISO] XLSX aberto, pulando: {p.name}")
        except Exception as e:
            print(f"[AVISO] Falha ao atualizar {p.name}: {e}")

    # Dados estruturados para Analise Blog (CSV auxiliares ou parse da aba)
    blog_resumo = None
    blog_dist = None
    blog_pages = None
    aux_dir = Path(r"C:\Users\Usuario\Desktop\Seo Conteudo")
    if (aux_dir / "_blog_dist.csv").exists():
        blog_resumo = pd.read_csv(aux_dir / "_blog_resumo.csv")
        blog_dist = pd.read_csv(aux_dir / "_blog_dist.csv")
        blog_pages = pd.read_csv(aux_dir / "_blog_pages.csv")

    # Dados JSON para tabelas interativas
    payload = {}
    for name, df in sheets.items():
        if name == "Contagem Unicos":
            df = df.dropna(how="all").copy()
        if name == "Analise Blog" and blog_dist is not None:
            payload[name] = {
                "label": SHEET_LABELS.get(name, name),
                "columns": list(blog_dist.columns),
                "rows": df_to_records(blog_dist),
                "blog_resumo": df_to_records(blog_resumo) if blog_resumo is not None else [],
                "blog_pages": {
                    "columns": list(blog_pages.columns),
                    "rows": df_to_records(blog_pages),
                }
                if blog_pages is not None
                else None,
            }
        else:
            entry = {
                "label": SHEET_LABELS.get(name, name),
                "columns": list(df.columns),
                "rows": df_to_records(df),
            }
            if name in {"Keywords", "Keywords por Topico"} and "Topico" in df.columns:
                entry["topicos"] = sorted(df["Topico"].dropna().astype(str).unique().tolist())
            payload[name] = entry

    data_json = json.dumps(payload, ensure_ascii=False, default=str)

    tabs_nav = "".join(
        f'<button class="tab-btn{" active" if i == 0 else ""}" data-tab="{html.escape(name)}" type="button">'
        f'{html.escape(SHEET_LABELS.get(name, name))}</button>'
        for i, name in enumerate(sheets)
    )

    def panel_html(i: int, name: str) -> str:
        extra = ""
        if name == "Analise":
            extra = (
                '<label class="opt-toggle" title="Mostrar colunas opcionais">'
                '<input type="checkbox" id="toggle-analise-extra" />'
                " Mostrar detalhes (tráfego, posição, tipo, SERP, queries)"
                "</label>"
            )
        if name == "Analise Concorrencia":
            extra = (
                '<label class="opt-toggle" title="Mostrar colunas opcionais">'
                '<input type="checkbox" id="toggle-concorrencia-extra" />'
                " Mostrar URL e Query"
                "</label>"
            )
        pos_filter = ""
        if name == "Keywords":
            pos_filter = (
                '<div class="pos-filters" data-sheet="Keywords" role="group" aria-label="Filtro de posição">'
                '<button type="button" class="pos-btn active" data-pos="all">Todos</button>'
                '<button type="button" class="pos-btn" data-pos="1">Top 1</button>'
                '<button type="button" class="pos-btn" data-pos="5">Top 5</button>'
                '<button type="button" class="pos-btn" data-pos="10">Top 10</button>'
                "</div>"
            )
        topic_filter = ""
        if name == "Keywords por Topico":
            topic_filter = (
                '<select class="topic-select" data-sheet="Keywords por Topico" aria-label="Filtrar por tópico">'
                '<option value="">Todos os tópicos</option>'
                "</select>"
                '<div class="pos-filters" data-sheet="Keywords por Topico" role="group" aria-label="Filtro de posição">'
                '<button type="button" class="pos-btn active" data-pos="all">Todos</button>'
                '<button type="button" class="pos-btn" data-pos="1">Top 1</button>'
                '<button type="button" class="pos-btn" data-pos="5">Top 5</button>'
                '<button type="button" class="pos-btn" data-pos="10">Top 10</button>'
                "</div>"
            )
        blog_extra = ""
        pages_block = ""
        if name == "Analise Blog":
            blog_extra = '<div id="blog-resumo" class="blog-resumo"></div><div id="blog-bars" class="blog-bars"></div>'
            pages_block = (
                '<div id="blog-pages-wrap" class="blog-pages-wrap" style="margin-top:18px">'
                '<h3 class="section-title">Páginas identificadas</h3>'
                '<div class="table-wrap"><table class="data-table" id="blog-pages-table"><thead></thead><tbody></tbody></table></div>'
                "</div>"
            )

        if name == "Topicos":
            return (
                f'<section class="tab-panel{" active" if i == 0 else ""}" id="panel-{html.escape(name)}" data-sheet="{html.escape(name)}">'
                f'<div class="topics-intro">'
                f'<p class="sub" style="margin:0">Visão por tópico: volume de keywords, rankings Top 1/5/10, tráfego e páginas do blog.</p>'
                f'<div class="topics-sort" role="group" aria-label="Ordenar tópicos">'
                f'<span class="muted-label">Ordenar por</span>'
                f'<button type="button" class="pos-btn active" data-topics-sort="Trafego_total">Tráfego</button>'
                f'<button type="button" class="pos-btn" data-topics-sort="Keywords">Keywords</button>'
                f'<button type="button" class="pos-btn" data-topics-sort="Paginas_blog">Páginas blog</button>'
                f'<button type="button" class="pos-btn" data-topics-sort="Keywords_Top5">Top 5</button>'
                f"</div>"
                f"</div>"
                f'<div id="topics-grid" class="topics-grid"></div>'
                f'<div class="table-wrap" style="margin-top:18px">'
                f'<table class="data-table" data-sheet="{html.escape(name)}"><thead></thead><tbody></tbody></table>'
                f"</div>"
                f'<span class="row-count" data-sheet="{html.escape(name)}" style="display:block;margin-top:8px"></span>'
                f"</section>"
            )

        return (
            f'<section class="tab-panel{" active" if i == 0 else ""}" id="panel-{html.escape(name)}" data-sheet="{html.escape(name)}">'
            f"{blog_extra}"
            f'<div class="toolbar">'
            f'<input type="search" class="search" placeholder="Filtrar nesta aba…" data-sheet="{html.escape(name)}" />'
            f"{pos_filter}{topic_filter}{extra}"
            f'<span class="row-count" data-sheet="{html.escape(name)}"></span>'
            f"</div>"
            f'<div class="table-wrap"><table class="data-table" data-sheet="{html.escape(name)}">'
            f"<thead></thead><tbody></tbody></table></div>"
            f"{pages_block}"
            f"</section>"
        )

    panels = "".join(panel_html(i, name) for i, name in enumerate(sheets))

    kpis = build_kpis(sheets)

    page = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<title>Mamba Digital — Relatório SEO por Tópicos</title>
<link rel="preconnect" href="https://fonts.googleapis.com" />
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin />
<link href="https://fonts.googleapis.com/css2?family=DM+Sans:ital,opsz,wght@0,9..40,400;0,9..40,500;0,9..40,600;0,9..40,700;1,9..40,400&family=Instrument+Serif:ital@0;1&display=swap" rel="stylesheet" />
<style>
:root {{
  --bg: #0f1412;
  --bg-elev: #171e1b;
  --bg-panel: #1c2420;
  --line: #2a3530;
  --text: #e8eee9;
  --muted: #8a9a91;
  --accent: #c8f542;
  --accent-dim: #9bc41f;
  --mamba: #ff4d2e;
  --chip: #243029;
  --danger: #ff6b6b;
  --ok: #6ee7a8;
  --shadow: 0 20px 50px rgba(0,0,0,.35);
}}
* {{ box-sizing: border-box; }}
html, body {{ margin: 0; padding: 0; background: var(--bg); color: var(--text); font-family: "DM Sans", system-ui, sans-serif; }}
body {{
  min-height: 100vh;
  background:
    radial-gradient(1200px 600px at 10% -10%, rgba(200,245,66,.12), transparent 55%),
    radial-gradient(900px 500px at 100% 0%, rgba(255,77,46,.10), transparent 50%),
    linear-gradient(180deg, #121916 0%, #0f1412 40%, #0c100e 100%);
}}
a {{ color: var(--accent); text-decoration: none; }}
a:hover {{ text-decoration: underline; }}
.wrap {{ max-width: 1400px; margin: 0 auto; padding: 28px 20px 60px; }}
.hero {{
  display: grid;
  gap: 10px;
  margin-bottom: 28px;
  padding: 28px 28px 24px;
  border: 1px solid var(--line);
  border-radius: 18px;
  background:
    linear-gradient(135deg, rgba(28,36,32,.95), rgba(15,20,18,.9)),
    repeating-linear-gradient(-45deg, transparent, transparent 8px, rgba(200,245,66,.03) 8px, rgba(200,245,66,.03) 9px);
  box-shadow: var(--shadow);
}}
.brand {{
  font-family: "Instrument Serif", Georgia, serif;
  font-size: clamp(2rem, 4vw, 3rem);
  line-height: 1.05;
  letter-spacing: -0.02em;
}}
.brand em {{
  font-style: italic;
  color: var(--accent);
}}
.sub {{
  color: var(--muted);
  max-width: 62ch;
  font-size: 1rem;
  line-height: 1.5;
}}
.meta {{
  display: flex; flex-wrap: wrap; gap: 8px; margin-top: 8px;
}}
.chip {{
  background: var(--chip);
  border: 1px solid var(--line);
  color: var(--muted);
  font-size: .78rem;
  padding: 5px 10px;
  border-radius: 999px;
}}
.kpis {{
  display: grid;
  grid-template-columns: repeat(5, minmax(0, 1fr));
  gap: 12px;
  margin: 18px 0 22px;
}}
@media (max-width: 900px) {{
  .kpis {{ grid-template-columns: repeat(2, 1fr); }}
}}
.kpi {{
  background: var(--bg-elev);
  border: 1px solid var(--line);
  border-radius: 14px;
  padding: 16px 14px;
  display: flex; flex-direction: column; gap: 4px;
}}
.kpi-val {{
  font-size: 1.55rem;
  font-weight: 700;
  letter-spacing: -0.03em;
  color: var(--accent);
}}
.kpi-label {{ font-size: .8rem; color: var(--muted); }}
.tabs {{
  display: flex; flex-wrap: wrap; gap: 8px;
  margin-bottom: 14px;
  position: sticky; top: 0; z-index: 20;
  padding: 10px 0;
  background: linear-gradient(180deg, rgba(15,20,18,.96), rgba(15,20,18,.88));
  backdrop-filter: blur(8px);
}}
.tab-btn {{
  appearance: none; border: 1px solid var(--line);
  background: var(--bg-elev); color: var(--text);
  padding: 10px 14px; border-radius: 999px;
  font: inherit; font-size: .9rem; font-weight: 600;
  cursor: pointer; transition: .15s ease;
}}
.tab-btn:hover {{ border-color: var(--accent-dim); }}
.tab-btn.active {{
  background: var(--accent);
  color: #111;
  border-color: var(--accent);
}}
.tab-panel {{ display: none; }}
.tab-panel.active {{ display: block; }}
.toolbar {{
  display: flex; align-items: center; justify-content: space-between;
  gap: 12px; margin-bottom: 10px; flex-wrap: wrap;
}}
.search {{
  flex: 1; min-width: 220px;
  background: var(--bg-panel);
  border: 1px solid var(--line);
  color: var(--text);
  border-radius: 10px;
  padding: 10px 12px;
  font: inherit;
}}
.search:focus {{ outline: 2px solid rgba(200,245,66,.35); border-color: var(--accent-dim); }}
.pos-filters {{
  display: inline-flex; gap: 6px; flex-wrap: wrap;
}}
.pos-btn {{
  appearance: none;
  border: 1px solid var(--line);
  background: var(--bg-elev);
  color: var(--muted);
  padding: 8px 12px;
  border-radius: 999px;
  font: inherit;
  font-size: .82rem;
  font-weight: 600;
  cursor: pointer;
}}
.pos-btn:hover {{ border-color: var(--accent-dim); color: var(--text); }}
.pos-btn.active {{
  background: var(--accent);
  color: #111;
  border-color: var(--accent);
}}
.topic-select {{
  background: var(--bg-panel);
  border: 1px solid var(--line);
  color: var(--text);
  border-radius: 10px;
  padding: 9px 12px;
  font: inherit;
  font-size: .86rem;
  min-width: 200px;
  max-width: 280px;
}}
.topic-select:focus {{ outline: 2px solid rgba(200,245,66,.35); border-color: var(--accent-dim); }}
.opt-toggle {{
  display: inline-flex; align-items: center; gap: 8px;
  color: var(--muted); font-size: .85rem;
  background: var(--bg-elev);
  border: 1px solid var(--line);
  border-radius: 999px;
  padding: 8px 12px;
  cursor: pointer; user-select: none; white-space: nowrap;
}}
.opt-toggle input {{ accent-color: var(--accent); cursor: pointer; }}
.row-count {{ color: var(--muted); font-size: .85rem; }}
.blog-resumo {{
  display: grid;
  grid-template-columns: repeat(4, minmax(0, 1fr));
  gap: 12px;
  margin-bottom: 16px;
}}
@media (max-width: 900px) {{
  .blog-resumo {{ grid-template-columns: repeat(2, 1fr); }}
}}
.blog-resumo .kpi-val {{ font-size: 1.4rem; }}
.blog-bars {{
  background: var(--bg-panel);
  border: 1px solid var(--line);
  border-radius: 14px;
  padding: 16px 18px;
  margin-bottom: 16px;
  box-shadow: var(--shadow);
}}
.bar-row {{
  display: grid;
  grid-template-columns: minmax(140px, 220px) 1fr 70px;
  gap: 10px;
  align-items: center;
  margin: 8px 0;
}}
.bar-label {{ font-size: .86rem; color: var(--text); }}
.bar-track {{
  height: 12px;
  background: #243029;
  border-radius: 999px;
  overflow: hidden;
  border: 1px solid var(--line);
}}
.bar-fill {{
  height: 100%;
  background: linear-gradient(90deg, var(--accent-dim), var(--accent));
  border-radius: 999px;
}}
.bar-pct {{
  text-align: right;
  font-variant-numeric: tabular-nums;
  font-weight: 700;
  color: var(--accent);
  font-size: .9rem;
}}
.section-title {{
  font-family: "Instrument Serif", Georgia, serif;
  font-size: 1.35rem;
  font-weight: 400;
  margin: 8px 0 10px;
}}
.blog-pages-wrap {{ display: block; }}
.topics-intro {{
  display: flex; flex-wrap: wrap; justify-content: space-between; gap: 12px;
  align-items: center; margin-bottom: 16px;
}}
.topics-sort {{ display: inline-flex; flex-wrap: wrap; gap: 6px; align-items: center; }}
.muted-label {{ color: var(--muted); font-size: .82rem; margin-right: 4px; }}
.topics-grid {{
  display: grid;
  grid-template-columns: repeat(auto-fill, minmax(300px, 1fr));
  gap: 14px;
  margin-bottom: 8px;
}}
.topic-card {{
  background: var(--bg-panel);
  border: 1px solid var(--line);
  border-radius: 16px;
  padding: 18px 16px 16px;
  box-shadow: var(--shadow);
  display: flex; flex-direction: column; gap: 12px;
  transition: border-color .15s ease, transform .15s ease;
}}
.topic-card:hover {{ border-color: rgba(200,245,66,.35); transform: translateY(-1px); }}
.topic-card h3 {{
  font-family: "Instrument Serif", Georgia, serif;
  font-size: 1.35rem;
  font-weight: 400;
  margin: 0;
  letter-spacing: -0.02em;
}}
.topic-stats {{
  display: grid;
  grid-template-columns: repeat(2, 1fr);
  gap: 8px;
}}
.topic-stat {{
  background: var(--bg-elev);
  border: 1px solid var(--line);
  border-radius: 10px;
  padding: 8px 10px;
}}
.topic-stat .val {{
  display: block;
  font-size: 1.15rem;
  font-weight: 700;
  color: var(--accent);
  font-variant-numeric: tabular-nums;
}}
.topic-stat .lbl {{
  display: block;
  font-size: .72rem;
  color: var(--muted);
  text-transform: uppercase;
  letter-spacing: .04em;
}}
.rank-block {{ display: flex; flex-direction: column; gap: 7px; }}
.rank-row {{
  display: grid;
  grid-template-columns: 52px 1fr 88px;
  gap: 8px;
  align-items: center;
  font-size: .82rem;
}}
.rank-row .rk {{ color: var(--muted); font-weight: 600; }}
.rank-row .nums {{
  text-align: right;
  color: var(--text);
  font-variant-numeric: tabular-nums;
  font-size: .78rem;
}}
.rank-track {{
  height: 8px;
  background: #243029;
  border-radius: 999px;
  overflow: hidden;
  border: 1px solid var(--line);
}}
.rank-fill {{
  height: 100%;
  border-radius: 999px;
  background: linear-gradient(90deg, var(--accent-dim), var(--accent));
}}
.rank-fill.t5 {{ background: linear-gradient(90deg, #6aa84f, #c8f542); }}
.rank-fill.t10 {{ background: linear-gradient(90deg, #3d7a3a, #9bc41f); }}
.topic-foot {{
  display: flex; justify-content: space-between; gap: 8px;
  color: var(--muted); font-size: .78rem; border-top: 1px solid var(--line);
  padding-top: 10px;
}}
.table-wrap {{
  overflow: auto;
  max-height: 70vh;
  border: 1px solid var(--line);
  border-radius: 14px;
  background: var(--bg-panel);
  box-shadow: var(--shadow);
}}
table.data-table {{
  width: 100%;
  border-collapse: collapse;
  font-size: .86rem;
}}
table.data-table th, table.data-table td {{
  padding: 10px 12px;
  border-bottom: 1px solid var(--line);
  text-align: left;
  white-space: nowrap;
}}
table.data-table th {{
  position: sticky; top: 0; z-index: 5;
  background: #222b27;
  color: var(--muted);
  font-weight: 600;
  font-size: .75rem;
  text-transform: uppercase;
  letter-spacing: .04em;
  cursor: pointer;
  user-select: none;
}}
table.data-table th:hover {{ color: var(--accent); }}
table.data-table tbody tr:hover {{ background: rgba(200,245,66,.04); }}
table.data-table td.num {{ font-variant-numeric: tabular-nums; }}
.badge {{
  display: inline-block;
  padding: 2px 8px;
  border-radius: 999px;
  font-size: .75rem;
  font-weight: 600;
  background: var(--chip);
  border: 1px solid var(--line);
}}
.badge.mamba {{ background: rgba(255,77,46,.18); color: #ffb3a6; border-color: rgba(255,77,46,.35); }}
.badge.google {{ background: rgba(200,245,66,.12); color: var(--accent); border-color: rgba(200,245,66,.25); }}
.dens-high {{ color: var(--ok); font-weight: 700; }}
.dens-mid {{ color: #fbbf24; font-weight: 600; }}
.dens-low {{ color: var(--muted); }}
.todos-row {{ background: rgba(200,245,66,.06); font-weight: 600; }}
.footer {{
  margin-top: 28px;
  color: var(--muted);
  font-size: .8rem;
}}
</style>
</head>
<body>
  <div class="wrap">
    <header class="hero">
      <div class="brand">Mamba <em>Digital</em></div>
      <p class="sub">Relatório SEO — keywords agrupadas por tópico e subtópico, tráfego por faixa de posição e análise competitiva (site: + densidade).</p>
      <div class="meta">
        <span class="chip">Cliente: Mamba Digital</span>
        <span class="chip">Fonte: Semrush Organic Positions</span>
        <span class="chip">Domínio: mambadigital.com.br</span>
      </div>
    </header>

    {kpis}

    <nav class="tabs" role="tablist">{tabs_nav}</nav>
    {panels}

    <p class="footer">Gerado a partir de mamba-digital-seo-topicos.xlsx · Densidade = resultados(site:domínio keyword) / resultados(site:domínio) × 100</p>
  </div>

<script>
const DATA = {data_json};

const NUM_COLS = new Set([
  "Traffic","Position","Trafego_Total","Trafego_Top1","Trafego_Top5","Trafego_Top10","Trafego_Top50",
  "Keywords","Keywords_Top1","Keywords_Top5","Keywords_Top10","Keywords_Top50","Trafego_keyword",
  "Posicao","Posicao_Google","Resultados_keyword","Resultados_total","Contagem","Subtopicos_unicos","Densidade_pct","Paginas","Pct_paginas","Valor","Resultados","Search Volume","Traffic","Keywords","Keywords_Top1","Keywords_Top5","Keywords_Top10","Trafego_total","Trafego_Top1","Trafego_Top5","Trafego_Top10","Paginas_blog","Subtopicos","Search_Volume_total","Pct_keywords","Pct_paginas_blog"
]);

/* Colunas opcionais da aba Analise — ocultas por padrao */
const ANALISE_OPTIONAL_COLS = new Set([
  "Trafego_keyword",
  "Posicao",
  "Posicao_Google",
  "Tipo",
  "Titulo_SERP",
  "URL_SERP",
  "Query_keyword",
  "Query_total",
]);

/* Colunas opcionais da aba Analise Concorrencia */
const CONCORRENCIA_OPTIONAL_COLS = new Set([
  "URL",
  "Query",
]);

const state = {{}};

function visibleColumns(name) {{
  const cols = DATA[name].columns;
  if (name === "Analise") {{
    if (state[name] && state[name].showOptional) return cols;
    return cols.filter(c => !ANALISE_OPTIONAL_COLS.has(c));
  }}
  if (name === "Analise Concorrencia") {{
    if (state[name] && state[name].showOptional) return cols;
    return cols.filter(c => !CONCORRENCIA_OPTIONAL_COLS.has(c));
  }}
  return cols;
}}

function fmtNumber(n) {{
  if (n === null || n === undefined || Number.isNaN(n)) return "—";
  const x = Number(n);
  if (!Number.isFinite(x)) return "—";
  if (Math.abs(x - Math.round(x)) < 1e-9) return Math.round(x).toLocaleString("pt-BR");
  return x.toLocaleString("pt-BR", {{ maximumFractionDigits: 2 }});
}}

function fmtCell(col, val, row) {{
  if (val === null || val === undefined || val === "") return "—";
  if (col === "URL" || col === "URL_SERP") {{
    const s = String(val);
    const short = s.length > 60 ? s.slice(0,57) + "…" : s;
    return `<a href="${{s.replace(/"/g,"&quot;")}}" target="_blank" rel="noopener">${{escapeHtml(short)}}</a>`;
  }}
  if (col === "Tipo") {{
    if (val === "mamba") return `<span class="badge mamba">Mamba</span>`;
    if (val === "google_top") return `<span class="badge google">Google Top</span>`;
    return `<span class="badge">${{escapeHtml(String(val))}}</span>`;
  }}
  if (col === "Densidade_pct") {{
    const n = Number(val);
    if (!Number.isFinite(n)) return "—";
    const cls = n >= 20 ? "dens-high" : n >= 5 ? "dens-mid" : "dens-low";
    return `<span class="${{cls}}">${{n.toLocaleString("pt-BR", {{ maximumFractionDigits: 2 }})}}%</span>`;
  }}
  if (col === "Nivel_concorrencia" || col === "Nivel_buscas") {{
    const s = String(val);
    let style = "";
    if (s === "Baixa") style = "color:#6ee7a8;border-color:rgba(110,231,168,.35)";
    else if (s === "Media") style = "color:#fbbf24;border-color:rgba(251,191,36,.35)";
    else if (s === "Alta" || s === "Muito Alta") style = "color:#ffb3a6;border-color:rgba(255,77,46,.35)";
    return `<span class="badge" style="${{style}}">${{escapeHtml(s)}}</span>`;
  }}
  if (col === "Tem_conteudo") {{
    const s = String(val);
    const style = s === "Sim"
      ? "color:#6ee7a8;border-color:rgba(110,231,168,.35)"
      : "color:#ffb3a6;border-color:rgba(255,77,46,.35)";
    return `<span class="badge" style="${{style}}">${{escapeHtml(s)}}</span>`;
  }}
  if (col === "URL" || col === "URL_SERP" || col === "URL_atual") {{
    const s = String(val);
    if (!s || s === "—" || s === "nan") return "—";
    const short = s.length > 60 ? s.slice(0,57) + "…" : s;
    return `<a href="${{s.replace(/"/g,"&quot;")}}" target="_blank" rel="noopener">${{escapeHtml(short)}}</a>`;
  }}
  if (NUM_COLS.has(col)) {{
    return `<span class="num">${{fmtNumber(val)}}</span>`;
  }}
  return escapeHtml(String(val));
}}

function escapeHtml(s) {{
  return s.replace(/[&<>"']/g, c => ({{"&":"&amp;","<":"&lt;",">":"&gt;","\\"":"&quot;","'":"&#39;"}}[c]));
}}

function rowMatches(row, q) {{
  if (!q) return true;
  const hay = Object.values(row).map(v => (v === null || v === undefined) ? "" : String(v).toLowerCase()).join(" ");
  return hay.includes(q);
}}

function sortRows(rows, col, dir) {{
  if (!col) return rows;
  const mul = dir === "desc" ? -1 : 1;
  return [...rows].sort((a, b) => {{
    let va = a[col], vb = b[col];
    const na = Number(va), nb = Number(vb);
    const aNum = va !== null && va !== "" && Number.isFinite(na);
    const bNum = vb !== null && vb !== "" && Number.isFinite(nb);
    if (aNum && bNum) return (na - nb) * mul;
    va = (va === null || va === undefined) ? "" : String(va).toLowerCase();
    vb = (vb === null || vb === undefined) ? "" : String(vb).toLowerCase();
    if (va < vb) return -1 * mul;
    if (va > vb) return 1 * mul;
    return 0;
  }});
}}

function renderBlogExtras(name) {{
  const meta = DATA[name];
  if (name !== "Analise Blog") return;
  const resumoEl = document.getElementById("blog-resumo");
  const barsEl = document.getElementById("blog-bars");
  const pagesWrap = document.getElementById("blog-pages-wrap");
  if (!resumoEl || !barsEl) return;

  if (meta.blog_resumo && meta.blog_resumo.length) {{
    resumoEl.innerHTML = meta.blog_resumo.map(r => {{
      const label = escapeHtml(String(r.Metrica || ""));
      const val = fmtNumber(r.Valor);
      return `<div class="kpi"><span class="kpi-val">${{val}}</span><span class="kpi-label">${{label}}</span></div>`;
    }}).join("");
  }}

  // barras a partir das rows de distribuicao
  const maxPct = Math.max(...meta.rows.map(r => Number(r.Pct_paginas) || 0), 1);
  barsEl.innerHTML = "<div class='section-title' style='margin-top:0'>Distribuição de páginas por tópico</div>" + meta.rows.map(r => {{
    const pct = Number(r.Pct_paginas) || 0;
    const w = Math.max(2, (pct / maxPct) * 100);
    return `<div class="bar-row">
      <div class="bar-label">${{escapeHtml(String(r.Topico))}} <span style="color:var(--muted);font-size:.75rem">(${{fmtNumber(r.Paginas)}})</span></div>
      <div class="bar-track"><div class="bar-fill" style="width:${{w}}%"></div></div>
      <div class="bar-pct">${{pct.toLocaleString("pt-BR", {{ maximumFractionDigits: 2 }})}}%</div>
    </div>`;
  }}).join("");

  if (meta.blog_pages && pagesWrap) {{
    pagesWrap.style.display = "block";
    const table = document.getElementById("blog-pages-table");
    const cols = meta.blog_pages.columns;
    const thead = table.querySelector("thead");
    const tbody = table.querySelector("tbody");
    thead.innerHTML = "<tr>" + cols.map(c => `<th>${{escapeHtml(c)}}</th>`).join("") + "</tr>";
    tbody.innerHTML = meta.blog_pages.rows.map(row =>
      "<tr>" + cols.map(c => `<td>${{fmtCell(c, row[c], row)}}</td>`).join("") + "</tr>"
    ).join("");
  }}
}}

function renderTopicsGrid(rows) {{
  const grid = document.getElementById("topics-grid");
  if (!grid) return;
  const maxKw = Math.max(...rows.map(r => Number(r.Keywords) || 0), 1);
  grid.innerHTML = rows.map(r => {{
    const kw = Number(r.Keywords) || 0;
    const t1 = Number(r.Keywords_Top1) || 0;
    const t5 = Number(r.Keywords_Top5) || 0;
    const t10 = Number(r.Keywords_Top10) || 0;
    const w1 = Math.max(3, (t1 / maxKw) * 100);
    const w5 = Math.max(3, (t5 / maxKw) * 100);
    const w10 = Math.max(3, (t10 / maxKw) * 100);
    return `<article class="topic-card">
      <h3>${{escapeHtml(String(r.Topico))}}</h3>
      <div class="topic-stats">
        <div class="topic-stat"><span class="val">${{fmtNumber(r.Keywords)}}</span><span class="lbl">Keywords</span></div>
        <div class="topic-stat"><span class="val">${{fmtNumber(r.Trafego_total)}}</span><span class="lbl">Tráfego</span></div>
        <div class="topic-stat"><span class="val">${{fmtNumber(r.Paginas_blog)}}</span><span class="lbl">Páginas blog</span></div>
        <div class="topic-stat"><span class="val">${{fmtNumber(r.Subtopicos)}}</span><span class="lbl">Subtópicos</span></div>
      </div>
      <div class="rank-block">
        <div class="rank-row"><span class="rk">Top 1</span><div class="rank-track"><div class="rank-fill" style="width:${{w1}}%"></div></div><span class="nums">${{fmtNumber(t1)}} kw · ${{fmtNumber(r.Trafego_Top1)}} tr</span></div>
        <div class="rank-row"><span class="rk">Top 5</span><div class="rank-track"><div class="rank-fill t5" style="width:${{w5}}%"></div></div><span class="nums">${{fmtNumber(t5)}} kw · ${{fmtNumber(r.Trafego_Top5)}} tr</span></div>
        <div class="rank-row"><span class="rk">Top 10</span><div class="rank-track"><div class="rank-fill t10" style="width:${{w10}}%"></div></div><span class="nums">${{fmtNumber(t10)}} kw · ${{fmtNumber(r.Trafego_Top10)}} tr</span></div>
      </div>
      <div class="topic-foot">
        <span>${{fmtNumber(r.Pct_keywords)}}% das keywords</span>
        <span>${{fmtNumber(r.Pct_paginas_blog)}}% das páginas blog</span>
      </div>
    </article>`;
  }}).join("");
}}

function renderSheet(name) {{
  const meta = DATA[name];
  const st = state[name];
  const q = (st.query || "").trim().toLowerCase();
  let rows = meta.rows.filter(r => rowMatches(r, q));

  // Filtro Top 1 / Top 5 / Top 10 (Position)
  if (st.posLimit && st.posLimit !== "all") {{
    const lim = Number(st.posLimit);
    rows = rows.filter(r => {{
      const p = Number(r.Position);
      return Number.isFinite(p) && p <= lim;
    }});
  }}

  // Filtro por topico (aba Keywords por Topico)
  if (st.topicFilter) {{
    rows = rows.filter(r => String(r.Topico || "") === st.topicFilter);
  }}

  if (name === "Topicos") {{
    const sortCol = st.topicsSort || "Trafego_total";
    rows = sortRows(rows, sortCol, "desc");
    renderTopicsGrid(rows);
  }} else {{
    rows = sortRows(rows, st.sortCol, st.sortDir);
  }}

  const cols = visibleColumns(name);

  const table = document.querySelector(`table.data-table[data-sheet="${{name}}"]`);
  const thead = table.querySelector("thead");
  const tbody = table.querySelector("tbody");
  const countEl = document.querySelector(`.row-count[data-sheet="${{name}}"]`);

  thead.innerHTML = "<tr>" + cols.map(c => {{
    const arrow = st.sortCol === c ? (st.sortDir === "asc" ? " ↑" : " ↓") : "";
    return `<th data-col="${{escapeHtml(c)}}">${{escapeHtml(c)}}${{arrow}}</th>`;
  }}).join("") + "</tr>";

  const MAX = (name === "Keywords" || name === "Keywords por Topico") ? 800 : 5000;
  const slice = rows.slice(0, MAX);
  tbody.innerHTML = slice.map(row => {{
    const isTodos = row.Subtopico === "(Todos)";
    const trClass = isTodos ? ' class="todos-row"' : "";
    return `<tr${{trClass}}>` + cols.map(c => `<td>${{fmtCell(c, row[c], row)}}</td>`).join("") + "</tr>";
  }}).join("");

  let msg = `${{rows.length.toLocaleString("pt-BR")}} linhas`;
  if (rows.length > MAX) msg += ` (mostrando ${{MAX}})`;
  if (name === "Analise Blog") msg = `Distribuição · ${{rows.length}} tópicos`;
  if (name === "Topicos") msg = `${{rows.length}} tópicos`;
  if (st.posLimit && st.posLimit !== "all") msg += ` · Posição ≤ ${{st.posLimit}}`;
  if (st.topicFilter) msg += ` · ${{st.topicFilter}}`;
  if (countEl) countEl.textContent = msg;

  thead.querySelectorAll("th").forEach(th => {{
    th.addEventListener("click", () => {{
      const col = th.getAttribute("data-col");
      if (st.sortCol === col) st.sortDir = st.sortDir === "asc" ? "desc" : "asc";
      else {{ st.sortCol = col; st.sortDir = NUM_COLS.has(col) ? "desc" : "asc"; }}
      renderSheet(name);
    }});
  }});

  if (name === "Analise Blog") renderBlogExtras(name);
}}

function activateTab(name) {{
  document.querySelectorAll(".tab-btn").forEach(b => b.classList.toggle("active", b.dataset.tab === name));
  document.querySelectorAll(".tab-panel").forEach(p => p.classList.toggle("active", p.dataset.sheet === name));
  if (!state[name].rendered) {{
    renderSheet(name);
    state[name].rendered = true;
  }}
}}

Object.keys(DATA).forEach(name => {{
  const defaultSort = {{ sortCol: null, sortDir: "asc" }};
  state[name] = {{
    query: "",
    sortCol: defaultSort.sortCol,
    sortDir: defaultSort.sortDir,
    rendered: false,
    showOptional: false,
    posLimit: "all",
    topicFilter: "",
    topicsSort: "Trafego_total",
  }};
  const input = document.querySelector(`.search[data-sheet="${{name}}"]`);
  let t = null;
  if (input) {{
    input.addEventListener("input", () => {{
      clearTimeout(t);
      t = setTimeout(() => {{
        state[name].query = input.value;
        renderSheet(name);
      }}, 120);
    }});
  }}

  // popula select de topicos
  if (name === "Keywords por Topico" && DATA[name].topicos) {{
    const sel = document.querySelector(`.topic-select[data-sheet="${{name}}"]`);
    if (sel) {{
      DATA[name].topicos.forEach(top => {{
        const opt = document.createElement("option");
        opt.value = top;
        opt.textContent = top;
        sel.appendChild(opt);
      }});
      sel.addEventListener("change", () => {{
        state[name].topicFilter = sel.value;
        renderSheet(name);
      }});
    }}
  }}
}});

document.querySelectorAll("[data-topics-sort]").forEach(btn => {{
  btn.addEventListener("click", () => {{
    document.querySelectorAll("[data-topics-sort]").forEach(b => b.classList.remove("active"));
    btn.classList.add("active");
    if (state["Topicos"]) {{
      state["Topicos"].topicsSort = btn.getAttribute("data-topics-sort");
      renderSheet("Topicos");
    }}
  }});
}});

// Filtros Top 1 / 5 / 10
document.querySelectorAll(".pos-filters").forEach(group => {{
  const sheet = group.getAttribute("data-sheet");
  group.querySelectorAll(".pos-btn").forEach(btn => {{
    btn.addEventListener("click", () => {{
      group.querySelectorAll(".pos-btn").forEach(b => b.classList.remove("active"));
      btn.classList.add("active");
      state[sheet].posLimit = btn.getAttribute("data-pos");
      renderSheet(sheet);
    }});
  }});
}});

const analiseToggle = document.getElementById("toggle-analise-extra");
if (analiseToggle) {{
  analiseToggle.addEventListener("change", () => {{
    state["Analise"].showOptional = analiseToggle.checked;
    renderSheet("Analise");
  }});
}}

const concorrenciaToggle = document.getElementById("toggle-concorrencia-extra");
if (concorrenciaToggle) {{
  concorrenciaToggle.addEventListener("change", () => {{
    state["Analise Concorrencia"].showOptional = concorrenciaToggle.checked;
    renderSheet("Analise Concorrencia");
  }});
}}

document.querySelectorAll(".tab-btn").forEach(btn => {{
  btn.addEventListener("click", () => activateTab(btn.dataset.tab));
}});

activateTab(Object.keys(DATA)[0]);
</script>
</body>
</html>
"""

    OUT.write_text(page, encoding="utf-8")
    print(f"HTML salvo: {OUT}")
    print(f"Tamanho: {OUT.stat().st_size / 1024:.1f} KB")


if __name__ == "__main__":
    main()
