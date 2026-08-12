# -*- coding: utf-8 -*-
"""
Aba Novos termos:
1) Match nas Keywords atuais (Semrush)
2) Fallback: URLs do sitemap https://mambadigital.com.br/sitemap.xml
Sem site: Google.
"""
from __future__ import annotations

import re
import unicodedata
from pathlib import Path
from urllib.parse import unquote, urlparse

import pandas as pd
import requests

SCRIPT_DIR = Path(__file__).resolve().parent
SITEMAP_URL = "https://mambadigital.com.br/sitemap.xml"
SITEMAP_FILE = SCRIPT_DIR / "mamba_sitemap_urls.txt"
SHEET = "Novos termos"

XLSX_CANDIDATES = [
    SCRIPT_DIR / "mamba-digital-seo-topicos.xlsx",
    SCRIPT_DIR / "mamba-digital-seo-topicos-com-blog.xlsx",
    SCRIPT_DIR / "mamba-digital-seo-topicos-com-blog_concorrencia.xlsx",
    SCRIPT_DIR
    / "mambadigital.com.br-organic.Positions-br-20260811-2026-08-12T17_22_42Z.xlsx",
]

RAW_TERMS = r"""
como vender na shopee	18100
como vender no mercado livre	14800
como vender na shopee sem estoque	3600
como vender na amazon	3600
shopee ads	1900
como vender no mercado livre sem estoque	1900
entrega full mercado livre	1600
como ver o estoque do full mercado livre	1300
mercado ads	1000
como vender na shopee do zero	880
curso mercado livre	720
como começar a vender no mercado livre	720
como começar a vender na shopee	720
o que significa anúncio pausado no mercado livre	720
problema anúncio grátis mercado livre	720
como vender na shopee como afiliado	590
o que significa full no mercado livre	480
como vender produtos no mercado livre	480
formulário full mercado livre	390
como vender no mercado livre passo a passo	390
tarifas mercado livre	320
envio full mercado livre	320
como vender no mercado livre como afiliado	320
como funciona o full do mercado livre	260
full super mercado livre	260
o que vender na shopee com baixo investimento	260
tamanho da imagem para anúncio no mercado livre	260
shopee ads como funciona	210
produtos full mercado livre	210
curso para vender no mercado livre	210
vender no magalu	210
como vender mais na shopee	210
vender na shopee com cpf	210
mamba assessoria de marketplace	210
como ser full no mercado livre	170
full mercado livre como funciona	170
como se tornar uma loja oficial no mercado livre	170
consultoria mercado livre	170
como vender mais no mercado livre	170
como vender no mercado livre pessoa física	170
como funciona venda na shopee	170
o que é shopee ads	140
shopee ads vale a pena	140
solicitar full mercado livre	140
como vender no mercado livre com cnpj	140
reputação mercado livre	110
vender em marketplace	110
como vender no magalu	110
como colocar vídeo no anúncio do mercado livre	110
product ads mercado livre	110
como vender em marketplace	90
como ver meus anúncios no mercado livre	90
display ads mercado livre	90
consultoria amazon	70
como vender na amazon brasil	70
como aumentar as vendas no mercado livre	70
anúncio pausado mercado livre significado	70
brand ads mercado livre	70
o que é ads no mercado livre	70
novas tarifas mercado livre 2026	50
tarifa de venda por categoria mercado livre	50
consultoria marketplace	50
assessoria mercado livre	50
como fazer ads no mercado livre	50
como fazer ads na shopee 2025	40
como vender na shopee sem ads	40
cnae para vender em marketplace	40
mercado livre ads preço	40
como criar anúncio na shopee ads	30
shopee video ads	30
produtos para vender em marketplace	30
consultoria shopee	30
mamba - assessoria de marketplace comentários	30
como colocar ads no mercado livre	30
mercado livre ads vale a pena	30
o que é mercado ads	30
calculadora tarifas mercado livre	20
como funciona a tarifa de venda do mercado livre	20
tabela de tarifas mercado livre	20
o que é tarifa de venda mercado livre	20
como melhorar as vendas em marketplace	20
curso para vender em marketplace	20
o que vender em marketplace	20
produtos mais vendidos em marketplace	20
vale a pena vender em marketplace	20
vender em vários marketplaces	20
consultor certificado mercado livre	20
consultor magalu	20
como aumentar as vendas no magalu	20
como cadastrar para vender no magalu	20
como vender no magalu com cpf	20
como vender no magalu como afiliado	20
como vender no magalu passo a passo	20
como vender produtos no magalu	20
produtos mais vendidos no magalu	20
avaliações sobre mamba assessoria de marketplace	20
gestão de marketplaces	20
gestão de ecommerce e marketplace	20
integração de vendas em marketplaces	10
consultor americanas marketplace	10
assessoria de marketplace	10
aumentar reputação mercado livre	5
como aumentar a reputação no mercado livre	5
barra de reputação mercado livre	5
categorias de reputação mercado livre	5
plataforma de gestão de marketplace	5
assessoria especializada mercado livre	5
"""

STOP = {
    "de", "da", "do", "das", "dos", "a", "o", "e", "em", "no", "na", "nos", "nas",
    "para", "com", "como", "que", "um", "uma", "os", "as", "ao", "se", "ou", "por",
    "mais", "meu", "meus", "sua", "seu", "the", "and", "blog", "page", "https", "www",
    "mambadigital", "com", "br",
}


def norm(s: str) -> str:
    s = str(s).lower().strip()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(c for c in s if not unicodedata.combining(c))
    s = s.replace("'", "")
    s = re.sub(r"[^a-z0-9\s]", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def tokens(s: str) -> set[str]:
    return {t for t in norm(s).split() if t not in STOP and len(t) > 1}


def jaccard(a: set[str], b: set[str]) -> float:
    if not a or not b:
        return 0.0
    return len(a & b) / len(a | b)


def platform(s: str) -> set[str]:
    n = norm(s)
    out = set()
    if "shopee" in n or "shoppe" in n:
        out.add("shopee")
    if "mercado livre" in n or "mercadolivre" in n:
        out.add("mercadolivre")
    if "amazon" in n:
        out.add("amazon")
    if "magalu" in n or "magazine luiza" in n:
        out.add("magalu")
    if "americanas" in n:
        out.add("americanas")
    return out


def platforms_ok(term: str, target: str) -> bool:
    specific = {"shopee", "mercadolivre", "amazon", "magalu", "americanas"}
    t = platform(term) & specific
    u = platform(target) & specific
    if t and u and t.isdisjoint(u):
        return False
    if t and not u:
        return False
    return True


def parse_terms() -> pd.DataFrame:
    rows = []
    for line in RAW_TERMS.strip().splitlines():
        line = line.strip()
        if not line:
            continue
        if "\t" in line:
            kw, vol = line.rsplit("\t", 1)
        else:
            kw, vol = line.rsplit(None, 1)
        rows.append({"Keyword": kw.strip(), "Search Volume": int(vol)})
    return pd.DataFrame(rows)


def download_sitemap_urls() -> list[str]:
    print(f"Baixando {SITEMAP_URL} ...")
    r = requests.get(
        SITEMAP_URL,
        timeout=90,
        headers={"User-Agent": "Mozilla/5.0 (compatible; MambaSeoBot/1.0)"},
    )
    r.raise_for_status()
    (SCRIPT_DIR / "sitemap.xml").write_bytes(r.content)
    text = r.text
    locs = re.findall(r"<loc>\s*([^<]+?)\s*</loc>", text, flags=re.I)
    urls = []
    seen = set()
    for loc in locs:
        u = unquote(loc.strip().replace("&amp;", "&"))
        # limpa UTM/query para canonical
        parsed = urlparse(u)
        clean = f"{parsed.scheme}://{parsed.netloc}{parsed.path}"
        if not clean.endswith("/") and "/blog/" in clean:
            clean += "/"
        if clean not in seen:
            seen.add(clean)
            urls.append(clean)
    SITEMAP_FILE.write_text("\n".join(urls), encoding="utf-8")
    print(f"Sitemap: {len(urls)} URLs unicas salvas em {SITEMAP_FILE.name}")
    return urls


def slug_text(url: str) -> str:
    path = urlparse(url).path.strip("/")
    # remove prefixos irrelevantes
    path = re.sub(r"^blog/", "", path)
    path = path.replace("-", " ").replace("/", " ")
    return path


def build_kw_inventory(kw: pd.DataFrame) -> pd.DataFrame:
    inv = kw.copy()
    inv["_n"] = inv["Keyword"].map(norm)
    inv["_tok"] = inv["Keyword"].map(tokens)
    inv["_blog"] = inv["URL"].astype(str).str.contains("/blog/", case=False, na=False)
    return inv.sort_values(["_blog", "Traffic", "Position"], ascending=[False, False, True])


def build_sitemap_inventory(urls: list[str]) -> pd.DataFrame:
    rows = []
    for u in urls:
        # prioriza blog; ignora home pura e query lixo ja limpo
        if urlparse(u).path in {"", "/"}:
            continue
        st = slug_text(u)
        rows.append(
            {
                "URL": u,
                "_slug": st,
                "_n": norm(st),
                "_tok": tokens(st),
                "_is_blog": "/blog/" in u.lower(),
            }
        )
    return pd.DataFrame(rows)


def match_keywords(term: str, inv: pd.DataFrame) -> dict | None:
    nt = norm(term)
    tt = tokens(term)

    exact = inv[inv["_n"] == nt]
    if len(exact):
        r = exact.iloc[0]
        return {
            "Tem_conteudo": "Sim",
            "Fonte_match": "Keywords",
            "Match_tipo": "Exato",
            "Keyword_existente": r["Keyword"],
            "URL_atual": r["URL"],
            "Position_atual": int(r["Position"]),
            "Score": 1.0,
        }

    best = None
    best_score = 0.0
    best_tipo = ""
    for _, r in inv.iterrows():
        if not platforms_ok(term, f"{r['Keyword']} {r['URL']}"):
            continue
        ne = r["_n"]
        score = 0.0
        tipo = ""
        if nt in ne or ne in nt:
            ratio = min(len(nt), len(ne)) / max(len(nt), len(ne))
            if ratio >= 0.55 or (nt in ne and len(tt) >= 3 and tt.issubset(tokens(ne))):
                score = 0.82 + 0.15 * ratio
                tipo = "Parcial"
        else:
            jac = jaccard(tt, r["_tok"])
            if jac >= 0.78 and len(tt & r["_tok"]) >= 3:
                score = 0.72 + 0.25 * jac
                tipo = "Similar"
        if score > best_score:
            best_score = score
            best = r
            best_tipo = tipo

    if best is not None and best_score >= 0.78:
        return {
            "Tem_conteudo": "Sim",
            "Fonte_match": "Keywords",
            "Match_tipo": best_tipo,
            "Keyword_existente": best["Keyword"],
            "URL_atual": best["URL"],
            "Position_atual": int(best["Position"]),
            "Score": round(best_score, 3),
        }
    return None


def match_sitemap(term: str, sm: pd.DataFrame) -> dict | None:
    """Fallback: compara o termo com o slug da URL do sitemap."""
    nt = norm(term)
    tt = tokens(term)
    if not tt:
        return None

    best = None
    best_score = 0.0
    best_tipo = ""

    for _, r in sm.iterrows():
        target = f"{r['_slug']} {r['URL']}"
        if not platforms_ok(term, target):
            continue
        ne = r["_n"]
        score = 0.0
        tipo = ""
        if nt == ne or nt.replace(" ", "") == ne.replace(" ", ""):
            score, tipo = 0.95, "Slug exato"
        elif nt in ne or ne in nt:
            ratio = min(len(nt), len(ne)) / max(len(nt), len(ne))
            if ratio >= 0.5:
                score = 0.8 + 0.12 * ratio
                tipo = "Slug parcial"
        else:
            jac = jaccard(tt, r["_tok"])
            inter = len(tt & r["_tok"])
            # exige boa cobertura dos tokens do termo no slug
            coverage = inter / max(len(tt), 1)
            if jac >= 0.55 and coverage >= 0.6 and inter >= 2:
                score = 0.55 + 0.35 * coverage + 0.1 * jac
                tipo = "Slug similar"
            elif coverage >= 0.75 and inter >= 3:
                score = 0.6 + 0.3 * coverage
                tipo = "Slug similar"

        # leve bonus para posts de blog
        if score > 0 and r["_is_blog"]:
            score += 0.03

        if score > best_score:
            best_score = score
            best = r
            best_tipo = tipo

    if best is not None and best_score >= 0.72:
        return {
            "Tem_conteudo": "Sim",
            "Fonte_match": "Sitemap",
            "Match_tipo": best_tipo,
            "Keyword_existente": best["_slug"],
            "URL_atual": best["URL"],
            "Position_atual": None,
            "Score": round(best_score, 3),
        }
    return None


def main() -> None:
    urls = download_sitemap_urls()
    sm = build_sitemap_inventory(urls)
    print(f"Inventario sitemap util: {len(sm)} (blog={int(sm['_is_blog'].sum())})")

    src = next(p for p in XLSX_CANDIDATES if p.exists())
    kw = pd.read_excel(src, sheet_name="Keywords")
    inv = build_kw_inventory(kw)
    terms = parse_terms()
    print(f"Novos termos: {len(terms)}")

    rows = []
    for i, r in terms.iterrows():
        term = r["Keyword"]
        vol = int(r["Search Volume"])
        hit = match_keywords(term, inv)
        if hit:
            print(f"[{i+1}/{len(terms)}] KW {hit['Match_tipo']}: {term}")
        else:
            hit = match_sitemap(term, sm)
            if hit:
                print(f"[{i+1}/{len(terms)}] SM {hit['Match_tipo']}: {term} -> {hit['URL_atual']}")
            else:
                hit = {
                    "Tem_conteudo": "Nao",
                    "Fonte_match": "",
                    "Match_tipo": "Sem match",
                    "Keyword_existente": "",
                    "URL_atual": "",
                    "Position_atual": None,
                    "Score": 0.0,
                }
                print(f"[{i+1}/{len(terms)}] NAO: {term}")
        rows.append({"Keyword": term, "Search Volume": vol, **hit})

    out = pd.DataFrame(rows).sort_values("Search Volume", ascending=False).reset_index(drop=True)
    cols = [
        "Keyword",
        "Search Volume",
        "Tem_conteudo",
        "Fonte_match",
        "Match_tipo",
        "Keyword_existente",
        "URL_atual",
        "Position_atual",
        "Score",
    ]
    out = out[cols]

    print("\n=== Resumo ===")
    print(out["Tem_conteudo"].value_counts().to_string())
    print(out.groupby(["Tem_conteudo", "Fonte_match"]).size().to_string())
    print("Oportunidades (Nao):", int((out["Tem_conteudo"] == "Nao").sum()))

    out.to_csv(SCRIPT_DIR / "_novos_termos.csv", index=False)

    # aba auxiliar com urls do sitemap
    sm_out = pd.DataFrame({"URL": urls})

    for p in XLSX_CANDIDATES:
        if not p.exists():
            continue
        try:
            existing = pd.read_excel(p, sheet_name=None)
            existing[SHEET] = out
            existing["Sitemap URLs"] = sm_out
            with pd.ExcelWriter(p, engine="openpyxl") as writer:
                for name, sheet in existing.items():
                    sheet.to_excel(writer, sheet_name=name, index=False)
            print("OK", p.name)
        except PermissionError:
            print("LOCKED", p.name)


if __name__ == "__main__":
    main()
