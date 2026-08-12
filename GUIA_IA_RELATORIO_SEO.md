# Guia para outra IA recriar o Relatório SEO Mamba Digital

Este documento descreve **como cada aba do relatório foi construída**, na ordem prática de execução, para que outra IA (ou analista) reproduza o mesmo pipeline.

Cliente: **Mamba Digital** (`mambadigital.com.br`)  
Entregáveis finais: XLSX multiabas + HTML com abas (`mamba-digital-seo-relatorio.html` / `index.html`).

---

## Visão geral do pipeline

```
1. Semrush Organic Positions (XLSX bruto)
2. Classificar Keyword → Topico + Subtopico
3. Agregar tráfego Top1/5/10/50
4. Contagem de únicos
5. Análise SERP (top 5 + Mamba) com site: e densidade
6. Análise Blog (páginas únicas e % por tópico)
7. Análise Concorrência (Top5 sem Marca) com "kw" -instagram -youtube -tiktok
8. Cruzar Search Volume (Semrush)
9. Novos termos × inventário Keywords + fallback Sitemap
10. Aba visual Tópicos
11. Aba Links (curadoria manual)
12. Exportar HTML com abas
```

**Ferramentas usadas nesta execução:** Python (`pandas`, `openpyxl`, `requests`), ScrapingBee Google Search API (só onde indicado), sitemap público da Mamba.

**Não usar `site:` do Google** na etapa de Novos termos — o fallback correto é o **sitemap**.

---

## Dados de entrada

| Arquivo | Uso |
|---------|-----|
| `mambadigital.com.br-organic.Positions-*.xlsx` | Keywords Semrush: `Keyword`, `Position`, `URL`, `Traffic` |
| `mamba nivel de pesquisa.xlsx` | `Keyword`, `Search Volume` (mesmo número de linhas / mesma ordem do Positions) |
| `https://mambadigital.com.br/sitemap.xml` | Lista de URLs publicadas no site |

---

## Aba: Keywords

**Objetivo:** base canônica de palavras-chave classificadas.

### Passos
1. Ler o XLSX de Organic Positions.
2. Criar colunas `Topico` e `Subtopico`.
3. Classificar **todas** as linhas (não só amostra):
   - Preferir mapeamento por **path da URL** do blog (ex.: `/blog/quais-sao-as-taxas-da-shopee` → `Shopee` / `Taxas e Calculadora`).
   - Fallback por regras na keyword (shopee, mercado livre, marca mamba, etc.).
   - Home / typos / CNPJ / concorrentes na home → `Ruido SERP` (não misturar com conteúdo).
4. Cruzar `Search Volume` do arquivo `mamba nivel de pesquisa.xlsx` (alinhamento por índice verificado; se a ordem for idêntica, `insert` por linha).
5. Colunas finais sugeridas:  
   `Keyword | Search Volume | Topico | Subtopico | Position | URL | Traffic`

### Taxonomia (exemplo usado)
- **Shopee**, **Mercado Livre**, **Marca Mamba**, **Comparativo de Marketplaces**, **Operacao Marketplace**, **Marketing Digital**, **Gestao e Financas**, **Promocoes e Datas Comerciais**, **Conteudo Geral**, **Ruido SERP**
- Subtópicos por intenção (Como Vender, Taxas, Full, Ads, etc.)

### Script de referência
- Classificação original foi feita em script ad hoc; manter regras em um `classificar_topicos.py` se for reexecutar.

---

## Aba: Tópicos

**Objetivo:** visão agregada **visual** (cards no HTML) por tópico único.

### Passos
1. A partir de `Keywords`, agrupar por `Topico` (excluir `Ruido SERP` da visão principal).
2. Calcular por tópico:
   - `Keywords` (contagem)
   - `Subtopicos` (nunique)
   - `Paginas_blog` = URLs únicas contendo `/blog/`
   - `Trafego_total`
   - `Keywords_Top1` / `Top5` / `Top10` onde `Position <= 1/5/10`
   - `Trafego_Top1` / `Top5` / `Top10`
   - `Search_Volume_total` (soma)
   - `%` de keywords e `%` de páginas blog
3. Ordenar por tráfego (ou permitir sort no HTML).

### Script
- `gerar_aba_topicos.py` → função `build_topicos()`

### HTML
- Cards com barras Top1/5/10 + tabela compacta abaixo.
- Aba fica **logo após Keywords**.

---

## Aba: Links

**Objetivo:** curadoria manual de keywords específicas (não é dump automático).

### Passos
1. Receber lista de termos do usuário.
2. Buscar em `Keywords` (match normalizado, sem acento).
3. Se houver duplicatas Semrush, **manter 1 linha**: melhor `Position` (menor número), desempate por `Traffic`.
4. Aplicar correções manuais pedidas (ex.: forçar Position=5 em termos específicos).
5. Não remover essas linhas de `Keywords` (senão quebra agregados), a menos que o usuário peça explicitamente.

### Lista atual neste projeto
- simulador de custos mercado livre  
- vender na shopee  
- quantos a shopee cobra por venda  
- como fazer dropshipping no mercado livre  
- maiores marketplaces do brasil  

---

## Aba: Keywords por Tópico

**Objetivo:** mesma base de Keywords, ordenada para navegação por tema.

### Passos
1. Copiar `Keywords`.
2. Ordenar por `Topico`, `Subtopico`, `Traffic` (desc).
3. No HTML: filtro dropdown de tópico + filtros Top1/Top5/Top10 (`Position <= N`).

**Atenção:** keywords quase iguais podem ter posições diferentes  
(ex.: `maior ecommerce do brasil` pos 6 vs `maior e commerce do brasil` pos 1).

---

## Aba: Tráfego por Tópico

**Objetivo:** tráfego agregado por tópico e por subtópico, em faixas de posição.

### Passos
1. Para cada grupo (`Topico`) e (`Topico`+`Subtopico`):
   - `Keywords`, `Trafego_Total`
   - Para limites 1, 5, 10, 50:  
     `Keywords_TopN` e `Trafego_TopN` com `Position <= N`
2. Incluir linhas `Subtopico = (Todos)` para o total do tópico.
3. Coluna `Subtopicos_unicos` só nas linhas `(Todos)`.

---

## Aba: Contagem Únicos

**Objetivo:** KPIs de cardinalidade da taxonomia.

### Passos
1. Calcular:
   - Tópicos únicos
   - Subtópicos únicos
   - Pares Topico+Subtopico únicos
   - Total de keywords
2. Listar tópicos com contagem de keywords e n° de subtópicos.
3. Listar detalhe Topico → Subtopico → Keywords.

---

## Aba: Análise

**Objetivo:** competidores SERP + densidade de conteúdo (`site:`).

### Passos (por tópico único, 1 keyword representativa — maior Traffic; excluir Ruido/Conteudo Geral se ruído)
1. Google BR: keyword → **top 5 orgânicos** (domínios únicos), **excluindo** `mambadigital.com.br`.
2. Sempre incluir **posição 6 = mambadigital.com.br**.
3. Para cada domínio:
   - `site:dominio keyword` → `Resultados_keyword`
   - `site:dominio` → `Resultados_total` (cachear por domínio)
   - `Densidade_pct = Resultados_keyword / Resultados_total * 100`
4. API: ScrapingBee Google Search (`seo_site_results.py`).

### HTML
- Colunas opcionais ocultas por padrão: tráfego keyword, posição, tipo, título/URL SERP, queries.  
  Toggle: “Mostrar detalhes…”.

### Script
- `gerar_aba_analise.py` + helpers em `seo_site_results.py`

---

## Aba: Análise Concorrência

**Objetivo:** nível de competição SERP para keywords em que a Mamba já está no Top 5 (exceto Marca).

### Passos
1. Filtrar `Keywords` com `Position <= 5` e `Topico` ∉ {`Marca Mamba`, `Ruido SERP`}.
2. Deduplicar keyword normalizada (maior Traffic).
3. **Amostra rápida:** top N por tráfego (neste projeto N=20; pode subir para ~225).
4. Query Google:  
   `"palavra chave" -instagram -youtube -tiktok`
5. Guardar `Resultados` (number_of_results).
6. `Nivel_concorrencia` por faixas de resultados (ex.: Baixa ≤1k, Média ≤50k, Alta ≤500k, Muito Alta >500k).
7. Cruzar `Search Volume` e recalcular `Nivel_buscas` por volume Semrush  
   (ex.: Baixa ≤500, Média ≤2k, Alta ≤10k, Muito Alta >10k).

### HTML
- Ocultar `URL` e `Query` por padrão (toggle).

### Script
- `gerar_analise_concorrencia.py`

---

## Aba: Novos termos

**Objetivo:** lista de termos novos (com volume) × cobertura do site.

### Passos
1. Receber lista `termo \t volume`.
2. **Match 1 — Keywords atuais** (exato → parcial → similar), respeitando plataforma (Shopee ≠ Magalu ≠ Amazon).
3. **Match 2 — fallback Sitemap** (não usar `site:` Google):
   - Baixar `https://mambadigital.com.br/sitemap.xml`
   - Extrair todas as `<loc>`
   - Comparar tokens do termo com o **slug** da URL
4. Colunas:  
   `Keyword | Search Volume | Tem_conteudo (Sim/Nao) | Fonte_match (Keywords/Sitemap) | Match_tipo | Keyword_existente | URL_atual | Position_atual | Score`

### Scripts
- `gerar_novos_termos_sitemap.py` (versão correta)
- Evitar `gerar_novos_termos.py` / `v2` se usarem `site:` como fallback principal

### Aba auxiliar
- `Sitemap URLs` = lista limpa das URLs do sitemap.

---

## Aba: Análise Blog

**Objetivo:** quantas páginas de blog existem no inventário Semrush e distribuição % por tópico.

### Passos
1. Filtrar URLs `mambadigital.com.br/blog`.
2. Deduplicar por path.
3. Atribuir tópico (sinais no path + majority vote das keywords).
4. KPIs: total páginas, artigos vs categorias, tópicos cobertos.
5. Tabela de distribuição: `Topico | Paginas | Pct_paginas | ...`

### Script
- `gerar_analise_blog.py`

---

## HTML final (abas)

### Passos
1. Rodar `gerar_html_relatorio.py`.
2. Embutir dados das abas em JSON no HTML.
3. UI: tabs, busca, sort por coluna, filtros Top1/5/10, cards de Tópicos, toggles de colunas opcionais.
4. Para GitHub Pages: publicar como `index.html` na raiz do repo.

### Script
- `gerar_html_relatorio.py` → `mamba-digital-seo-relatorio.html`

---

## Ordem recomendada das abas no entregável

1. Keywords  
2. Tópicos  
3. Links  
4. Keywords por Tópico  
5. Tráfego por Tópico  
6. Contagem Únicos  
7. Análise  
8. Análise Concorrência  
9. Novos termos  
10. Análise Blog  
11. Sitemap URLs  

---

## Checklist anti-erro (para a IA seguinte)

- [ ] Não classificar ruído de marca/home como conteúdo editorial.  
- [ ] Não usar `site:` Google para decidir se “já tem conteúdo” em Novos termos — usar sitemap.  
- [ ] Em Links: 1 linha por keyword; melhor posição; aplicar overrides manuais do usuário.  
- [ ] Search Volume: validar alinhamento de linhas antes do merge.  
- [ ] Densidade >100% pode ocorrer (estimativa do Google) — documentar, não “corrigir” artificialmente sem critério.  
- [ ] Não commitar API keys ScrapingBee / configs Semrush.  
- [ ] Ao regenerar HTML, preservar aba `Links` e ordem das sheets.

---

## Comandos úteis (PowerShell)

```powershell
cd "C:\Users\Usuario\Desktop\Seo Conteudo"
python gerar_aba_topicos.py
python gerar_novos_termos_sitemap.py
python gerar_analise_concorrencia.py   # ajustar limit=20 ou completo
python gerar_html_relatorio.py
```

Variável necessária só para etapas ScrapingBee:

```powershell
$env:SCRAPINGBEE_API_KEY = "sua_chave"
```

---

## Critério de “igual ao relatório atual”

Uma recriação é considerada equivalente se:
1. Taxonomia Topico/Subtopico cobre ≥95% das keywords sem `Outros`.  
2. Abas de agregação batem com as mesmas definições de Top1/5/10/50.  
3. Análise / Concorrência usam a mesma lógica de queries.  
4. Novos termos usam Keywords + Sitemap (sem `site:`).  
5. HTML tem as mesmas abas, filtros e colunas opcionais.
