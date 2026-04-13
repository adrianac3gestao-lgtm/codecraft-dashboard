#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
gerar_dados_dashboard.py
CodeCraft - Atualizador de Dashboard
--------------------------------------
COMO USAR:
  1. Verifique os caminhos EXCEL_PATH e HTML_PATH abaixo
  2. Execute: python gerar_dados_dashboard.py
  3. O script le o Excel, calcula tudo e atualiza o index.html

INSTALAR DEPENDENCIAS (apenas uma vez):
  pip install pandas openpyxl
"""

import json, sys, shutil, datetime, unicodedata
from pathlib import Path

# ============================================================
#  CONFIGURACAO - ajuste os caminhos se necessario
# ============================================================
EXCEL_PATH = r"C:\Users\adria\OneDrive\C3 Gestao\EquipeC3\CODECRAFT\Relatorio Gerencial_Codecraft2026.xlsx"
HTML_PATH  = r"C:\Users\adria\OneDrive\C3 Gestao\EquipeC3\CODECRAFT\7. Dashboard\index.html"
SHEET_NAME = "4.Consolidado"

# Marcadores no HTML que delimitam o bloco de dados
MARKER_START = "// DATA BLOCK - update with gerar_dados_dashboard.py"
MARKER_END   = "const LABEL_MAP"

# ============================================================
#  DEPENDENCIAS
# ============================================================
try:
    import pandas as pd
except ImportError:
    print("ERRO: pandas nao instalado.")
    print("Execute: pip install pandas openpyxl")
    sys.exit(1)

# ============================================================
#  HELPERS
# ============================================================
def pausar():
    try:
        if sys.stdin.isatty():
            input("\nPressione Enter para fechar...")
    except Exception:
        pass

def rnd(v):
    return round(float(v or 0), 2)

def fmt_brl(v):
    s = "{:,.2f}".format(abs(v)).replace(",","X").replace(".",",").replace("X",".")
    return "R$ " + s

def norm_cat(s):
    """Remove acentos e normaliza nome de categoria para comparacao."""
    s2 = unicodedata.normalize("NFD", str(s))
    return "".join(c for c in s2 if unicodedata.category(c) != "Mn").upper().strip()

# ============================================================
#  CATEGORIAS
# ============================================================
CAT_PESSOAL  = ["SALARIOS E ORDENADOS ADM", "PREMIOS E GRATIFICACOES ADM"]
CAT_MKT      = ["SERVS. DE PUBLICIDADE E PROPAGANDA"]
CAT_JURIDICO = ["SERVS. ADVOCATICIOS"]
CAT_OUTROS   = [
    "SERVS. DE ASSESSORIA E CONSULTORIA", "SERVICO PRESTADOS POR TERCEIROS",
    "SERVS. DE CONTABILIDADE", "SISTEMAS OPERACIONAIS", "SERVS. CERTIFICADO DIGITAL",
    "COMPRA DE ATIVO IMOBILIZADO", "ISS", "CSRF", "REEMB DESPESAS GERAIS",
    "ESTORNO_DEVOLUCAO (-)", "SERVS. ADMINISTRATIVOS",
    # Categorias Bruno Niro (periodo inicial abr-out/25)
    "SERVICOS PRESTADOS PJ - INICIAL", "IMPOSTOS - INICIAL",
    "SERVS. DE RECRUTAMENTO E SELECAO - PJ"
]

CAT_COLORS = {
    "SALARIOS E ORDENADOS ADM":           "#f59e0b",
    "SERVS. DE PUBLICIDADE E PROPAGANDA": "#8b5cf6",
    "SERVS. ADVOCATICIOS":                "#ef4444",
    "SERVS. DE ASSESSORIA E CONSULTORIA": "#10b981",
    "SERVS. DE CONTABILIDADE":            "#14b8a6",
    "SERVICO PRESTADOS POR TERCEIROS":    "#a855f7",
    "CSRF":                               "#64748b",
    "PREMIOS E GRATIFICACOES ADM":        "#d97706",
    "SERVS. ADMINISTRATIVOS":             "#475569",
    "SISTEMAS OPERACIONAIS":              "#f97316",
    "ISS":                                "#be185d",
    "REEMB DESPESAS GERAIS":              "#65a30d",
    "SERVS. CERTIFICADO DIGITAL":         "#b45309",
    "ESTORNO_DEVOLUCAO (-)":              "#9ca3af",
    "COMPRA DE ATIVO IMOBILIZADO":        "#0284c7",
}

CC_COLORS = {
    "TECNOLOGIA - DESENVOLVIMENTO": "#8b5cf6",
    "MKT - INFLUENCERS":            "#f59e0b",
    "ADMINISTRATIVO":               "#14b8a6",
    "SOCIAL MEDIA - DESIGN":        "#ec4899",
    "OPERACIONAL":                  "#10b981",
    "COMMUNITY MANAGER":            "#f97316",
    "TRANSITORIO":                  "#9ca3af",
}

# Subgrupos a excluir (movimentacoes financeiras, nao operacionais)
EXCL_SUBGRUPOS = ["INVESTIMENTO CDI", "CONTA INVESTIMENTO", "ADIANTAMENTO APORTE"]

# ============================================================
#  FUNCAO PRINCIPAL
# ============================================================
def main():
    print("\n" + "=" * 58)
    print("  CodeCraft - Atualizador de Dashboard")
    print("=" * 58)

    excel_path = Path(EXCEL_PATH)
    html_path  = Path(HTML_PATH)

    # Verificar arquivos
    if not excel_path.exists():
        print("\nERRO: Excel nao encontrado:")
        print("  " + str(EXCEL_PATH))
        pausar(); sys.exit(1)
    if not html_path.exists():
        print("\nERRO: HTML nao encontrado:")
        print("  " + str(HTML_PATH))
        pausar(); sys.exit(1)

    print(f"\n  Excel : {excel_path.name}")
    print(f"  HTML  : {html_path.name}")

    # Backup
    data_hoje = datetime.date.today().strftime("%Y%m%d")
    backup_dir = html_path.parent / "_backups_dashboard"
    backup_dir.mkdir(exist_ok=True)
    backup = backup_dir / f"{html_path.stem}.backup_{data_hoje}.html"
    shutil.copy2(html_path, backup)
    print(f"  Backup: _backups_dashboard/{backup.name}")

    # -- Ler Excel --------------------------------------------
    print("\n  Lendo Excel...")
    try:
        df = pd.read_excel(str(excel_path), sheet_name=SHEET_NAME)
    except Exception as e:
        print(f"\nERRO ao abrir Excel: {e}")
        pausar(); sys.exit(1)

    # Colunas de data
    df["Data Financ"] = pd.to_datetime(df["Data Financ"], errors="coerce")
    df["mes_fin"] = df["Data Financ"].dt.to_period("M").astype(str)

    # Normalizar banco e categoria
    df["banco_norm"] = (df["Banco"].astype(str)
                        .str.encode("ascii","ignore")
                        .str.decode("ascii")
                        .str.upper()
                        .str.strip())
    df["cat_norm"] = df["Categoria"].apply(norm_cat)

    # Meses realizados = todos com Data Financ preenchida
    # Realizados = todos com Data Financ preenchida
    meses_reais = sorted(
        m for m in df["mes_fin"].dropna().unique()
        if len(m) == 7 and m[4] == "-"
    )
    ultimo_real = meses_reais[-1]
    meses_prev = []  # sem previstos

    print(f"  Realizados: {meses_reais[0]} a {ultimo_real} ({len(meses_reais)} meses)")

    # -- Custodia MP: creditos e saques -----------------------
    # Creditos = Receita MP com Subgrupo SALDO RECEITA
    # Saques   = Despesa MP com Subgrupo SALDO RECEITA
    mp = df[df["banco_norm"] == "MERCADO PAGO"]
    cred_m  = (mp[(mp["Tipo"] == "Receita") &
                  (mp["Subgrupo"].str.contains("SALDO RECEITA", na=False))]
               .groupby("mes_fin")["Valor"].sum())
    saques_m = (mp[(mp["Tipo"] == "Despesa") &
                   (mp["cat_norm"].str.contains("SAQUE GAMERS", na=False))]
                .groupby("mes_fin")["Valor"].sum().abs())

    # -- Receita Inter: APENAS TAXA DE ADMINISTRACAO (+) ------
    rec_taxa = (df[(df["banco_norm"] == "BANCO INTER") &
                   (df["Tipo"] == "Receita") &
                   (df["cat_norm"].str.contains("TAXA DE ADMINISTRA"))]
                .groupby("mes_fin")["Valor"].sum())

    rec_detail = []
    for _, r in df[(df["banco_norm"] == "BANCO INTER") &
                   (df["Tipo"] == "Receita") &
                   (df["cat_norm"].str.contains("TAXA DE ADMINISTRA"))].iterrows():
        desc_col = "Descricao" if "Descricao" in df.columns else ("Descricao" if "Descricao" in df.columns else "")
        desc = str(r.get(desc_col, "")) if desc_col else ""
        rec_detail.append({
            "data_fin": str(r["Data Financ"])[:10],
            "banco":    str(r.get("Banco", "")),
            "subgrupo": str(r.get("Subgrupo", "")),
            "descricao": desc,
            "val":      rnd(r["Valor"]),
            "status":   "Recebido"
        })

    # -- 20% devida: 20% do credito MP de cada mes ------------
    # Logica: 20% do credito bruto MP acumulado por mes
    rec20 = {}
    for mes in meses_reais:
        c = float(cred_m.get(mes, 0))
        if c > 0:
            rec20[mes] = rnd(c * 0.20)
    total_rec20 = rnd(sum(rec20.values()))

    # -- Despesas B4: Inter + Cartao + Bruno (so Despesas) ----
    # Bruno: apenas Despesas (excluir Receita INVESTIMENTO BRUNO = ficticio)
    # Inter e Cartao: excluir subgrupos de CDI/investimento
    desp_b4 = df[
        (df["Tipo"] == "Despesa") &
        (df["banco_norm"].isin(["BANCO INTER", "CARTO DE CREDITO INTER", "BRUNO NIRO"])) &
        (~df["Subgrupo"].isin(EXCL_SUBGRUPOS)) &
        (df["mes_fin"].isin(meses_reais))
    ].copy()

    def gr(dff, cats):
        return dff[dff["cat_norm"].isin(cats)].groupby("mes_fin")["Valor"].sum().abs()

    dp = gr(desp_b4, CAT_PESSOAL)
    dm = gr(desp_b4, CAT_MKT)
    dj = gr(desp_b4, CAT_JURIDICO)
    do = gr(desp_b4, CAT_OUTROS)

    # Sem previstos - nao usa Orcado
    dpp = {}; dpm = {}; dpj = {}; dpo = {}
    DB_PREV = {}

    # -- Taxa ADM 20%% paga ao Inter (saida do MP) ----------
    taxa_mp = df[
        (df["banco_norm"] == "MERCADO PAGO") &
        (df["cat_norm"] == "TAXA ADM - 20% (-)")
    ].groupby("mes_fin")["Valor"].sum().abs()
    DB_TAXA_PAGA_RAW = {m: round(float(v),2) for m,v in taxa_mp.items()}

    # -- Outras taxas MP (CSRF, ISS, Estornos, Taxas plataforma etc) ----
    outras_mp = df[
        (df["banco_norm"] == "MERCADO PAGO") &
        (df["Tipo"] == "Despesa") &
        (~df["cat_norm"].str.contains("SAQUE GAMERS")) &
        (~df["cat_norm"].str.contains("TAXA ADM"))
    ].groupby("mes_fin")["Valor"].sum().abs()
    DB_OUTRAS_TAXAS_RAW = {m: round(float(v),2) for m,v in outras_mp.items()}

    # -- B2_ROWS: Capital & Investimento (BANCO INTER - APLICACAO) ------
    inv_df   = df[df["banco_norm"].str.contains("APLICAC", na=False)]
    aplic_df = df[(df["banco_norm"] == "BANCO INTER") & (df["Categoria"] == "APLIC FINANCEIRA (-)")]
    resg_df  = df[(df["banco_norm"] == "BANCO INTER") & (df["Categoria"] == "RESGATE APLIC (+)")]
    b2_meses = sorted(set(
        list(inv_df["mes_fin"].dropna().unique()) +
        list(aplic_df["mes_fin"].dropna().unique()) +
        list(resg_df["mes_fin"].dropna().unique())
    ))
    mnames = {"01":"jan","02":"fev","03":"mar","04":"abr","05":"mai","06":"jun",
              "07":"jul","08":"ago","09":"set","10":"out","11":"nov","12":"dez"}
    b2_rows_list = []
    for mes in b2_meses:
        month  = mes[5:]
        year   = mes[:4]
        aplic  = round(float(aplic_df[aplic_df["mes_fin"]==mes]["Valor"].abs().sum()), 0)
        resg   = round(float(resg_df[resg_df["mes_fin"]==mes]["Valor"].sum()), 0)
        grp    = inv_df[inv_df["mes_fin"] == mes]
        rend_b = round(float(grp[grp["Categoria"] == "RENDIMENTO INVEST. (+)"]["Valor"].sum()), 2)
        irf    = round(float(grp[grp["Categoria"].str.contains("IRRF-IOF|PREVISAO IR", na=False)]["Valor"].sum()), 2)
        rend_liq = int(round(rend_b + irf, 0))
        key   = mnames[month] + year[2:]
        label = mnames[month].capitalize() + "/" + year[2:]
        rend_b_int = int(round(rend_b, 0))
        irf_int    = int(round(irf, 0))
        b2_rows_list.append(
            "{key:'" + key + "', label:'" + label + "', year:'" + year + "', month:'" + month +
            "', aplicacao:" + str(int(aplic)) + ", resgate:" + str(int(resg)) +
            ", rend_b:" + str(rend_b_int) + ", irf:" + str(irf_int) + ", rend:" + str(rend_liq) + "}"
        )
    B2_ROWS_JS = "const B2_ROWS = [\n  " + ",\n  ".join(b2_rows_list) + "\n];"

    # -- DB_CAT: despesas por categoria e mes -----------------
    DB_CAT_RAW = {}
    for _, row in desp_b4.iterrows():
        cat = row["cat_norm"]
        mes = row["mes_fin"]
        val = abs(float(row["Valor"] or 0))
        if cat not in DB_CAT_RAW:
            DB_CAT_RAW[cat] = {}
        DB_CAT_RAW[cat][mes] = rnd(DB_CAT_RAW[cat].get(mes, 0) + val)

    # -- DB_CC: despesas por Centro de Custo e mes ------------------
    DB_CC_RAW = {}
    for _, row in desp_b4.iterrows():
        cc  = norm_cat(str(row.get("Centro de Custo", "") or "SEM CC"))
        mes = row["mes_fin"]
        val = abs(float(row["Valor"] or 0))
        if cc not in DB_CC_RAW: DB_CC_RAW[cc] = {}
        DB_CC_RAW[cc][mes] = rnd(DB_CC_RAW[cc].get(mes, 0) + val)

    # -- Montar DB principal -----------------------------------
    print("\n  Calculando dados...")
    DB = {}
    for mes in meses_reais:
        p  = rnd(dp.get(mes, 0))
        m  = rnd(dm.get(mes, 0))
        j  = rnd(dj.get(mes, 0))
        o  = rnd(do.get(mes, 0))
        DB[mes] = {
            "cred":          rnd(cred_m.get(mes, 0)),
            "saques":        -rnd(saques_m.get(mes, 0)),
            "rec_inter":     rnd(rec_taxa.get(mes, 0)),
            "desp_pessoal":  p,
            "desp_mkt":      m,
            "desp_juridico": j,
            "desp_outros":   o,
            "desp_total":    rnd(p + m + j + o)
        }

    # -- DB_PREVISTO -------------------------------------------
    DB_PREV = {}
    for mes in meses_prev:
        pp = rnd(dpp.get(mes, 0))
        pm = rnd(dpm.get(mes, 0))
        pj = rnd(dpj.get(mes, 0))
        po = rnd(dpo.get(mes, 0))
        DB_PREV[mes] = {
            "desp_pessoal":  pp,
            "desp_mkt":      pm,
            "desp_juridico": pj,
            "desp_outros":   po,
            "desp_total":    rnd(pp + pm + pj + po)
        }

    # -- Resumo ------------------------------------------------
    tot_rec  = sum(v["rec_inter"] for v in DB.values())
    tot_desp = sum(v["desp_total"] for v in DB.values())
    n_real   = sum(1 for v in DB.values() if v["desp_total"] > 0)

    # -- Gerar bloco JS ----------------------------------------
    def db_line(mes, d):
        return ('  "%s":{"cred":%s,"saques":%s,"rec_inter":%s,'
                '"desp_pessoal":%s,"desp_mkt":%s,"desp_juridico":%s,'
                '"desp_outros":%s,"desp_total":%s}') % (
            mes, d["cred"], d["saques"], d["rec_inter"],
            d["desp_pessoal"], d["desp_mkt"], d["desp_juridico"],
            d["desp_outros"], d["desp_total"])

    def prev_line(mes, d):
        return ('  "%s":{"desp_pessoal":%s,"desp_mkt":%s,'
                '"desp_juridico":%s,"desp_outros":%s,"desp_total":%s}') % (
            mes, d["desp_pessoal"], d["desp_mkt"],
            d["desp_juridico"], d["desp_outros"], d["desp_total"])

    cat_lines = []
    for cat, mdata in sorted(DB_CAT_RAW.items(), key=lambda x: -sum(x[1].values())):
        parts = ", ".join('"%s":%s' % (m, v) for m, v in sorted(mdata.items()))
        cat_lines.append('  "%s": {%s}' % (cat, parts))

    col_lines = ['  "%s": "%s"' % (k, v) for k, v in CAT_COLORS.items()]

    rec20_lines = ['  "%s": %s' % (k, v) for k, v in sorted(rec20.items())]

    rec_detail_js = json.dumps(rec_detail, indent=2, ensure_ascii=False)

    hoje_str = datetime.date.today().strftime("%d/%m/%Y")

    bloco = "\n".join([
        "// ============================================================",
        "// DATA BLOCK - update with gerar_dados_dashboard.py",
        "// Gerado em: %s  |  Arquivo: %s" % (hoje_str, excel_path.name),
        "// ============================================================",
        "const DB = {",
        ",\n".join(db_line(m, d) for m, d in sorted(DB.items())),
        "};",
        "",
        "// Receita Inter - somente TAXA DE ADMINISTRACAO (+)",
        "const REC_INTER_DETAIL = %s;" % rec_detail_js,
        "",
        "// 20% devida por mes (20% do credito bruto MP de cada mes)",
        "const REC20_2025 = {",
        ",\n".join(rec20_lines),
        "};",
        "const TOTAL_REC20_2025 = %s;" % total_rec20,
        "",
        "// Despesas PREVISTAS - coluna Orcado",
        "const DB_PREVISTO = {",
        ",\n".join(prev_line(m, d) for m, d in sorted(DB_PREV.items())),
        "};",
        "",
        "// DB_CAT: despesas por categoria (Inter + Cartao + Bruno)",
        "const DB_CAT = {",
        ",\n".join(cat_lines),
        "};",
        "",
        "// Cores por categoria",
        "const CAT_COLORS = {",
        ",\n".join(col_lines),
        "};",
        "",
        "// DB_CC: despesas por Centro de Custo e mes",
        "const DB_CC = {",
        ",\n".join(
            '  "%s": {%s}' % (cc, ", ".join('"%s":%s' % (m,v) for m,v in sorted(mdata.items())))
            for cc, mdata in sorted(DB_CC_RAW.items(), key=lambda x: -sum(x[1].values()))
        ),
        "};",
        "",
        "// Cores por Centro de Custo",
        "const CC_COLORS = {",
        ",\n".join('  "%s": "%s"' % (k,v) for k,v in CC_COLORS.items()),
        "};",
        "",
        "// Taxa ADM 20%% efetivamente paga ao Inter via MP",
        "const DB_TAXA_PAGA = {%s};" % ", ".join('"%s":%s' % (m, round(float(v),2)) for m,v in sorted(DB_TAXA_PAGA_RAW.items())),
        "const DB_OUTRAS_TAXAS = {%s};" % ", ".join('"%s":%s' % (m, round(float(v),2)) for m,v in sorted(DB_OUTRAS_TAXAS_RAW.items())),
        "",
        B2_ROWS_JS,
        "",
    ])

    # -- Atualizar HTML ----------------------------------------
    print("  Atualizando HTML...")
    texto = html_path.read_text(encoding="utf-8")
    idx_s = texto.find(MARKER_START)
    idx_e = texto.find(MARKER_END, idx_s)
    if idx_s == -1 or idx_e == -1:
        print("\nERRO: Marcadores nao encontrados no HTML.")
        print("  Certifique-se de usar o index.html correto.")
        pausar(); sys.exit(1)

    html_novo = texto[:idx_s] + bloco + texto[idx_e:]
    html_path.write_text(html_novo, encoding="utf-8")

    # -- Resumo final ------------------------------------------
    print("\n  " + "=" * 54)
    print("  HTML atualizado com sucesso!")
    print("  " + "=" * 54)
    print("  Receita Inter (TAXA ADM):  %s" % fmt_brl(tot_rec))
    print("  20%% devida total:          %s" % fmt_brl(total_rec20))
    print("  Despesa realizada:         %s" % fmt_brl(tot_desp))
    print("  Meses realizados:          %d" % len(meses_reais))
    print("  Meses previstos:           %d" % len(meses_prev))
    print("  " + "=" * 54)
    pausar()


if __name__ == "__main__":
    main()
