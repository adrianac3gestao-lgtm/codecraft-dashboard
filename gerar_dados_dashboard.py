#!/usr/bin/env python3
"""
gerar_dados_dashboard.py
------------------------------------------------------------
COMO USAR - 3 passos simples:

  1. Abra este arquivo no Bloco de Notas ou qualquer editor
  2. Preencha EXCEL_PATH e HTML_PATH logo abaixo
  3. De dois cliques no arquivo (ou execute no terminal:
         python gerar_dados_dashboard.py)

  O script atualiza o HTML automaticamente.
  Nao e necessario copiar nada do terminal.

INSTALAR DEPENDENCIAS (apenas uma vez):
  Abra o Prompt de Comando e execute:
      pip install pandas openpyxl
------------------------------------------------------------
"""

# +---------------------------------------------------------+
# |  >  PREENCHA AQUI - apenas estas 2 linhas               |
# +---------------------------------------------------------+

EXCEL_PATH = r"C:\Users\adria\OneDrive\C3 Gestao\EquipeC3\CODECRAFT\Relatorio Gerencial_Codecraft2026.xlsx"
HTML_PATH  = r"C:\Users\adria\OneDrive\C3 Gestao\EquipeC3\CODECRAFT\7. Dashboard\index.html"

# ----------------------------------------------------------
#  Nao altere nada abaixo desta linha
# ----------------------------------------------------------

import json, sys, shutil, datetime
import io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

def pausar():
    """So pausa se rodando interativamente (nao via bat/agendador)"""
    try:
        if sys.stdin.isatty():
            pausar()
    except Exception:
        pass
from pathlib import Path

try:
    import pandas as pd
except ImportError:
    print("\nX pandas nao esta instalado.")
    print("   Abra o Prompt de Comando e execute:")
    print("       pip install pandas openpyxl\n")
    pausar()
    sys.exit(1)

SHEET_NAME = "4.Consolidado"

CAT_PESSOAL  = ['SALARIOS E ORDENADOS ADM', 'PREMIOS E GRATIFICACOES ADM']
CAT_MKT      = ['SERVS. DE PUBLICIDADE E PROPAGANDA']
CAT_JURIDICO = ['SERVS. ADVOCATICIOS']
CAT_OUTROS   = [
    'SERVS. DE ASSESSORIA E CONSULTORIA', 'SERVICO PRESTADOS POR TERCEIROS',
    'SERVS. DE CONTABILIDADE', 'SISTEMAS OPERACIONAIS', 'SERVS. CERTIFICADO DIGITAL',
    'COMPRA DE ATIVO IMOBILIZADO', 'ISS', 'CSRF',
    'COFINS Retido sobre Pagamentos', 'IRPJ Retido sobre Pagamentos',
    'CSLL Retido sobre Pagamentos', 'PIS Retido sobre Pagamentos',
    'PREVISAO IR - IOF INVEST (-)', 'IRRF-IOF APLIC. (-)',
    'REEMB DESPESAS GERAIS', 'ESTORNO_DEVOLUCAO (-)', 'SERVS. ADMINISTRATIVOS'
]
EXCL_SUBGRUPOS = [
    'INVESTIMENTO BRUNO', 'INVESTIMENTO CDI',
    'CONTA INVESTIMENTO', 'ADIANTAMENTO APORTE'
]

# Apenas bancos operacionais (Inter + Cartao Inter)
BANCOS_OP = ['BANCO INTER', 'CARTO DE CREDITO INTER']  # Cartao name after ASCII encoding

# Marcadores dentro do HTML que delimitam o bloco de dados
MARKER_START = "// DATA BLOCK - update with gerar_dados_dashboard.py"
MARKER_END   = "const LABEL_MAP"


def rnd(v):
    return round(float(v or 0), 2)


def fmt_brl(v):
    s = f"{abs(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"R$ {s}"


def gerar_bloco_js(DB, rec_detail, rec20, total_rec20, DB_PREV, DB_CAT_RAW, CAT_COLORS_MAP, excel_name):
    hoje = datetime.date.today().strftime('%d/%m/%Y')
    linhas = [
        "// ============================================================",
        "// DATA BLOCK - update with gerar_dados_dashboard.py",
        f"// Gerado em: {hoje}  |  Arquivo: {excel_name}",
        "// ============================================================",
        f"const DB = {json.dumps(DB, indent=2, ensure_ascii=False)};",
        "",
        "// Receita Inter - somente TAXA DE ADMINISTRACAO (+)",
        f"const REC_INTER_DETAIL = {json.dumps(rec_detail, indent=2, ensure_ascii=False)};",
        "",
        "// 20% devida 2025 (por competencia, nao recebida no Inter)",
        f"const REC20_2025 = {json.dumps(rec20, indent=2, ensure_ascii=False)};",
        f"const TOTAL_REC20_2025 = {total_rec20};",
        "",
        "// Despesas PREVISTAS - coluna Orcado (datas futuras)",
        f"const DB_PREVISTO = {json.dumps(DB_PREV, indent=2, ensure_ascii=False)};",
        "",
        "// DB_CATEGORIAS - despesas reais por categoria e mes (Inter + Cartao)",
        f"const DB_CAT = {json.dumps(DB_CAT_RAW, indent=2, ensure_ascii=True)};",
        "",
        "// Category color map",
        f"const CAT_COLORS = {json.dumps(CAT_COLORS_MAP, indent=2, ensure_ascii=True)};",
        "",
    ]
    return "\n".join(linhas)


def atualizar_html(html_path, novo_bloco):
    texto = html_path.read_text(encoding='utf-8')
    idx_s = texto.find(MARKER_START)
    idx_e = texto.find(MARKER_END, idx_s)
    if idx_s == -1 or idx_e == -1:
        return False
    html_path.write_text(texto[:idx_s] + novo_bloco + texto[idx_e:], encoding='utf-8')
    return True


def main():
    import unicodedata
    def norm(s):
        s2 = unicodedata.normalize('NFD', str(s))
        s2 = ''.join(c for c in s2 if unicodedata.category(c) != 'Mn')
        return s2.upper().strip()

        print("\n" + "=" * 58)
    print("  CodeCraft - Atualizador de Dashboard")
    print("=" * 58)

    excel_path = Path(EXCEL_PATH)
    html_path  = Path(HTML_PATH)

    # -- Verificar se os arquivos existem ---------------------
    if not excel_path.exists():
        print(f"\nX Arquivo Excel nao encontrado:")
        print(f"   {EXCEL_PATH}")
        print("\n   Corrija o caminho em EXCEL_PATH no inicio do script.")
        pausar()
        sys.exit(1)

    if not html_path.exists():
        print(f"\nX Arquivo HTML nao encontrado:")
        print(f"   {HTML_PATH}")
        print("\n   Corrija o caminho em HTML_PATH no inicio do script.")
        pausar()
        sys.exit(1)

    print(f"\n? Excel : {excel_path.name}")
    print(f"? HTML  : {html_path.name}")

    # -- Backup automatico numa subpasta _backups --------------
    # (nao cria arquivo extra ao lado do HTML original)
    data_hoje = datetime.date.today().strftime('%Y%m%d')
    backup_dir = html_path.parent / "_backups_dashboard"
    backup_dir.mkdir(exist_ok=True)
    backup = backup_dir / f"{html_path.stem}.backup_{data_hoje}.html"
    shutil.copy2(html_path, backup)
    print(f"? Backup: _backups_dashboard\\{backup.name}")

    # -- Ler Excel --------------------------------------------
    print("\n? Lendo Excel...")
    try:
        df = pd.read_excel(str(excel_path), sheet_name=SHEET_NAME)
    except Exception as e:
        print(f"\nX Erro ao abrir o Excel: {e}")
        pausar()
        sys.exit(1)

    df['Data Financ'] = pd.to_datetime(df['Data Financ'], errors='coerce')
    # Detecta coluna Orcado independente de acento
    col_orc = next((c for c in df.columns if c.encode('ascii','ignore').decode().lower() in ['orcado','or?ado']), None)
    if col_orc is None:
        col_orc = next((c for c in df.columns if 'ado' in c.lower() and c.lower().startswith('or')), None)
    if col_orc is None:
        raise KeyError("Coluna Orcado nao encontrada. Colunas disponiveis: " + str(list(df.columns)))
    df['Orcado'] = pd.to_datetime(df[col_orc], errors='coerce')
    df['mes_fin']     = df['Data Financ'].dt.to_period('M').astype(str)
    df['mes_orc']     = df['Orcado'].dt.to_period('M').astype(str)

    hoje_mes = datetime.date.today().strftime('%Y-%m')

    meses_reais = sorted(
        m for m in df['mes_fin'].dropna().unique()
        if len(m) == 7 and m[4] == '-' and m <= hoje_mes
    )
    ultimo_real = meses_reais[-1]

    meses_previstos = sorted(
        m for m in df['mes_orc'].dropna().unique()
        if len(m) == 7 and m[4] == '-' and m > ultimo_real
    )

    print(f"? Realizados : {meses_reais[0]} ? {ultimo_real}  ({len(meses_reais)} meses)")
    prev_resumo = ', '.join(meses_previstos[:4]) + ('...' if len(meses_previstos) > 4 else '')
    print(f"? Previstos  : {len(meses_previstos)} meses  ({prev_resumo})")

    # -- Custodia ---------------------------------------------
    cred_m  = df[df['Categoria'].str.contains('CREDITO_CUSTODIA|OUTRAS RECEITAS_TESTE', na=False) &
                 df['mes_fin'].isin(meses_reais)].groupby('mes_fin')['Valor'].sum()
    saques_m= df[df['Categoria'].str.contains('SAQUE GAMERS', na=False) &
                 df['mes_fin'].isin(meses_reais)].groupby('mes_fin')['Valor'].sum()

    # -- Receita Inter - TAXA ADM (+) ------------------------
    inter_taxa = df[
        (df['Banco'] == 'BANCO INTER') &
        (df['Tipo']  == 'Receita') &
        (df['Categoria'].astype(str).str.encode('ascii','ignore').str.decode('ascii').str.upper().str.contains('TAXA DE ADMINISTRA')) &
        df['mes_fin'].isin(meses_reais)
    ]
    rec_inter_m = inter_taxa.groupby('mes_fin')['Valor'].sum()
    rec_detail  = [
        {
            'data_fin': str(r['Data Financ'])[:10],
            'banco':    str(r.get('Banco', '')),
            'subgrupo': str(r.get('Subgrupo', '')),
            'descricao':str(r.get('Descricao', '')),
            'val':      rnd(r['Valor']),
            'status':   'Recebido'
        }
        for _, r in inter_taxa.iterrows()
    ]

    # -- Despesas realizadas ----------------------------------
    # Normaliza nome do banco (remove acento para comparar)
    df['Banco_norm'] = df['Banco'].astype(str).str.encode('ascii','ignore').str.decode('ascii').str.upper()
    desp_r = df[
        (df['Tipo'] == 'Despesa') &
        (~df['Subgrupo'].isin(EXCL_SUBGRUPOS)) &
        (df['Banco_norm'].isin(['BANCO INTER', 'CARTO DE CREDITO INTER'])) &
        df['mes_fin'].isin(meses_reais)
    ]

    def gr(df_, cats, col='mes_fin'):
        # Use normalized category column for accent-insensitive matching
        if 'cat_norm' in df_.columns:
            return df_[df_['cat_norm'].isin(cats)].groupby(col)['Valor'].sum().abs()
        return df_[df_['Categoria'].isin(cats)].groupby(col)['Valor'].sum().abs()

    dp = gr(desp_r, CAT_PESSOAL)
    dm = gr(desp_r, CAT_MKT)
    dj = gr(desp_r, CAT_JURIDICO)
    do = gr(desp_r, CAT_OUTROS)

    # -- Despesas previstas -----------------------------------
    desp_p = df[
        (df['Tipo'] == 'Despesa') &
        (~df['Subgrupo'].isin(EXCL_SUBGRUPOS)) &
        (df['Banco_norm'].isin(['BANCO INTER', 'CARTO DE CREDITO INTER'])) &
        df['mes_orc'].isin(meses_previstos)
    ].copy()
    desp_p['cat_norm'] = desp_p['Categoria'].apply(norm)
    dpp = gr(desp_p, CAT_PESSOAL,  'mes_orc')
    dpm = gr(desp_p, CAT_MKT,      'mes_orc')
    dpj = gr(desp_p, CAT_JURIDICO, 'mes_orc')
    dpo = gr(desp_p, CAT_OUTROS,   'mes_orc')

    # -- 20% devida 2025 --------------------------------------
    rec20    = {m: rnd(cred_m.get(m, 0) * 0.20) for m in meses_reais if m.startswith('2025')}
    tot_r20  = rnd(sum(rec20.values()))

    # -- DB_CAT: despesas reais por categoria e mes (Inter + Cartao) --------
    import unicodedata
    def norm(s):
        s2 = unicodedata.normalize('NFD', str(s))
        s2 = ''.join(c for c in s2 if unicodedata.category(c) != 'Mn')
        return s2.upper().replace('(','(').strip()

    cat_mes = desp_r.copy()
    cat_mes['cat_norm'] = cat_mes['Categoria'].apply(norm)
    DB_CAT_RAW = {}
    for _, row in cat_mes.iterrows():
        cat = row['cat_norm']
        mes = row['mes_fin']
        val = abs(float(row['Valor'] or 0))
        if cat not in DB_CAT_RAW:
            DB_CAT_RAW[cat] = {}
        DB_CAT_RAW[cat][mes] = round(DB_CAT_RAW[cat].get(mes, 0) + val, 2)

    CAT_COLORS_MAP = {
        'SALARIOS E ORDENADOS ADM':           '#f59e0b',
        'SERVS. DE PUBLICIDADE E PROPAGANDA': '#8b5cf6',
        'SERVS. ADVOCATICIOS':                '#ef4444',
        'SERVS. DE ASSESSORIA E CONSULTORIA': '#10b981',
        'SERVS. DE CONTABILIDADE':            '#14b8a6',
        'SERVICO PRESTADOS POR TERCEIROS':    '#a855f7',
        'CSRF':                               '#64748b',
        'PREMIOS E GRATIFICACOES ADM':        '#d97706',
        'SERVS. ADMINISTRATIVOS':             '#475569',
        'SISTEMAS OPERACIONAIS':              '#f97316',
        'ISS':                                '#be185d',
        'REEMB DESPESAS GERAIS':              '#65a30d',
        'SERVS. CERTIFICADO DIGITAL':         '#b45309',
        'ESTORNO_DEVOLUCAO (-)':              '#9ca3af',
        'COMPRA DE ATIVO IMOBILIZADO':        '#0284c7',
    }

    # -- Montar DB --------------------------------------------
    DB = {}
    for m in meses_reais:
        p = rnd(dp.get(m, 0)); mk = rnd(dm.get(m, 0))
        j = rnd(dj.get(m, 0)); o  = rnd(do.get(m, 0))
        DB[m] = {
            'cred':          rnd(cred_m.get(m, 0)),
            'saques':        rnd(saques_m.get(m, 0)),
            'rec_inter':     rnd(rec_inter_m.get(m, 0)),
            'desp_pessoal':  p, 'desp_mkt': mk,
            'desp_juridico': j, 'desp_outros': o,
            'desp_total':    rnd(p + mk + j + o)
        }

    DB_PREV = {}
    for m in meses_previstos:
        p = rnd(dpp.get(m, 0)); mk = rnd(dpm.get(m, 0))
        j = rnd(dpj.get(m, 0)); o  = rnd(dpo.get(m, 0))
        DB_PREV[m] = {
            'desp_pessoal': p, 'desp_mkt': mk,
            'desp_juridico': j, 'desp_outros': o,
            'desp_total': rnd(p + mk + j + o)
        }

    # -- Atualizar HTML ---------------------------------------
    print("\n??  Calculando dados...")
    bloco = gerar_bloco_js(DB, rec_detail, rec20, tot_r20, DB_PREV, DB_CAT_RAW, CAT_COLORS_MAP, excel_path.name)

    print("??  Atualizando HTML...")
    ok = atualizar_html(html_path, bloco)

    if ok:
        print(f"\nOK HTML atualizado com sucesso!")
        print(f"   ? Abra o arquivo no navegador para ver os novos dados.")
    else:
        # Fallback: salvar bloco em arquivo .js separado
        js_out = html_path.parent / "bloco_dados_gerado.txt"
        js_out.write_text(bloco, encoding='utf-8')
        print(f"\n??  Nao foi possivel atualizar o HTML automaticamente.")
        print(f"   O bloco de dados foi salvo em:")
        print(f"   ? {js_out}")
        print(f"\n   Para atualizar manualmente:")
        print(f"   1. Abra o HTML em um editor de texto (ex: Notepad++)")
        print(f"   2. Use Ctrl+H (Localizar e Substituir)")
        print(f"   3. Procure por:  // DB_REALIZADO - dados por Data Financ")
        print(f"   4. Substitua o bloco (ate 'const LABEL_MAP') pelo conteudo de bloco_dados_gerado.txt")

    # -- Resumo -----------------------------------------------
    tot_rec  = sum(v['rec_inter'] for v in DB.values())
    tot_desp = sum(v['desp_total'] for v in DB.values())
    tot_prev = sum(v['desp_total'] for v in DB_PREV.values())

    print("\n" + "-" * 58)
    print("  RESUMO")
    print("-" * 58)
    print(f"  Receita Inter (TAXA ADM):   {fmt_brl(tot_rec):>18}")
    print(f"  20% devida 2025:            {fmt_brl(tot_r20):>18}")
    print(f"  Despesa realizada (total):  {fmt_brl(tot_desp):>18}")
    print(f"  Despesa prevista (total):   {fmt_brl(tot_prev):>18}")
    print(f"  Meses realizados:           {len(DB):>18}")
    print(f"  Meses previstos:            {len(DB_PREV):>18}")
    print("-" * 58)

    pausar()


if __name__ == '__main__':
    main()
