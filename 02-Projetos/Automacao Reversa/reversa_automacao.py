# -*- coding: utf-8 -*-
"""
Automação Orçado vs Realizado – Reversa
=======================================
Fluxo:
  Realizado Reversa (documento + série + filial)
      -> NEOGRID_CTE  (XML do CT-e: chave da NF-e de origem em infDoc/infNFe/chave)
      -> RE0501       (.txt com itens/quantidades da NF-e, pela mesma chave)
  Saída: Orc_vs_Real_Reversa_python.xlsx (na pasta RAIZ)

Uso:  python reversa_automacao.py            (ou dê duplo clique em rodar_reversa.bat)
Requisitos: pip install pandas openpyxl xlrd
"""
import os
import re
import sys
import zipfile
import pickle
import xml.etree.ElementTree as ET
from datetime import datetime

import pandas as pd
from openpyxl.styles import Font

# ============================== CONFIGURAÇÃO ==============================
CONFIG = {
    "RAIZ": r"K:\Logistica\CENTRAL DE FRETES\CUSTOS\Reversa",
    "PASTA_REALIZADO": "Realizado Reversa",
    "PASTA_CTE": "NEOGRID_CTE",
    "PASTA_RE0501": "RE0501",
    "ARQUIVO_SAIDA": "Orc_vs_Real_Reversa_python.xlsx",

    # Nomes possíveis das colunas (comparação sem acento/maiúsculas/espaços)
    "COLS_REALIZADO": {
        "documento": ["documento", "doc", "num documento", "nr documento", "numero documento", "nota", "conhecimento", "cte"],
        "serie":     ["serie", "serie documento", "serie doc", "ser"],
        "filial":    ["filial", "fil", "cod filial", "codigo filial", "estabelecimento", "estab"],
        "valor":     ["valor", "vlr", "valor total", "vl total", "total", "valor frete", "realizado"],
    },
    "COLS_RE0501": {
        "chave":     ["chave", "chave acesso", "chave nfe", "chave de acesso", "chave nf", "chv nfe"],
        "item":      ["item", "cod item", "codigo item", "produto", "cod produto", "codigo", "material", "sku"],
        "descricao": ["descricao", "desc item", "descricao item", "descricao produto", "desc"],
        "quantidade":["quantidade", "qtd", "qtde", "quant", "qtd item"],
        "unidade":   ["un", "unidade", "um", "unid"],
        "valor_item":["valor item", "vlr item", "valor total item", "vl item", "valor mercadoria"],
    },

    # Classificação do item (1ª regra que casar vence; regex sobre código + descrição)
    "REGRAS_CATEGORIA": [
        ("Pallet",    r"PALLET|PALETE|PALLETS|\bPBR\b|\bCHEP\b"),
        ("Embalagem", r"CAIXA|EMBALAG|BOMBONA|TAMBOR|CONTAINER|\bIBC\b|ENGRADADO|\bGAIOLA\b"),
    ],
    "CATEGORIA_PADRAO": "Produto",

    # Opcional: filial do relatório -> CNPJ (14 dígitos). Usado para desempatar
    # quando há mais de um CT-e com mesmo número/série. Se vazio, compara a
    # filial com o trecho de filial do CNPJ (posições 9-12, ex.: 0001).
    "MAPA_FILIAL_CNPJ": {
        # "01": "12345678000190",
    },
}
# =========================================================================

CHAVE_RE = re.compile(r"\d{44}")


def norm_txt(s):
    import unicodedata
    s = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode()
    return re.sub(r"[^a-z0-9]+", " ", s.lower()).strip()


def norm_num(v):
    """'000123' / 123.0 / ' 123 ' -> '123'"""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    s = str(v).strip()
    if re.fullmatch(r"\d+\.0+", s):
        s = s.split(".")[0]
    s = s.lstrip("0")
    return s or ("0" if str(v).strip() else "")


def achar_colunas(cols, aliases):
    """Mapeia nome lógico -> coluna real usando os aliases."""
    nc = {norm_txt(c): c for c in cols}
    out = {}
    for logico, nomes in aliases.items():
        for n in nomes:
            if norm_txt(n) in nc:
                out[logico] = nc[norm_txt(n)]
                break
    return out


def log(msg):
    print(f"[{datetime.now():%H:%M:%S}] {msg}", flush=True)


# ----------------------------- REALIZADO ---------------------------------
def ler_planilha_com_cabecalho(caminho, obrigatorias):
    """Lê xlsx/xls/csv procurando a linha de cabeçalho nas 30 primeiras linhas."""
    ext = os.path.splitext(caminho)[1].lower()
    if ext == ".csv" or ext == ".txt":
        bruto = pd.read_csv(caminho, sep=None, engine="python", header=None, dtype=str,
                            encoding="latin-1", on_bad_lines="skip")
        abas = {"csv": bruto}
    else:
        abas = pd.read_excel(caminho, sheet_name=None, header=None, dtype=str)
    resultados = []
    for aba, bruto in abas.items():
        for i in range(min(30, len(bruto))):
            linha = [str(x) for x in bruto.iloc[i].tolist()]
            m = achar_colunas(linha, obrigatorias)
            if {"documento", "serie", "filial"} <= set(m):
                df = bruto.iloc[i + 1:].copy()
                df.columns = [str(c).strip() for c in bruto.iloc[i].tolist()]
                df = df.loc[:, [c for c in df.columns if c and c != "nan"]]
                df = df.dropna(how="all")
                df["_arquivo"] = os.path.basename(caminho)
                df["_aba"] = aba
                resultados.append(df)
                break
    return resultados


def ler_realizado(pasta):
    frames = []
    for raiz, _, arqs in os.walk(pasta):
        for a in arqs:
            if a.startswith("~$") or not a.lower().endswith((".xlsx", ".xlsm", ".xls", ".csv")):
                continue
            p = os.path.join(raiz, a)
            try:
                frames += ler_planilha_com_cabecalho(p, CONFIG["COLS_REALIZADO"])
            except Exception as e:
                log(f"  ! erro lendo {a}: {e}")
    if not frames:
        raise SystemExit("Nenhum relatório com colunas Documento/Série/Filial encontrado em 'Realizado Reversa'.")
    df = pd.concat(frames, ignore_index=True, sort=False)
    m = achar_colunas(df.columns, CONFIG["COLS_REALIZADO"])
    df["_doc"] = df[m["documento"]].map(norm_num)
    df["_serie"] = df[m["serie"]].map(norm_num)
    df["_filial"] = df[m["filial"]].map(lambda x: str(x).strip() if pd.notna(x) else "")
    df = df[df["_doc"] != ""]
    if "valor" in m:
        df["_valor"] = pd.to_numeric(df[m["valor"]].astype(str).str.replace(r"\.(?=\d{3}(\D|$))", "", regex=True)
                                     .str.replace(",", "."), errors="coerce")
    else:
        df["_valor"] = pd.NA
    df.insert(0, "ID_Linha", range(1, len(df) + 1))
    log(f"Realizado: {len(df)} linhas | colunas usadas: {m}")
    return df


# ------------------------------- CT-e ------------------------------------
def _sem_ns(tag):
    return tag.split("}", 1)[-1]


def parse_cte(conteudo, origem):
    try:
        root = ET.fromstring(conteudo)
    except ET.ParseError:
        return None
    idx = {}
    for el in root.iter():
        idx.setdefault(_sem_ns(el.tag), []).append(el)

    def txt(tag, pai=None):
        base = pai.iter() if pai is not None else (e for e in idx.get(tag, []))
        for e in base:
            if _sem_ns(e.tag) == tag and e.text:
                return e.text.strip()
        return ""

    if "infCte" not in idx:
        return None
    inf = idx["infCte"][0]
    ide = idx.get("ide", [None])[0]
    chave_cte = re.sub(r"\D", "", inf.attrib.get("Id", "")) or txt("chCTe")

    def cnpj_de(bloco):
        el = idx.get(bloco, [None])[0]
        return (txt("CNPJ", el) or txt("CPF", el)) if el is not None else ""

    # tomador: toma3/toma4 indicam quem é (0 rem,1 exped,2 receb,3 dest,4 outros)
    toma = txt("toma")
    mapa_toma = {"0": "rem", "1": "exped", "2": "receb", "3": "dest"}
    cnpj_toma = cnpj_de(mapa_toma[toma]) if toma in mapa_toma else cnpj_de("toma4")

    nfs = [e.text.strip() for e in idx.get("chave", []) if e.text and CHAVE_RE.fullmatch(e.text.strip())]
    # infNFe/chave (CT-e normal) – algumas versões usam chNFe
    nfs += [e.text.strip() for e in idx.get("chNFe", []) if e.text and CHAVE_RE.fullmatch(e.text.strip())]
    nfs = list(dict.fromkeys(n for n in nfs if n != chave_cte))

    return {
        "Chave_CTe": chave_cte,
        "Num_CTe": norm_num(txt("nCT", ide)),
        "Serie_CTe": norm_num(txt("serie", ide)),
        "Emissao_CTe": (txt("dhEmi", ide) or txt("dEmi", ide))[:10],
        "CNPJ_Emitente": cnpj_de("emit"),
        "CNPJ_Remetente": cnpj_de("rem"),
        "CNPJ_Destinatario": cnpj_de("dest"),
        "CNPJ_Tomador": cnpj_toma,
        "Valor_CTe": pd.to_numeric(txt("vTPrest"), errors="coerce"),
        "Chaves_NF_Origem": nfs,
        "Arquivo_CTe": origem,
    }


def ler_ctes(pasta, cache_path):
    cache = {}
    if os.path.exists(cache_path):
        try:
            with open(cache_path, "rb") as f:
                cache = pickle.load(f)
        except Exception:
            cache = {}
    novo_cache, regs, novos = {}, [], 0
    for raiz, _, arqs in os.walk(pasta):
        for a in arqs:
            p = os.path.join(raiz, a)
            low = a.lower()
            if not low.endswith((".xml", ".zip")):
                continue
            try:
                assin = (os.path.getsize(p), os.path.getmtime(p))
            except OSError:
                continue
            if p in cache and cache[p][0] == assin:
                itens = cache[p][1]
            else:
                novos += 1
                itens = []
                try:
                    if low.endswith(".zip"):
                        with zipfile.ZipFile(p) as z:
                            for n in z.namelist():
                                if n.lower().endswith(".xml"):
                                    r = parse_cte(z.read(n), f"{a}::{n}")
                                    if r:
                                        itens.append(r)
                    else:
                        with open(p, "rb") as f:
                            r = parse_cte(f.read(), a)
                        if r:
                            itens.append(r)
                except Exception as e:
                    log(f"  ! erro no CT-e {a}: {e}")
            novo_cache[p] = (assin, itens)
            regs += itens
    with open(cache_path, "wb") as f:
        pickle.dump(novo_cache, f)
    df = pd.DataFrame(regs)
    if df.empty:
        raise SystemExit("Nenhum XML de CT-e válido em NEOGRID_CTE.")
    df = df.drop_duplicates("Chave_CTe")
    log(f"CT-e: {len(df)} XMLs ({novos} novos/alterados lidos agora)")
    return df


# ------------------------------ RE0501 -----------------------------------
def ler_txt_tabela(caminho):
    with open(caminho, "rb") as f:
        bruto = f.read()
    for enc in ("utf-8-sig", "latin-1"):
        try:
            texto = bruto.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    linhas = [l for l in texto.splitlines() if l.strip()]
    if not linhas:
        return None
    amostra = "\n".join(linhas[:50])
    sep = max([";", "|", "\t", ","], key=amostra.count)
    tab = [[c.strip() for c in l.split(sep)] for l in linhas]
    # acha cabeçalho
    for i, lin in enumerate(tab[:40]):
        m = achar_colunas(lin, CONFIG["COLS_RE0501"])
        if "item" in m and "quantidade" in m:
            cab = [c if c else f"col{j}" for j, c in enumerate(lin)]
            corpo = [r + [""] * (len(cab) - len(r)) for r in tab[i + 1:]]
            df = pd.DataFrame([r[:len(cab)] for r in corpo], columns=cab)
            break
    else:
        return None
    m = achar_colunas(df.columns, CONFIG["COLS_RE0501"])
    if "chave" in m:
        ch = df[m["chave"]].astype(str).str.extract(r"(\d{44})")[0]
    else:  # chave em qualquer coluna da linha
        ch = df.astype(str).agg(" ".join, axis=1).str.extract(r"(\d{44})")[0]
    # relatórios em bloco: a chave aparece numa linha e os itens abaixo
    linhas_txt = df.astype(str).agg(" ".join, axis=1)
    ch = ch.fillna(linhas_txt.str.extract(r"(\d{44})")[0]).ffill()
    out = pd.DataFrame({
        "Chave_NF": ch,
        "Cod_Item": df[m["item"]],
        "Desc_Item": df[m["descricao"]] if "descricao" in m else "",
        "Qtd_Item": pd.to_numeric(df[m["quantidade"]].str.replace(".", "", regex=False)
                                    .str.replace(",", ".", regex=False), errors="coerce"),
        "Un_Item": df[m["unidade"]] if "unidade" in m else "",
        "Valor_Item": pd.to_numeric(df[m["valor_item"]].str.replace(".", "", regex=False)
                                    .str.replace(",", ".", regex=False), errors="coerce") if "valor_item" in m else pd.NA,
    })
    out = out[out["Chave_NF"].notna() & out["Cod_Item"].astype(str).str.strip().ne("") & out["Qtd_Item"].notna()]
    out["Arquivo_RE0501"] = os.path.basename(caminho)
    return out


def classificar(cod, desc):
    alvo = f"{cod} {desc}".upper()
    for cat, rx in CONFIG["REGRAS_CATEGORIA"]:
        if re.search(rx, alvo):
            return cat
    return CONFIG["CATEGORIA_PADRAO"]


def ler_re0501(pasta):
    frames, sem_layout = [], []
    for raiz, _, arqs in os.walk(pasta):
        for a in arqs:
            if not a.lower().endswith((".txt", ".csv")):
                continue
            try:
                df = ler_txt_tabela(os.path.join(raiz, a))
            except Exception as e:
                log(f"  ! erro no RE0501 {a}: {e}")
                df = None
            if df is None:
                sem_layout.append(a)
            elif not df.empty:
                frames.append(df)
    if sem_layout:
        log(f"  ! {len(sem_layout)} arquivo(s) RE0501 sem cabeçalho reconhecido (ex.: {sem_layout[:3]}). "
            "Ajuste CONFIG['COLS_RE0501'].")
    if not frames:
        raise SystemExit("Nenhum item lido do RE0501 – ajuste CONFIG['COLS_RE0501'].")
    df = pd.concat(frames, ignore_index=True).drop_duplicates()
    df["Categoria"] = [classificar(c, d) for c, d in zip(df["Cod_Item"], df["Desc_Item"])]
    log(f"RE0501: {len(df)} linhas de item | {df['Chave_NF'].nunique()} NF-e")
    return df, sem_layout


# ------------------------------- VÍNCULO ---------------------------------
def filial_bate(filial, cte):
    cnpjs = [cte.get(k, "") for k in ("CNPJ_Tomador", "CNPJ_Remetente", "CNPJ_Destinatario", "CNPJ_Emitente")]
    alvo = CONFIG["MAPA_FILIAL_CNPJ"].get(filial)
    if alvo:
        return re.sub(r"\D", "", alvo) in cnpjs
    f = norm_num(filial)
    return any(len(c) == 14 and norm_num(c[8:12]) == f for c in cnpjs)


def vincular(real, ctes, itens):
    por_doc = {}
    for r in ctes.to_dict("records"):
        por_doc.setdefault((r["Num_CTe"], r["Serie_CTe"]), []).append(r)

    vinc = []
    for id_linha, doc, serie, filial in real[["ID_Linha", "_doc", "_serie", "_filial"]].itertuples(index=False, name=None):
        cands = por_doc.get((doc, serie), [])
        status = "OK"
        if len(cands) > 1:
            filtr = [c for c in cands if filial_bate(filial, c)]
            if len(filtr) >= 1:
                cands = filtr[:1]
                status = "OK (desempate por filial)"
            else:
                cands = cands[:1]
                status = "ATENÇÃO: CT-e duplicado, filial não confirmada"
        if not cands:
            vinc.append({"ID_Linha": id_linha, "Status_Vinculo": "CT-e não encontrado no NEOGRID"})
            continue
        c = cands[0]
        nfs = c["Chaves_NF_Origem"] or [None]
        for nf in nfs:
            vinc.append({"ID_Linha": id_linha, "Status_Vinculo": status if nf else "CT-e sem NF-e de origem",
                         **{k: c[k] for k in ("Chave_CTe", "Emissao_CTe", "CNPJ_Emitente", "CNPJ_Tomador", "Valor_CTe")},
                         "Chave_NF_Origem": nf})
    vinc = pd.DataFrame(vinc)

    det = vinc.merge(itens, left_on="Chave_NF_Origem", right_on="Chave_NF", how="left")
    sem_item = det["Chave_NF_Origem"].notna() & det["Cod_Item"].isna()
    det.loc[sem_item, "Status_Vinculo"] = "NF-e não encontrada no RE0501"
    det = det.drop(columns=["Chave_NF"])
    det["Categoria"] = det["Categoria"].fillna("Sem vínculo")

    # rateio do valor realizado por quantidade dentro de cada linha do realizado
    det = det.merge(real[["ID_Linha", "_valor"]], on="ID_Linha", how="left")
    qtd_tot = det.groupby("ID_Linha")["Qtd_Item"].transform("sum")
    det["Valor_Rateado"] = (det["_valor"] * det["Qtd_Item"] / qtd_tot).where(qtd_tot > 0)
    det = det.drop(columns=["_valor"])

    cols_real = [c for c in real.columns if not c.startswith("_")] + ["_arquivo"]
    final = real[cols_real].rename(columns={"_arquivo": "Arquivo_Realizado"}).merge(det, on="ID_Linha", how="left")
    return final


# -------------------------------- SAÍDA ----------------------------------
def gravar(final, ctes, itens, sem_layout, saida):
    resumo = (final.groupby(["Categoria"], dropna=False)
              .agg(Linhas=("ID_Linha", "nunique"), Quantidade=("Qtd_Item", "sum"),
                   Valor_Rateado=("Valor_Rateado", "sum")).reset_index())
    resumo_item = (final.dropna(subset=["Cod_Item"])
                   .groupby(["Categoria", "Cod_Item", "Desc_Item"], dropna=False)
                   .agg(Quantidade=("Qtd_Item", "sum"), Valor_Rateado=("Valor_Rateado", "sum"),
                        Documentos=("ID_Linha", "nunique")).reset_index()
                   .sort_values("Valor_Rateado", ascending=False))
    status = final.groupby("Status_Vinculo", dropna=False)["ID_Linha"].nunique().reset_index(name="Linhas_Realizado")
    pend = final[~final["Status_Vinculo"].astype(str).str.startswith("OK")]
    cte_tab = ctes.assign(Chaves_NF_Origem=ctes["Chaves_NF_Origem"].map(" | ".join))
    params = pd.DataFrame({
        "Parâmetro": ["Gerado em", "Raiz", "Arquivos RE0501 sem layout reconhecido"],
        "Valor": [f"{datetime.now():%d/%m/%Y %H:%M}", CONFIG["RAIZ"], ", ".join(sem_layout) or "-"],
    })

    abas = {
        "Realizado_Itens": final, "Resumo_Categoria": resumo, "Resumo_Item": resumo_item,
        "Status_Vinculo": status, "Pendencias": pend, "Base_CTe": cte_tab,
        "Base_RE0501": itens, "Parametros": params,
    }
    with pd.ExcelWriter(saida, engine="openpyxl") as xw:
        for nome, df in abas.items():
            df.to_excel(xw, sheet_name=nome, index=False)
            ws = xw.sheets[nome]
            ws.freeze_panes = "A2"
            if ws.max_row > 1:
                ws.auto_filter.ref = ws.dimensions
            for col in ws.columns:
                larg = max(len(str(c.value)) if c.value is not None else 0 for c in list(col)[:500])
                ws.column_dimensions[col[0].column_letter].width = min(max(10, larg + 2), 60)
            for cel in ws[1]:
                cel.font = Font(bold=True)
    log(f"Arquivo gerado: {saida}")


def main():
    raiz = sys.argv[1] if len(sys.argv) > 1 else CONFIG["RAIZ"]
    CONFIG["RAIZ"] = raiz
    real = ler_realizado(os.path.join(raiz, CONFIG["PASTA_REALIZADO"]))
    ctes = ler_ctes(os.path.join(raiz, CONFIG["PASTA_CTE"]), os.path.join(raiz, "_cache_cte.pkl"))
    itens, sem_layout = ler_re0501(os.path.join(raiz, CONFIG["PASTA_RE0501"]))
    final = vincular(real, ctes, itens)
    saida = os.path.join(raiz, CONFIG["ARQUIVO_SAIDA"])
    try:
        gravar(final, ctes, itens, sem_layout, saida)
    except PermissionError:
        alt = saida.replace(".xlsx", f"_{datetime.now():%Y%m%d_%H%M}.xlsx")
        log("Arquivo de saída está aberto no Excel; gravando cópia.")
        gravar(final, ctes, itens, sem_layout, alt)


if __name__ == "__main__":
    main()
