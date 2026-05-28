"""
╔══════════════════════════════════════════════════════════════════╗
║   Análise de Evasão Free Flow — Aplicação Web (Streamlit)       ║
║   Execução: streamlit run app_evasao.py                         ║
╚══════════════════════════════════════════════════════════════════╝

Dependências:
    pip install streamlit pandas numpy openpyxl
"""

import io
import re

import importlib
import numpy as np
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import appextemporaneo # Importa o módulo inteiro
from appextemporaneo import renderizar_validacao_extemporaneo
# ═══════════════════════════════════════════════════════════════════
# CONFIGURAÇÕES DE NEGÓCIO  (não alterar sem alinhamento com o DER)
# ═══════════════════════════════════════════════════════════════════

MOTIVOS_FRAUDE = {
    "41 - Veículo não encontrado no RENAVAM",
    "40 - Placa",
    "37 - Placa estrangeira",
    "38 - Placa estrangeira",
    "41 - Placa",
    "42 - Marca/modelo ou tipo",
    "42 - Veículo não encontrado no RENAVAM",
}

SITUACAO_VALIDO = "Concluído / Registro Válido"

PADRAO_MERCOSUL = re.compile(r"^[A-Z]{3}\d[A-Z]\d{2}$")
PADRAO_ANTIGO   = re.compile(r"^[A-Z]{3}\d{4}$")

# ═══════════════════════════════════════════════════════════════════
# ESTILOS OPENPYXL
# ═══════════════════════════════════════════════════════════════════

def _fill(cor: str) -> PatternFill:
    return PatternFill("solid", start_color=cor, end_color=cor)


def _font(bold: bool = False, color: str = "000000", size: int = 10) -> Font:
    return Font(bold=bold, color=color, name="Arial", size=size)


THIN = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin"),
)

HEADER_CONFIGS = {
    "DADOS PASSAGENS CONCESSAO": ("375623", "FFFFFF"),
    "DADOS PASSAGENS DER":       ("375623", "FFFFFF"),
    "Valor Encontrado":          ("1F4E79", "FFFFFF"),
    "Valor Válido":              ("375623", "FFFFFF"),
    "Valor Fraude":              ("833C00", "FFFFFF"),
}


def auto_width(ws, extra: int = 4, max_w: int = 55) -> None:
    for col in ws.columns:
        letra = get_column_letter(col[0].column)
        larg  = max((len(str(c.value or "")) for c in col), default=8)
        ws.column_dimensions[letra].width = min(larg + extra, max_w)


def write_df_to_sheet(ws, df: pd.DataFrame, start_row: int = 1) -> None:
    """Escreve DataFrame com cabeçalho azul e linhas zebradas."""
    hdr_fill = _fill("1F4E79")
    hdr_font = _font(bold=True, color="FFFFFF")
    alt_fill = _fill("D6E4F0")

    for c, col in enumerate(df.columns, 1):
        cell = ws.cell(row=start_row, column=c, value=col)
        cell.fill      = hdr_fill
        cell.font      = hdr_font
        cell.border    = THIN
        cell.alignment = Alignment(horizontal="center", wrap_text=True)

    for r, row in enumerate(df.itertuples(index=False), start_row + 1):
        fll = alt_fill if r % 2 == 0 else PatternFill()
        for c, val in enumerate(row, 1):
            v    = None if (not isinstance(val, str) and pd.isna(val)) else val
            cell = ws.cell(row=r, column=c, value=v)
            cell.fill   = fll
            cell.font   = _font()
            cell.border = THIN


# ═══════════════════════════════════════════════════════════════════
# LEITURA DOS DADOS
# ═══════════════════════════════════════════════════════════════════
def get_col(df, *nomes):
    for nome in nomes:
        for col in df.columns:
            if col.strip().lower() == nome.lower():
                return col
    raise KeyError(f"Nenhuma das colunas {nomes} encontrada. Disponíveis: {list(df.columns)}")

def ler_tamoios(path_original) -> pd.DataFrame:
    """
    Lê a primeira aba VISÍVEL — Concessionária.
    O cabeçalho real está na linha 5 do Excel (header=4 no pandas).
    Abas ocultas (sheetState != 'visible') são ignoradas para evitar
    leitura de planilhas de suporte ou auxiliares escondidas.
    Aceita caminho em disco ou objeto BytesIO.
    """
    # Identifica a primeira aba visível via openpyxl antes de ler com pandas
    if isinstance(path_original, io.BytesIO):
        path_original.seek(0)
    _wb_probe = load_workbook(path_original, read_only=True, data_only=True)
    _nome_aba_visivel = next(
        (ws.title for ws in _wb_probe.worksheets if ws.sheet_state == "visible"),
        _wb_probe.sheetnames[0],   # fallback: usa a primeira se todas estiverem ocultas
    )
    _wb_probe.close()

    if isinstance(path_original, io.BytesIO):
        path_original.seek(0)
    df = pd.read_excel(path_original, sheet_name=0, header=4)

    # Data vem como string "dd/mm/yyyy hh:mm:ss"
    df.columns = df.columns.str.strip()

    col_data  = get_col(df, "Data", "Data da Passagem", "Data Passagem", "Data e Hora da Passagem")
    col_hora  = get_col(df, "Hora")
    col_placa = get_col(df, "Placa")
    col_valor = get_col(df, "Valor")

    df["Data"] = pd.to_datetime(df[col_data], dayfirst=True, errors="coerce")

    df["Hora_str"] = pd.to_datetime(
        df[col_hora].astype(str),
        errors="coerce"
    ).dt.strftime("%H:%M:%S")

    df["Placa"] = df[col_placa].astype(str).str.strip().str.upper()
    df["Placa7"] = df["Placa"].str[:7]
    # Chave de cruzamento: dd/mm/aa hh:mm:ss placa
    df["DADOS PASSAGENS CONCESSAO"] = (
        df["Data"].dt.strftime("%d/%m/%y")
        + " "
        + df["Hora_str"]
        + " "
        + df["Placa7"]
    )

    df["Valor"] = pd.to_numeric(df[col_valor], errors="coerce")
    return df


def ler_der(path_original) -> pd.DataFrame:
    """
    Lê a Aba 2 (índice 1) — DER.
    Aceita caminho em disco ou objeto BytesIO.
    """
    df = pd.read_excel(path_original, sheet_name=0)
    df.columns = df.columns.str.strip()

    # ─────────────────────────────────────────────
    # BUSCA FLEXÍVEL DA COLUNA "DATA"
    # ─────────────────────────────────────────────
    col_data = get_col(
        df, 
        "Data", 
        "Data da Passagem", 
        "Data Passagem", 
        "Data e Hora da Passagem"
    )

    df["Data"] = pd.to_datetime(
        df[col_data], dayfirst=True, errors="coerce"
    )
    
    df["Placa7"] = df["Veículo"].astype(str).str.strip().str.upper().str[:7]

    # Chave de cruzamento: dd/mm/aa hh:mm:ss placa
    df["DADOS PASSAGENS DER"] = (
        df["Data"].dt.strftime("%d/%m/%y")
        + " "
        + df["Data"].dt.strftime("%H:%M:%S")
        + " "
        + df["Placa7"]
    )

    # ─────────────────────────────────────────────
    # PADRONIZAÇÃO DAS COLUNAS
    # ─────────────────────────────────────────────
    df.columns = (
        df.columns
        .astype(str)
        .str.replace("\n", " ", regex=False)
        .str.replace("\r", " ", regex=False)
        .str.replace("\xa0", " ", regex=False)
        .str.strip()
    )

    # ─────────────────────────────────────────────
    # COLUNAS DE MOTIVO
    # ─────────────────────────────────────────────
    cols_motivo = [
        c for c in df.columns
        if "motivo invalidação" in c.lower() or "motivo da invalidação" in c.lower()
    ]
    
    # Tratativa de segurança para evitar erro de índice
    if not cols_motivo:
        raise KeyError(f"Coluna de motivo de invalidação não encontrada. Colunas disponíveis: {list(df.columns)}")

    # Normaliza todas
    for c in cols_motivo:
        df[c] = (
            df[c]
            .astype(str)
            .str.strip()
        )

    # ─────────────────────────────────────────────
    # BUSCA FLEXÍVEL DA COLUNA "SITUAÇÃO / FASE ANÁLISE"
    # ─────────────────────────────────────────────
    col_situacao = get_col(
        df,
        "Situação / Fase Análise",
        "Situacao / Fase Analise",
        "Situação/Fase Análise",
        "Situação / Fase"
    )

    df["Situação / Fase Análise"] = (
        df[col_situacao]
        .astype(str)
        .str.strip()
    )

    # ─────────────────────────────────────────────
    # MOTIVO CONSOLIDADO DE FRAUDE
    # ─────────────────────────────────────────────
    df["Motivo Invalidação Fraude"] = df[cols_motivo].apply(
        lambda row: next(
            (
                v for v in row
                if v in MOTIVOS_FRAUDE
            ),
            np.nan
        ),
        axis=1
    )

    # ─────────────────────────────────────────────
    # FLAG FRAUDE
    # ─────────────────────────────────────────────
    df["_fraude"] = (
        df["Motivo Invalidação Fraude"]
        .notna()
    )

    # ─────────────────────────────────────────────
    # MANTÉM COLUNA PRINCIPAL VISUAL
    # ─────────────────────────────────────────────
    df["Motivo Invalidação"] = df[cols_motivo[0]]
    
    return df


# ═══════════════════════════════════════════════════════════════════
# CRUZAMENTO E CÁLCULOS
# ═══════════════════════════════════════════════════════════════════

def cruzar_dados(df_der: pd.DataFrame, df_tam: pd.DataFrame) -> pd.DataFrame:
    """Matching DER × Concessionária e cálculo das colunas de valor."""

    # ── Mapeamento de valor ───────────────────────────────────────
    mapa_valor = (
        df_tam.drop_duplicates("DADOS PASSAGENS CONCESSAO")
        .set_index("DADOS PASSAGENS CONCESSAO")["Valor"]
    )

    df_der["Valor Encontrado"] = df_der["DADOS PASSAGENS DER"].map(mapa_valor)

    # ── NOVA LÓGICA: último status encontrado na linha vence ─────
    def linha_valida(row):
        ultimo_status = None

        for v in row:
            texto = str(v).strip().lower()

            tem_invalido = (
                "inválido" in texto or
                "invalido" in texto
            )

            tem_valido = (
                "válido" in texto or
                "valido" in texto
            )

            # prioridade pela posição na linha
            # o último encontrado sobrescreve o anterior
            if tem_invalido:
                ultimo_status = False

            elif tem_valido:
                ultimo_status = True

        return ultimo_status is True

    # ── FLAG VÁLIDO ──────────────────────────────────────────────
    df_der["_valido"] = df_der.apply(linha_valida, axis=1)

    # ── Valor Válido ─────────────────────────────────────────────
    df_der["Valor Válido"] = np.where(
        df_der["_valido"],
        df_der["Valor Encontrado"].fillna(0),
        0,
    )

    # ── FLAG FRAUDE ──────────────────────────────────────────────
    df_der["_fraude"] = (
        df_der["Motivo Invalidação Fraude"]
        .isin(MOTIVOS_FRAUDE)
    )

    # ── Valor Fraude ─────────────────────────────────────────────
    df_der["Valor Fraude"] = np.where(
        (~df_der["_valido"]) & df_der["_fraude"],
        df_der["Valor Encontrado"].fillna(0),
        0,
    )

    return df_der


# ═══════════════════════════════════════════════════════════════════
# RELATÓRIOS ANALÍTICOS
# ═══════════════════════════════════════════════════════════════════

def gerar_inconsistencias(df_der: pd.DataFrame, df_tam: pd.DataFrame) -> pd.DataFrame:
    """Registros com Valor Válido = 0 e Valor Fraude = 0, com justificativa."""
    mask   = (df_der["Valor Válido"] == 0) & (df_der["Valor Fraude"] == 0)
    df_inc = df_der[mask].copy()
    tam_set = set(df_tam["DADOS PASSAGENS CONCESSAO"])

    def justificativa(row):
        dado = row["DADOS PASSAGENS DER"]
        sit  = row["Situação / Fase Análise"]
        mot  = row["Motivo Invalidação"]
        if dado not in tam_set:
            return "DADOS PASSAGENS DER não encontrado em DADOS PASSAGENS CONCESSAO"
        if sit == SITUACAO_VALIDO and pd.isna(row["Valor Encontrado"]):
            return "Registro válido sem valor correspondente na Concessionária"
        if mot in ("nan", "", "NaN"):
            return "Situação inválida sem motivo de invalidação registrado"
        return f"Motivo '{mot}' não classificado como fraude nem como válido"

    df_inc["Justificativa"] = df_inc.apply(justificativa, axis=1)
    cols = [
        "DADOS PASSAGENS DER", "Data", "Veículo",
        "Situação / Fase Análise", "Motivo Invalidação",
        "Valor Encontrado", "Valor Válido", "Valor Fraude", "Justificativa",
    ]
    return df_inc[cols].reset_index(drop=True)


def detectar_erros_concessao(df_tam: pd.DataFrame) -> pd.DataFrame:
    """
    Detecta erros na Aba 1 (Concessionária) com duplicatas resumidas.
    """
    erros     = []
    base_cols = ["DADOS PASSAGENS CONCESSAO", "Placa", "Data", "Hora_str", "Valor"]

    # ── Etapa 1: linhas com valores ausentes ─────────────────────────────────
    campos_obrigatorios = {
        "Data":  df_tam["Data"].isna(),
        "Hora":  df_tam["Hora_str"].isin(["", "nan", "NaT"]) | df_tam["Hora_str"].isna(),
        "Placa": df_tam["Placa7"].isin(["", "NAN", "NAT"])   | df_tam["Placa7"].isna(),
        "Valor": df_tam["Valor"].isna(),
    }

    mask_incompleto = df_tam["Data"].isna()
    for m in campos_obrigatorios.values():
        mask_incompleto = mask_incompleto | m

    df_incompleto = df_tam[mask_incompleto].copy()
    df_completo   = df_tam[~mask_incompleto].copy()

    if not df_incompleto.empty:
        def campos_faltando(row):
            faltam = [
                campo
                for campo, mask in campos_obrigatorios.items()
                if mask.loc[row.name]
            ]
            return "Dados insuficientes — campo(s) ausente(s): " + ", ".join(faltam)

        df_incompleto["Tipo Erro"] = df_incompleto.apply(campos_faltando, axis=1)
        erros.append(df_incompleto[base_cols + ["Tipo Erro"]])

    # ── Etapa 2a: duplicatas (RESUMIDAS) ─────────────────────────────────────
    dup = (
        df_completo[df_completo.duplicated("DADOS PASSAGENS CONCESSAO", keep=False)]
        .groupby("DADOS PASSAGENS CONCESSAO")
        .agg({
            "Placa": "first",
            "Data": "first",
            "Hora_str": "first",
            "Valor": "first",
            "DADOS PASSAGENS CONCESSAO": "count"
        })
        .rename(columns={"DADOS PASSAGENS CONCESSAO": "Quantidade"})
        .reset_index()
    )

    if not dup.empty:
        dup["Tipo Erro"] = dup["Quantidade"].apply(
            lambda x: f"Passagem duplicada na Concessionária / {x} vezes"
        )

        dup = dup[
            ["DADOS PASSAGENS CONCESSAO", "Placa", "Data", "Hora_str", "Valor", "Tipo Erro"]
        ]

        erros.append(dup)

    # ── Etapa 2b: placas "FRAUDE" ────────────────────────────────────────────
    frd = df_completo[df_completo["Placa7"].str.upper() == "FRAUDE"].copy()
    if not frd.empty:
        frd["Tipo Erro"] = "Placa registrada como FRAUDE pela Concessionária"
        erros.append(frd[base_cols + ["Tipo Erro"]])

    # ── Etapa 2c: formato inválido ───────────────────────────────────────────
    def invalido(p):
        p = str(p).strip().upper()
        return p not in ("FRAUDE", "NAN", "") and not (
            PADRAO_MERCOSUL.match(p) or PADRAO_ANTIGO.match(p)
        )

    inv = df_completo[df_completo["Placa7"].apply(invalido)].copy()
    if not inv.empty:
        inv["Tipo Erro"] = "Formato de placa inválido (não Mercosul nem padrão antigo)"
        erros.append(inv[base_cols + ["Tipo Erro"]])

    # ── Resultado final ──────────────────────────────────────────────────────
    if erros:
        result = pd.concat(erros).drop_duplicates().reset_index(drop=True)
    else:
        result = pd.DataFrame(columns=base_cols + ["Tipo Erro"])

    result.columns = ["Chave Passagem", "Placa", "Data", "Hora", "Valor", "Tipo Erro"]

    return result

def detectar_passagens_proximas(df_tam: pd.DataFrame) -> pd.DataFrame:
    """
    Detecta passagens da mesma placa com intervalo menor ou igual a 60 segundos.
    Validação puramente informativa/auditável.
    """
    df = df_tam.copy()
    
    # 1. Cria uma coluna de Data/Hora consolidada para cálculo matemático
    df["_DataHora"] = pd.to_datetime(
        df["Data"].dt.strftime("%Y-%m-%d") + " " + df["Hora_str"],
        errors="coerce"
    )
    
    # Remove linhas sem data/hora ou sem placa para evitar erro no cálculo
    df = df.dropna(subset=["_DataHora", "Placa"])
    
    # 2. Ordena por Placa e cronologicamente
    df = df.sort_values(by=["Placa", "_DataHora"])
    
    # 3. Compara a linha atual com a linha imediatamente anterior
    df["Placa_prev"] = df["Placa"].shift(1)
    df["DataHora_prev"] = df["_DataHora"].shift(1)
    df["Valor_prev"] = df["Valor"].shift(1)
    
    # 4. Calcula a diferença em segundos
    df["Dif_Segundos"] = (df["_DataHora"] - df["DataHora_prev"]).dt.total_seconds()
    
    # 5. Filtra: Mesma placa E diferença <= 60 segundos (e >= 0 para garantir coerência)
    mask = (df["Placa"] == df["Placa_prev"]) & (df["Dif_Segundos"] <= 60) & (df["Dif_Segundos"] >= 0)
    df_proximas = df[mask].copy()
    
    # 6. Trata o caso de não encontrar nenhuma ocorrência
    if df_proximas.empty:
        return pd.DataFrame({"Mensagem": ["Nenhuma passagem próxima encontrada"]})
        
    # 7. Monta o DataFrame final com as colunas exatas solicitadas
    result = pd.DataFrame({
        "Placa": df_proximas["Placa"],
        "Primeira Passagem": df_proximas["DataHora_prev"].dt.strftime("%d/%m/%Y %H:%M:%S"),
        "Segunda Passagem Próxima": df_proximas["_DataHora"].dt.strftime("%d/%m/%Y %H:%M:%S"),
        "Diferença em segundos": df_proximas["Dif_Segundos"].astype(int),
        "Valor da primeira": df_proximas["Valor_prev"],
        "Valor da segunda": df_proximas["Valor"]
    })
    
    return result.reset_index(drop=True)

def detectar_erros_der(df_der: pd.DataFrame) -> pd.DataFrame:
    if df_der.empty:
        return pd.DataFrame()

    # considerar só chaves válidas
    df_base = df_der[
        df_der["DADOS PASSAGENS DER"].notna() &
        (df_der["DADOS PASSAGENS DER"] != "")
    ].copy()

    # AGRUPAMENTO IGUAL AO DA CONCESSIONÁRIA
    agrupado = (
        df_base
        .groupby("DADOS PASSAGENS DER")
        .agg({
            "Veículo": "first",
            "Data": "first",
            "DADOS PASSAGENS DER": "count"
        })
        .rename(columns={"DADOS PASSAGENS DER": "Quantidade"})
        .reset_index()
    )

    # FILTRA SÓ QUEM É DUPLICADO (>=2)
    duplicados = agrupado[agrupado["Quantidade"] > 1].copy()

    if duplicados.empty:
        return pd.DataFrame()

    # TEXTO IGUAL AO PADRÃO
    duplicados["Tipo Erro"] = duplicados["Quantidade"].apply(
        lambda x: f"Passagem duplicada no DER / {x} vezes"
    )

    duplicados = duplicados[
        ["DADOS PASSAGENS DER", "Veículo", "Data", "Tipo Erro"]
    ]

    duplicados.columns = ["Chave Passagem", "Placa", "Data", "Tipo Erro"]

    return duplicados.reset_index(drop=True)

def detectar_divergencias(df_der: pd.DataFrame, df_tam: pd.DataFrame) -> pd.DataFrame:
    """Registros presentes em apenas uma das fontes."""
    der_set = set(df_der["DADOS PASSAGENS DER"])
    tam_set = set(df_tam["DADOS PASSAGENS CONCESSAO"])

    so_der = df_der[~df_der["DADOS PASSAGENS DER"].isin(tam_set)][
        ["DADOS PASSAGENS DER", "Data", "Veículo",
         "Situação / Fase Análise", "Motivo Invalidação"]
    ].copy()
    so_der.columns = ["Chave Passagem", "Data/Hora", "Placa/Veículo", "Situação", "Motivo"]
    so_der["Origem"] = "Apenas no DER"

    so_tam = df_tam[~df_tam["DADOS PASSAGENS CONCESSAO"].isin(der_set)][
        ["DADOS PASSAGENS CONCESSAO", "Data", "Placa"]
    ].copy()
    so_tam.columns   = ["Chave Passagem", "Data/Hora", "Placa/Veículo"]
    so_tam["Situação"] = ""
    so_tam["Motivo"]   = ""
    so_tam["Origem"]   = "Apenas na Concessionária"

    return pd.concat([so_der, so_tam], ignore_index=True)


def consolidacao(df_der: pd.DataFrame) -> pd.DataFrame:
    total_val  = df_der["Valor Válido"].sum()
    total_fra  = df_der["Valor Fraude"].sum()
    qtd_total  = len(df_der)

    # Contabiliza apenas registros com Valor Encontrado > 0
    qtd_val = int(
    (df_der["_valido"] & (df_der["Valor Encontrado"].fillna(0) > 0)).sum()
    )
    qtd_fraude = int(
        (df_der["_fraude"] & (df_der["Valor Encontrado"].fillna(0) > 0)).sum()
    )

    rows = [
        ["Valores das Evasões DER", ""],
        ["DER: Valor de inadimplência (válidos)",     total_val],
        ["DER: Valor de fraude (inválidos c/ motivo)", total_fra],
        ["DER: Valor total evasão (válidos + fraudes)", total_val + total_fra],
        # Linha "Valor total encontrado" removida conforme item 4
        ["", ""],
        ["Quantidade de Passagens de Evasões", ""],
        ["DER: Quantidade total de registros",               qtd_total],
        ["DER: Quantidade de inadimplência (válidos)",       qtd_val],
        ["DER: Quantidade de fraudes (inválidos c/ motivo)", qtd_fraude],
    ]
    return pd.DataFrame(rows, columns=["Descrição", "Valor"])


# ═══════════════════════════════════════════════════════════════════
# ESCRITA DO EXCEL → BytesIO
# ═══════════════════════════════════════════════════════════════════

def escrever_resultado_otimizado(
    df_der, df_tam, df_inc, df_erros, df_div, df_cons, df_erros_der=None, df_proximas=None
):
    import io
    buffer = io.BytesIO()

    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        df_der.to_excel(writer, sheet_name="DER Processado", index=False)
        df_tam.to_excel(writer, sheet_name="Concessionária", index=False)
        df_inc.to_excel(writer, sheet_name="Inconsistências", index=False)
        df_erros.to_excel(writer, sheet_name="Erros Concessionária", index=False)
        
        if df_erros_der is not None and not df_erros_der.empty:
            df_erros_der.to_excel(writer, sheet_name="Erros DER", index=False)
            
        df_div.to_excel(writer, sheet_name="Divergências", index=False)
        df_cons.to_excel(writer, sheet_name="Resultado", index=False)
        
        # 🔥 NOVA ABA INSERIDA AQUI
        if df_proximas is not None:
            df_proximas.to_excel(writer, sheet_name="Passagens Próximas", index=False)

    buffer.seek(0)
    return buffer


# ═══════════════════════════════════════════════════════════════════
# INTERFACE STREAMLIT
# ═══════════════════════════════════════════════════════════════════

def formatar_brl(valor: float) -> str:
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def render_metric(label: str, valor: float, cor: str = "#1F4E79") -> str:
    """Retorna HTML para um card de métrica."""
    return f"""
    <div style="
        background:{cor}18;
        border-left: 5px solid {cor};
        border-radius: 6px;
        padding: 14px 18px;
        margin-bottom: 8px;
    ">
        <div style="font-size:0.78rem;color:#555;font-weight:600;
                    text-transform:uppercase;letter-spacing:.05em">
            {label}
        </div>
        <div style="font-size:1.45rem;font-weight:700;color:{cor};margin-top:4px">
            {formatar_brl(valor)}
        </div>
    </div>
    """

def renderizar_aba_analise():    
    st.markdown('<div class="section-title">📂 Upload das Planilhas</div>', unsafe_allow_html=True)

    # Cria duas colunas para os uploads ficarem organizados
    col_up1, col_up2 = st.columns(2)

    with col_up1:
        file_concessao = st.file_uploader(
            label="Planilha Concessionária",
            type=["xlsx"],
            help="Selecione o arquivo enviado pela concessionária (Aba 1 será lida)."
        )

    with col_up2:
        file_der = st.file_uploader(
            label="Planilha DER",
            type=["xlsx"],
            help="Selecione o arquivo extraído do sistema do DER (Aba 1 será lida)."
        )

    # Interrompe a execução se um dos dois não for enviado
    if not file_concessao or not file_der:
        st.info("⬆️ Por favor, faça o upload de ambas as planilhas para começar.")
        return

    # Informações sobre os arquivos carregados
    col_info1, col_info2 = st.columns(2)
    with col_info1:
        st.success(f"✅ Concessionária: **{file_concessao.name}**")
    with col_info2:
        st.success(f"✅ DER: **{file_der.name}**")

    # ── Botão de processamento ────────────────────────────────────────────────
    st.markdown('<div class="section-title">⚙️ Processamento</div>', unsafe_allow_html=True)

    if not st.button("▶ Processar Planilha", type="primary", use_container_width=True):
        st.caption("Clique no botão acima para iniciar a análise.")
        st.stop()

    # ── Pipeline de análise ───────────────────────────────────────────────────
    progress = st.progress(0, text="Iniciando…")

    try:
        # 1. Leitura
        progress.progress(10, text="📂 Lendo planilha da Concessionária...")
        df_tam = ler_tamoios(file_concessao)

        progress.progress(25, text="📂 Lendo planilha do DER...")
        df_der = ler_der(file_der)

        # 2. Cruzamento
        progress.progress(40, text="🔗 Cruzando dados DER × Concessionária…")
        df_der = cruzar_dados(df_der, df_tam)

        # 3. Detectar duplicatas DER (RELATÓRIO)
        progress.progress(55, text="🔍 Detectando duplicatas no DER…")
        df_erros_der = detectar_erros_der(df_der)

        # ── DUPLICIDADES FINANCEIRAS (válidos com valor) ─────────────
        duplicados_validos = df_der[
            (
                df_der.duplicated("DADOS PASSAGENS DER", keep=False)
            )
            &
            (
                df_der["_valido"]
            )
            &
            (
                df_der["Valor Encontrado"].fillna(0) > 0
            )
        ].copy()

        df_der_unico = df_der.drop_duplicates("DADOS PASSAGENS DER", keep="first")

        # 4. Relatórios
        progress.progress(65, text="🔍 Gerando relatório de inconsistências…")
        df_inc = gerar_inconsistencias(df_der_unico, df_tam)

        progress.progress(75, text="🔍 Detectando erros da Concessionária…")
        df_erros = detectar_erros_concessao(df_tam)

        progress.progress(78, text="🔍 Procurando passagens próximas (<= 60s)…")
        df_proximas = detectar_passagens_proximas(df_tam)

        progress.progress(82, text="🔍 Identificando divergências…")
        df_div  = detectar_divergencias(df_der_unico, df_tam)

        # 🔥 CONSOLIDAÇÃO USANDO BASE SEM DUPLICATA
        df_cons = consolidacao(df_der_unico)

        # 5. Geração do Excel
        progress.progress(90, text="💾 Gerando planilha de saída…")
        buffer_saida = escrever_resultado_otimizado(
            df_der, df_tam, df_inc, df_erros, df_div, df_cons,
            df_erros_der=df_erros_der,
            df_proximas=df_proximas,
        )

        progress.progress(100, text="✅ Concluído!")

    except Exception as e:
        progress.empty()
        st.error(f"❌ Erro durante o processamento: {e}")
        st.exception(e)
        st.stop()

    # ── Resultados ────────────────────────────────────────────────────────────
    enc = df_der_unico["Valor Encontrado"].fillna(0).sum()
    val = df_der_unico["Valor Válido"].sum()
    fra = df_der_unico["Valor Fraude"].sum()

    so_der_n = (df_div["Origem"] == "Apenas no DER").sum()
    so_tam_n = (df_div["Origem"] == "Apenas na Concessionária").sum()

    st.markdown('<div class="section-title">📊 Resultados do Processamento</div>', unsafe_allow_html=True)

    # Métricas financeiras
    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown(render_metric("Valor Encontrado (cruzamento)", val+fra, "#1F4E79"), unsafe_allow_html=True)
    with col2:
        st.markdown(render_metric("Valor Válido (inadimplência)", val, "#375623"), unsafe_allow_html=True)
    with col3:
        st.markdown(render_metric("Valor Fraude", fra, "#833C00"), unsafe_allow_html=True)

    # Métricas operacionais
    # Quantidades filtradas: apenas registros com Valor Encontrado > 0
    qtd_val_dash = int(
        (df_der_unico["_valido"] & (df_der_unico["Valor Encontrado"].fillna(0) > 0)).sum()
    )
    qtd_fraude_dash = int(
        (df_der_unico["_fraude"] & (df_der_unico["Valor Encontrado"].fillna(0) > 0)).sum()
    )

    col4, col5, col6, col7, col8 = st.columns(5)
    col4.metric("Registros DER",            f"{len(df_der):,}")
    col5.metric("Registros Concessionária", f"{len(df_tam):,}")
    col6.metric("Inconsistências",          f"{len(df_inc):,}")
    col7.metric("Erros Concessionária",     f"{len(df_erros):,}")
    col8.metric("Duplicatas DER",           f"{len(df_erros_der):,}")

    col8, col9, col10, col11 = st.columns(4)
    col8.metric("Qtd. Inadimplência (válidos c/ valor)",       f"{qtd_val_dash:,}")
    col9.metric("Qtd. Fraudes (inválidos c/ motivo e valor)",  f"{qtd_fraude_dash:,}")
    col10.metric("Divergências — só no DER",                   f"{so_der_n:,}")
    col11.metric("Divergências — só na Concessionária",        f"{so_tam_n:,}")

    duplicados_financeiros = (
        duplicados_validos
        .groupby("DADOS PASSAGENS DER", as_index=False)
        .agg({
            "Valor Encontrado": "first"
        })
    )

    valor_removido_duplicidades = (
        duplicados_financeiros["Valor Encontrado"]
        .fillna(0)
        .sum()
    )

    qtd_duplicidades_financeiras = len(duplicados_financeiros)

    if valor_removido_duplicidades > 0:

        st.warning(
            f"""
            ⚠️ Foram identificadas duplicidades financeiras no DER.

            Para evitar cobrança duplicada, o sistema removeu automaticamente:
            
            • {qtd_duplicidades_financeiras} chave(s) duplicada(s)
            • Valor total removido: {formatar_brl(valor_removido_duplicidades)}

            Por isso o valor final de inadimplência pode ser menor que a soma original dos registros válidos.
            """
        )

        st.dataframe(
            duplicados_financeiros.sort_values("DADOS PASSAGENS DER"),
            use_container_width=True,
            height=250
        )    

    # ── Prévia das tabelas ────────────────────────────────────────────────────
    st.markdown('<div class="section-title">🔎 Prévia dos Dados</div>', unsafe_allow_html=True)

    qtd_proximas = len(df_proximas) if "Placa" in df_proximas.columns else 0

    tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
        f"Inconsistências ({len(df_inc):,})",
        f"Erros Concessionária ({len(df_erros):,})",
        f"Erros DER ({len(df_erros_der):,})",
        f"Divergências ({len(df_div):,})",
        "Consolidação",
        f"Auditoria: Próximas ({qtd_proximas})",
    ])
    MAX_PREVIEW = 500

    with tab1:
        st.caption(f"Mostrando até {MAX_PREVIEW} linhas de {len(df_inc):,}")
        st.dataframe(df_inc.head(MAX_PREVIEW), use_container_width=True, height=320)

    with tab2:
        st.caption(f"Mostrando até {MAX_PREVIEW} linhas de {len(df_erros):,}")
        st.dataframe(df_erros.head(MAX_PREVIEW), use_container_width=True, height=320)

    with tab3:
        if df_erros_der.empty:
            st.success("✅ Nenhuma passagem duplicada encontrada no DER.")
        else:
            st.caption(f"Mostrando até {MAX_PREVIEW} linhas de {len(df_erros_der):,}")
            st.dataframe(df_erros_der.head(MAX_PREVIEW), use_container_width=True, height=320)

    with tab4:
        st.caption(f"Mostrando até {MAX_PREVIEW} linhas de {len(df_div):,}")
        st.dataframe(df_div.head(MAX_PREVIEW), use_container_width=True, height=320)

    with tab5:
        st.dataframe(df_cons, use_container_width=True, hide_index=True)

    # Adicione a renderização da tab6 lá embaixo, junto das outras:
    with tab6:
        st.dataframe(df_proximas.head(MAX_PREVIEW), use_container_width=True, hide_index=True)

    # ── Download ──────────────────────────────────────────────────────────────
    st.markdown('<div class="section-title">⬇️ Download do Resultado</div>', unsafe_allow_html=True)

    nome_saida = file_concessao.name.replace(".xlsx", "_Processado.xlsx")

    st.download_button(
        label="⬇️ Baixar Planilha Processada (.xlsx)",
        data=buffer_saida,
        file_name=nome_saida,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

    st.success(
        f"✅ Processamento concluído! O arquivo **{nome_saida}** está pronto para download."
    )

def renderizar_aba_uniao():
    st.markdown('<div class="section-title">📂 União de Planilhas DER</div>', unsafe_allow_html=True)
    st.write("Selecione até 12 arquivos do DER para consolidar em um único arquivo.")

    arquivos_der = st.file_uploader(
        "Upload das planilhas DER", 
        type=["xlsx"], 
        accept_multiple_files=True,
        key="uniao_der"
    )

    if arquivos_der:
        st.info(f"📁 {len(arquivos_der)} arquivo(s) carregado(s)")

        if st.button("Gerar Planilha Única"):
            progress = st.progress(0)
            logs = st.empty()

            try:
                lista_dfs = []

                total = len(arquivos_der[:12])

                # ── Etapa 1: leitura ─────────────────────────────
                for i, arq in enumerate(arquivos_der[:12], start=1):
                    logs.info(f"📂 Lendo arquivo {i}/{total}: {arq.name}")

                    try:
                        df = pd.read_excel(arq, sheet_name=0)
                        lista_dfs.append(df)

                    except Exception as e:
                        logs.error(f"❌ Erro ao ler {arq.name}: {e}")
                        continue

                    progress.progress(int((i / total) * 40))

                if not lista_dfs:
                    st.error("❌ Nenhum arquivo válido foi lido.")
                    return

                # ── Etapa 2: união ───────────────────────────────
                logs.info("🔗 Unificando planilhas...")
                df_consolidado = pd.concat(lista_dfs, ignore_index=True)

                progress.progress(60)

                # ── Etapa 3: validação básica ────────────────────
                logs.info("🔍 Validando estrutura dos dados...")

                linhas = len(df_consolidado)
                colunas = len(df_consolidado.columns)

                st.write(f"📊 Total de linhas: {linhas:,}")
                st.write(f"📊 Total de colunas: {colunas}")

                progress.progress(75)

                # ── Etapa 4: geração do arquivo ──────────────────
                logs.info("💾 Gerando arquivo final...")

                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df_consolidado.to_excel(writer, index=False)

                progress.progress(100)
                logs.success("✅ Processo concluído com sucesso!")

                # ── Download ─────────────────────────────────────
                st.download_button(
                    "⬇️ Baixar Planilha Consolidada",
                    data=output.getvalue(),
                    file_name="DER_Unificado.xlsx"
                )

            except Exception as e:
                progress.empty()
                logs.error("❌ Erro geral no processo")
                st.exception(e)

def main() -> None:
    # ── Configuração da página ────────────────────────────────────────────────

    tab_freeflow, tab_ext = st.tabs([
        "📊 Evasão FreeFlow",
        "🕵️ Extemporâneo"
    ])

    # ── ABA FREEFLOW ─────────────────────────────
    with tab_freeflow:

        # 🔥 HEADER AQUI
        st.markdown(
            """
            <div class="main-header">
                <h1>🛣️ Validação de Evasão Free Flow</h1>
                <p>Cruzamento e consolidação de inadimplências válidas e fraudes especificas</p>
            </div>
            """,
            unsafe_allow_html=True,
        )

        # ── SUBABAS ─────────────────────────────
        sub_analise, sub_uniao = st.tabs([
            "🔗 Cruzar Informações",
            "📂 Unificar Base do DER"
        ])

        with sub_analise:
            renderizar_aba_analise()

        with sub_uniao:
            renderizar_aba_uniao()


    # ── ABA EXTEMPORÂNEO ─────────────────────────
    with tab_ext:
        renderizar_validacao_extemporaneo()

    # ── CSS customizado ───────────────────────────────────────────────────────
    st.markdown(
        """
        <style>
        /* Cabeçalho principal */
        .main-header {
            background: linear-gradient(135deg, #1F4E79 0%, #2E75B6 100%);
            padding: 28px 32px;
            border-radius: 10px;
            margin-bottom: 28px;
            color: white;
        }
        .main-header h1 { margin: 0; font-size: 1.8rem; }
        .main-header p  { margin: 6px 0 0; opacity: .85; font-size: .95rem; }

        /* Separador de seção */
        .section-title {
            font-size: 1.05rem;
            font-weight: 700;
            color: #1F4E79;
            border-bottom: 2px solid #1F4E79;
            padding-bottom: 5px;
            margin: 24px 0 14px;
        }

        /* Badge de status */
        .badge-ok    { background:#C6EFCE; color:#375623;
                       padding:3px 10px; border-radius:12px; font-size:.8rem; font-weight:600; }
        .badge-warn  { background:#FFE699; color:#7F6000;
                       padding:3px 10px; border-radius:12px; font-size:.8rem; font-weight:600; }
        .badge-error { background:#FFD7D7; color:#9C0006;
                       padding:3px 10px; border-radius:12px; font-size:.8rem; font-weight:600; }

        /* Botão de download */
        div[data-testid="stDownloadButton"] button {
            background: #1F4E79 !important;
            color: white !important;
            font-weight: 600 !important;
            border-radius: 6px !important;
            padding: 10px 24px !important;
            border: none !important;
            width: 100%;
        }
        div[data-testid="stDownloadButton"] button:hover {
            background: #2E75B6 !important;
        }

        /* Upload widget */
        div[data-testid="stFileUploader"] {
            border: 2px dashed #2E75B6;
            border-radius: 8px;
            padding: 12px;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    # ── Cabeçalho ─────────────────────────────────────────────────────────────

    # ── Sidebar ───────────────────────────────────────────────────────────────
    with st.sidebar:
        st.image(
            "https://img.icons8.com/color/96/highway.png",
            width=72,
        )
        st.markdown("### ℹ️ Instruções de uso")
        st.markdown(
            """
            1. Faça o **upload** da planilha `.xlsx` original (sem fórmulas).
            2. Clique em **▶ Processar Planilha**.
            3. Aguarde o processamento.
            4. Baixe o arquivo processado com o botão **⬇️ Download**.

            ---
            **Estrutura esperada das planilhas:**
            - **Planilha 1** — Dados da Concessionária  
              *(cabeçalho na linha 5)*
            -  *Colunas: "Data", "Hora", "Placa", "Valor".*
            - *Obs: Conter Coluna: "Data do Pagamento"(SOMENTE NA DE EXTEMPORANEO)*
            - **Planilha 2** — Dados do DER
              *(cabeçalho na linha 1)*
            -  *Colunas: "Data Passagem", "Hora", "Placa"*
            ---
            **Motivos classificados como fraude:**
            """
        )
        for m in sorted(MOTIVOS_FRAUDE):
            st.markdown(f"- {m}")
        st.markdown("---")
        st.caption("Desenvolvido para análise de evasão de pedágio — Free Flow L27")



# ═══════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    main()