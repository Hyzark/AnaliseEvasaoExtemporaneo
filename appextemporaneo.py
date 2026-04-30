"""
╔══════════════════════════════════════════════════════════════════╗
║   Validação de Extemporâneo — Módulo complementar               ║
║   Integrar ao app_evasao.py                                     ║
╚══════════════════════════════════════════════════════════════════╝

Como integrar:
    1. Copie este arquivo para o mesmo diretório do app_evasao.py
    2. No app_evasao.py, adicione o import:
           from appextemporaneo import renderizar_validacao_extemporaneo
    3. No main(), adicione uma nova aba:
           tab_analise, tab_ext = st.tabs([
               "📊 Evasão FreeFlow", "🕵️ Extemporâneo"
           ])
           with tab_ext:
               renderizar_validacao_extemporaneo()
"""

import io

import numpy as np
import pandas as pd
import streamlit as st
from openpyxl.utils import get_column_letter


# ═══════════════════════════════════════════════════════════════════
# CONSTANTES
# ═══════════════════════════════════════════════════════════════════

JANELA_DIAS = 30  # Regra de negócio: pagamento extemporâneo deve ter ≥ 30 dias após passagem


# ═══════════════════════════════════════════════════════════════════
# UTILITÁRIOS INTERNOS DE VALIDAÇÃO
# ═══════════════════════════════════════════════════════════════════

def _linha_valida(row) -> bool:
    """
    Replica a lógica de linha_valida de cruzar_dados.
    Retorna True se qualquer célula da linha contiver 'válido'
    (sem 'inválido').
    """
    for v in row:
        texto = str(v).strip().lower()
        if "válido" in texto and "inválido" not in texto:
            return True
    return False


def _aplicar_validacao(df: pd.DataFrame) -> pd.DataFrame:
    """Acrescenta coluna _valido e filtra apenas linhas válidas."""
    df = df.copy()
    df["_valido"] = df.apply(_linha_valida, axis=1)
    return df[df["_valido"]].drop(columns=["_valido"]).reset_index(drop=True)


# ═══════════════════════════════════════════════════════════════════
# REGRA DE NEGÓCIO — VALIDADOR TEMPORAL (janela de 30 dias)
# ═══════════════════════════════════════════════════════════════════

def _calcular_conclusao(data_passagem, data_pagamento) -> str:
    """
    Compara Data da passagem com Data de pagamento.

    Regras:
      • pagamento ≥ 30 dias após passagem  → "OK"
      • pagamento < 30 dias após passagem  → "INVÁLIDO: PAGAMENTO ANTES DE GERAR MULTA"
      • pagamento anterior à passagem      → "INVÁLIDO: MULTA GERADA ANTES DA PASSAGEM"
      • datas ausentes                     → "SEM DATA"
    """
    if pd.isna(data_passagem) or pd.isna(data_pagamento):
        return "SEM DATA"

    delta = (data_pagamento - data_passagem).days

    if delta >= JANELA_DIAS:
        return "OK"
    elif delta >= 0:
        return "INVÁLIDO: PAGAMENTO ANTES DE GERAR MULTA"
    else:
        return "INVÁLIDO: MULTA GERADA ANTES DA PASSAGEM"


# ═══════════════════════════════════════════════════════════════════
# LEITURA DAS PLANILHAS (re-exporta de app_evasao para conveniência)
# ═══════════════════════════════════════════════════════════════════

def _ler_der(path_original) -> pd.DataFrame:
    from app_evasao import ler_der
    return ler_der(path_original)


def _ler_tamoios(path_original) -> pd.DataFrame:
    from app_evasao import ler_tamoios
    return ler_tamoios(path_original)

def calcular_passagens_por_mes(df, coluna_data="Data"):
    """
    Agrupa passagens por mês, considerando apenas matches (DER e Concessionária)
    e removendo duplicatas de registros.
    """
    df_temp = df.copy()

    # 1. Filtra apenas os registros que deram Match (Sim)
    # Isso garante que a passagem "bateu" com o DER
    if "DER x CONCESSIONARIA" in df_temp.columns:
        df_temp = df_temp[df_temp["DER x CONCESSIONARIA"] == "Sim"]

    # 2. Remove duplicatas (contabiliza apenas um registro por passagem)
    # Utilizamos a chave 'DADOS PASSAGENS CONCESSAO' para garantir a unicidade
    if "DADOS PASSAGENS CONCESSAO" in df_temp.columns:
        df_temp = df_temp.drop_duplicates("DADOS PASSAGENS CONCESSAO")

    # 3. Converter para datetime
    df_temp[coluna_data] = pd.to_datetime(
        df_temp[coluna_data],
        dayfirst=True,
        errors="coerce"
    )

    # Remover datas inválidas que possam ter sobrado
    df_temp = df_temp.dropna(subset=[coluna_data])

    if df_temp.empty:
        return pd.DataFrame(columns=["Mes/Ano", "Quantidade"])

    # 4. Agrupar por mês/ano
    agrupado = (
        df_temp
        .groupby(df_temp[coluna_data].dt.to_period("M"))
        .size()
        .reset_index(name="Quantidade")
    )

    # 5. Formatar para exibição (MM/YYYY)
    agrupado["Mes/Ano"] = agrupado[coluna_data].dt.strftime("%m/%Y")

    # Ordenar cronologicamente
    agrupado = agrupado.sort_values(by=coluna_data)

    return agrupado[["Mes/Ano", "Quantidade"]]
# ═══════════════════════════════════════════════════════════════════
# DETECÇÃO DE DUPLICATAS (mesma lógica do app_evasao)
# ═══════════════════════════════════════════════════════════════════

def _detectar_duplicatas_concessao(df_tam: pd.DataFrame) -> pd.DataFrame:
    """
    Duplicatas na planilha da Concessionária (mesma chave DADOS PASSAGENS CONCESSAO).
    Retorna resumo agrupado — mesmo padrão de detectar_erros_concessao.
    """
    df_completo = df_tam[df_tam["DADOS PASSAGENS CONCESSAO"].notna()].copy()

    dup = (
        df_completo[df_completo.duplicated("DADOS PASSAGENS CONCESSAO", keep=False)]
        .groupby("DADOS PASSAGENS CONCESSAO")
        .agg({"Placa": "first", "Data": "first", "Hora_str": "first",
              "Valor": "first", "DADOS PASSAGENS CONCESSAO": "count"})
        .rename(columns={"DADOS PASSAGENS CONCESSAO": "Quantidade"})
        .reset_index()
    )

    if dup.empty:
        return pd.DataFrame(columns=["Chave Passagem", "Placa", "Data", "Hora", "Valor", "Tipo Erro"])

    dup["Tipo Erro"] = dup["Quantidade"].apply(
        lambda x: f"Passagem duplicada na Consessionária / {x} vezes"
    )
    dup = dup[["DADOS PASSAGENS CONCESSAO", "Placa", "Data", "Hora_str", "Valor", "Tipo Erro"]]
    dup.columns = ["Chave Passagem", "Placa", "Data", "Hora", "Valor", "Tipo Erro"]
    return dup.reset_index(drop=True)


def _detectar_duplicatas_der(df_der: pd.DataFrame) -> pd.DataFrame:
    """
    Duplicatas no DER (mesma chave DADOS PASSAGENS DER).
    Retorna resumo agrupado — mesmo padrão de detectar_erros_der.
    """
    df_base = df_der[df_der["DADOS PASSAGENS DER"].notna()].copy()

    agrupado = (
        df_base.groupby("DADOS PASSAGENS DER")
        .agg({"Veículo": "first", "Data Passagem": "first",
              "DADOS PASSAGENS DER": "count"})
        .rename(columns={"DADOS PASSAGENS DER": "Quantidade"})
        .reset_index()
    )

    duplicados = agrupado[agrupado["Quantidade"] > 1].copy()

    if duplicados.empty:
        return pd.DataFrame(columns=["Chave Passagem", "Placa", "Data", "Tipo Erro"])

    duplicados["Tipo Erro"] = duplicados["Quantidade"].apply(
        lambda x: f"Passagem duplicada no DER / {x} vezes"
    )
    duplicados = duplicados[["DADOS PASSAGENS DER", "Veículo", "Data Passagem", "Tipo Erro"]]
    duplicados.columns = ["Chave Passagem", "Placa", "Data", "Tipo Erro"]
    return duplicados.reset_index(drop=True)


# ═══════════════════════════════════════════════════════════════════
# CRUZAMENTO PRINCIPAL — EXTEMPORÂNEO
# ═══════════════════════════════════════════════════════════════════

def _cruzar_extemporaneo_completo(
    df_der: pd.DataFrame,
    df_tam: pd.DataFrame,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """
    Realiza o cruzamento DER × Concessionária com:
      • Validação de presença (match)
      • Validação temporal (janela de 30 dias)
      • Cálculo financeiro apenas para as passagens válidas (OK)
    """

    # ── 1. Uma entrada por chave (sem duplicatas financeiras) ─────
    der_unico = df_der.drop_duplicates("DADOS PASSAGENS DER", keep="first").copy()
    tam_unico = df_tam.drop_duplicates("DADOS PASSAGENS CONCESSAO", keep="first").copy()

    # ── 2. Mapeamento: chave DER → valor e colunas da Concessionária
    mapa_valor = tam_unico.set_index("DADOS PASSAGENS CONCESSAO")["Valor"]
    chaves_tam = set(tam_unico["DADOS PASSAGENS CONCESSAO"])
    
    # 🔥 CORREÇÃO: Pegar apenas chaves reais (ignorar "nan" ou vazias)
    chaves_der = {
        k for k in der_unico["DADOS PASSAGENS DER"] 
        if pd.notna(k) and str(k).strip() != "" and "nan" not in str(k).lower()
    }

    # ── 3. Enriquecer Concessionária com colunas de validação ────
    df_tam_proc = df_tam.copy()

    if "Data do Pagamento" not in df_tam_proc.columns:
        df_tam_proc["Data do Pagamento"] = pd.NaT
    else:
        df_tam_proc["Data do Pagamento"] = pd.to_datetime(
            df_tam_proc["Data do Pagamento"], dayfirst=True, errors="coerce"
        )

    # 🔥 CORREÇÃO: Só dá "Sim" se a chave da concessionária também não for "nan" e existir no DER
    df_tam_proc["DER x CONCESSIONARIA"] = df_tam_proc["DADOS PASSAGENS CONCESSAO"].apply(
        lambda k: "Sim" if (
            pd.notna(k) and "nan" not in str(k).lower() and k in chaves_der
        ) else "Não"
    )

    # Conclusão temporal
    df_tam_proc["Conclusão"] = df_tam_proc.apply(
        lambda row: (
            _calcular_conclusao(row["Data"], row["Data do Pagamento"])
            if row["DER x CONCESSIONARIA"] == "Sim"
            else "Não Encontrado no DER"
        ),
        axis=1,
    )

    chaves_concessao_encontradas = df_tam_proc["DADOS PASSAGENS CONCESSAO"].apply(
        lambda k: k if (
            pd.notna(k) and "nan" not in str(k).lower() and k in chaves_der
        ) else np.nan
    )
    valores_repitidos = (
        chaves_concessao_encontradas.notna()
        & chaves_concessao_encontradas.duplicated(keep="first")
    )
    df_tam_proc.loc[valores_repitidos, "Conclusão"] = "VALOR REPITIDO"

    # ── 4. Ausentes ───────────────────────────────────────────────
    so_der = der_unico[~der_unico["DADOS PASSAGENS DER"].isin(chaves_tam)][
        ["DADOS PASSAGENS DER", "Data Passagem", "Veículo",
         "Situação / Fase Análise", "Motivo Invalidação"]
    ].copy()
    so_der.columns = ["Chave Passagem", "Data/Hora", "Placa/Veículo", "Situação", "Motivo"]
    so_der["Origem"] = "Apenas no DER"

    so_tam = tam_unico[~tam_unico["DADOS PASSAGENS CONCESSAO"].isin(chaves_der)][
        ["DADOS PASSAGENS CONCESSAO", "Data", "Placa"]
    ].copy()
    so_tam.columns = ["Chave Passagem", "Data/Hora", "Placa/Veículo"]
    so_tam["Situação"] = ""
    so_tam["Motivo"]   = ""
    so_tam["Origem"]   = "Apenas na Concessionária"

    df_ausentes = pd.concat([so_der, so_tam], ignore_index=True)

    # ── 5. Resultado executivo (Tabela Resumo) ────────────────────
    
    # Contagens organizadas
    n_der       = len(der_unico)
    n_tam       = len(tam_unico)
    n_match     = (df_tam_proc["DER x CONCESSIONARIA"] == "Sim").sum()
    n_ok        = (df_tam_proc["Conclusão"] == "OK").sum()
    n_inv_antes = (df_tam_proc["Conclusão"] == "INVÁLIDO: PAGAMENTO ANTES DE GERAR MULTA").sum()
    n_inv_multa = (df_tam_proc["Conclusão"] == "INVÁLIDO: MULTA GERADA ANTES DA PASSAGEM").sum()
    n_sem       = (df_tam_proc["Conclusão"] == "SEM DATA").sum()
    n_repitido  = (df_tam_proc["Conclusão"] == "VALOR REPITIDO").sum()

    # Cálculo do Financeiro SOMENTE para os que deram "OK"
    chaves_ok = set(df_tam_proc.loc[df_tam_proc["Conclusão"] == "OK", "DADOS PASSAGENS CONCESSAO"])

    def obter_valor_ok(row):
        chave = row["DADOS PASSAGENS DER"]
        if chave in chaves_ok:
            valor = mapa_valor.get(chave, 0)
            return valor if pd.notna(valor) else 0
        return 0

    der_unico["Valor Válido"] = der_unico.apply(obter_valor_ok, axis=1)
    total_valor_valido = der_unico["Valor Válido"].sum()

    # Montando a Tabela
    df_resultado = pd.DataFrame([
        ["Resumo Extemporâneo", ""],
        ["Registros DER",                int(n_der)],
        ["Registros Concessionária",     int(n_tam)],
        ["Registros DER x CONCESSIONARIA com (Sim)",             int(n_match)],
        ["", ""],
        ["Validação Temporal (Concessionária)", ""],
        ["Conclusão de Valores OK (≥ 30 dias)",              int(n_ok)],
        ["VALOR REPITIDO CONCESSIONÁRIA",                        int(n_repitido)],
        ["INVÁLIDO: Pagamento antes de gerar multa", int(n_inv_antes)],
        ["INVÁLIDO: Multa gerada antes da passagem", int(n_inv_multa)],
        ["SEM DATA (data ausente)",               int(n_sem)],
        ["", ""],
        ["Financeiro", ""],
        ["Soma total do Valor VÁLIDO (R$)",       total_valor_valido],
    ], columns=["Descrição", "Valor"])

    # ── 6. Duplicatas ─────────────────────────────────────────────
    dup_tam = _detectar_duplicatas_concessao(df_tam)
    dup_der = _detectar_duplicatas_der(df_der)
    df_duplicatas = pd.concat([dup_tam, dup_der], ignore_index=True)

    return df_tam_proc, df_ausentes, df_resultado, df_duplicatas


# ═══════════════════════════════════════════════════════════════════
# GERAÇÃO DO EXCEL (4 abas)
# ═══════════════════════════════════════════════════════════════════

def _gerar_excel_extemporaneo(
    df_der: pd.DataFrame,
    df_tam_proc: pd.DataFrame,
    df_ausentes: pd.DataFrame,
    df_resultado: pd.DataFrame,
    df_duplicatas: pd.DataFrame,
) -> bytes:

    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        # Aba 1
        df_ausentes.to_excel(writer, sheet_name="Ausentes", index=False)

        # Aba 2
        df_resultado.to_excel(writer, sheet_name="Resultado", index=False)

        # Aba 3
        df_tam_proc.to_excel(writer, sheet_name="Concessionária Processada", index=False)
        #aba 4
        df_der.to_excel(writer, sheet_name="DER", index=False)

        # Aba 5 — somente passagens válidas
        df_ok = df_tam_proc[df_tam_proc["Conclusão"] == "OK"]
        resumo_mes = calcular_passagens_por_mes(df_ok, "Data")
        resumo_mes.to_excel(writer, sheet_name="Passagens OK por Mês", index=False)

        # Aba 5
        if df_duplicatas.empty:
            pd.DataFrame({"Mensagem": ["Nenhuma duplicata encontrada"]}) \
                .to_excel(writer, sheet_name="Duplicadas", index=False)
        else:
            df_duplicatas.to_excel(writer, sheet_name="Duplicadas", index=False)

    output.seek(0)
    return output.getvalue()


# ═══════════════════════════════════════════════════════════════════
# ABA — CRUZAR INFORMAÇÕES (aba principal do extemporâneo)
# ═══════════════════════════════════════════════════════════════════

def _renderizar_cruzar_informacoes() -> None:
    st.markdown(
        '<div class="section-title">🔗 Cruzamento Extemporâneo DER × Concessionária</div>',
        unsafe_allow_html=True,
    )
    st.write(
        "Faça o upload das duas planilhas para identificar correspondências e validar "
        "a janela temporal de **30 dias** entre a passagem e o pagamento."
    )

    col_up1, col_up2 = st.columns(2)

    with col_up1:
        file_concessao = st.file_uploader(
            "Planilha Concessionária",
            type=["xlsx"],
            key="ext_cruzar_concessao",
            help="Arquivo enviado pela concessionária (cabeçalho na linha 5).",
        )

    with col_up2:
        file_der = st.file_uploader(
            "Planilha DER",
            type=["xlsx"],
            key="ext_cruzar_der",
            help="Arquivo extraído do sistema do DER.",
        )

    if not file_der or not file_concessao:
        st.info("⬆️ Faça o upload de ambas as planilhas para habilitar o cruzamento.")
        return

    col_i1, col_i2 = st.columns(2)
    with col_i1:
        st.success(f"✅ Concessionária: **{file_concessao.name}**")
    with col_i2:
        st.success(f"✅ DER: **{file_der.name}**")

    if not st.button("▶ Cruzar Informações", type="primary",
                     use_container_width=True, key="ext_btn_cruzar"):
        st.caption("Clique no botão acima para iniciar o cruzamento.")
        return

    progress = st.progress(0, text="Iniciando…")

    try:
        progress.progress(15, text="📂 Lendo planilha do DER...")
        df_der = _ler_der(file_der)

        progress.progress(30, text="📂 Lendo planilha da Concessionária...")
        df_tam = _ler_tamoios(file_concessao)

        progress.progress(50, text="🔗 Cruzando dados e aplicando validação temporal...")
        df_tam_proc, df_ausentes, df_resultado, df_duplicatas = _cruzar_extemporaneo_completo(
            df_der, df_tam
        )

        progress.progress(75, text="💾 Gerando arquivo de saída...")
        excel_bytes = _gerar_excel_extemporaneo(
            df_der, df_tam_proc, df_ausentes, df_resultado, df_duplicatas
        )
        progress.progress(100, text="✅ Concluído!")

    except Exception as e:
        progress.empty()
        st.error(f"❌ Erro durante o cruzamento: {e}")
        st.exception(e)
        return

    # ── Métricas ─────────────────────────────────────────────────
    st.markdown(
        '<div class="section-title">📊 Resultado do Cruzamento</div>',
        unsafe_allow_html=True,
    )

    # Recontagem correta
    total_conc   = len(df_tam)
    total_der    = len(df_der)
    total_match  = (df_tam_proc["DER x CONCESSIONARIA"] == "Sim").sum()
    total_nao    = (df_tam_proc["DER x CONCESSIONARIA"] == "Não").sum()
    n_ok         = (df_tam_proc["Conclusão"] == "OK").sum()
    n_inv_antes  = (df_tam_proc["Conclusão"] == "INVÁLIDO: PAGAMENTO ANTES DE GERAR MULTA").sum()
    n_inv_multa  = (df_tam_proc["Conclusão"] == "INVÁLIDO: MULTA GERADA ANTES DA PASSAGEM").sum()
    n_repitido   = (df_tam_proc["Conclusão"] == "VALOR REPITIDO").sum()

    col1, col2, col3, col4 = st.columns(4)
    # Agora sim mostra o total real do arquivo vs o que deu Match
    col1.metric("Total Concessionária",    f"{total_conc:,}")
    col2.metric("Total DER",               f"{total_der:,}")
    col3.metric("Match (Deu Sim)",         f"{total_match:,}")
    col4.metric("Ausentes (Deu Não)",      f"{total_nao:,}")

    st.markdown("**Resumo Temporal (Apenas os que deram Match):**")
    col5, col6, col7, col8 = st.columns(4)
    col5.metric("✅ Conclusão OK", f"{n_ok:,}")
    col6.metric("Valor Repitido", f"{n_repitido:,}")
    col7.metric("❌ Pagamento Antes da Multa", f"{n_inv_antes:,}")
    col8.metric("❌ Multa Antes da Passagem", f"{n_inv_multa:,}")
    # ── Prévia ───────────────────────────────────────────────────
    total_conc_proc = int(len(df_tam_proc))
    total_ausent    = int(len(df_ausentes))
    n_dup           = int(len(df_duplicatas))
    
    # Se MAX_PREVIEW não estiver definido no seu arquivo, defina aqui:
    if 'MAX_PREVIEW' not in locals():
        MAX_PREVIEW = 1000

    st.markdown(
        '<div class="section-title">🔎 Prévia dos Dados</div>',
        unsafe_allow_html=True,
    )

    # ── 2. Renderização das Abas ────────────────────────────────
    prev_conc, prev_aus, prev_res, prev_dup = st.tabs([
        f"Concessionária Processada ({total_conc_proc:,})",
        f"Ausentes ({total_ausent:,})",
        "Resultado",
        f"Duplicadas ({n_dup:,})",
    ])

    with prev_conc:
        st.caption(f"Mostrando até {MAX_PREVIEW} de {total_conc_proc:,} linhas")
        st.dataframe(df_tam_proc.head(MAX_PREVIEW), use_container_width=True, height=340)

    with prev_aus:
        st.caption(f"Mostrando até {MAX_PREVIEW} de {total_ausent:,} linhas")
        # .head() garante que o Streamlit não tente carregar 1 milhão de linhas de uma vez
        st.dataframe(df_ausentes.head(MAX_PREVIEW), use_container_width=True, height=340)

    with prev_res:
        # Mostra a tabela de resumo executivo
        st.dataframe(df_resultado, use_container_width=True, hide_index=True)

    with prev_dup:
        if n_dup == 0:
            st.success("✅ Nenhuma duplicata encontrada nas planilhas.")
        else:
            st.caption(f"Mostrando até {MAX_PREVIEW} de {n_dup:,} linhas")
            st.dataframe(df_duplicatas.head(MAX_PREVIEW), use_container_width=True, height=340)

    # ── Download ─────────────────────────────────────────────────
    st.markdown(
        '<div class="section-title">⬇️ Download do Resultado</div>',
        unsafe_allow_html=True,
    )
    nome_saida = file_concessao.name.replace(".xlsx", "Processado.xlsx")
    st.download_button(
        label="⬇️ Baixar Resultado Extemporâneo (.xlsx)",
        data=excel_bytes,
        file_name=nome_saida,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
    st.success(f"✅ Arquivo **{nome_saida}** pronto para download.")


# ═══════════════════════════════════════════════════════════════════
# ABA — UNIFICAR INADIMPLÊNCIAS (somente registros válidos)
# ═══════════════════════════════════════════════════════════════════

def _renderizar_unificar_inadimplencias() -> None:
    st.markdown(
        '<div class="section-title">📂 União de Planilhas DER — Apenas Válidos</div>',
        unsafe_allow_html=True,
    )
    st.write(
        "Selecione até 12 arquivos do DER. "
        "A planilha gerada conterá **somente** os registros considerados válidos."
    )

    arquivos_der = st.file_uploader(
        "Upload das planilhas DER",
        type=["xlsx"],
        accept_multiple_files=True,
        key="ext_uniao_der",
    )

    if not arquivos_der:
        return

    st.info(f"📁 {len(arquivos_der)} arquivo(s) carregado(s)")

    if not st.button("Gerar Planilha de Inadimplências Válidas", key="ext_btn_uniao"):
        return

    progress = st.progress(0)
    logs = st.empty()

    try:
        lista_dfs = []
        total = len(arquivos_der[:12])

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

        logs.info("🔗 Unificando planilhas...")
        df_consolidado = pd.concat(lista_dfs, ignore_index=True)
        progress.progress(55)

        logs.info("🔍 Aplicando filtro de registros válidos...")
        df_validos = _aplicar_validacao(df_consolidado)
        progress.progress(75)

        total_linhas   = len(df_consolidado)
        validas_linhas = len(df_validos)

        col1, col2, col3 = st.columns(3)
        col1.metric("Total de registros lidos", f"{total_linhas:,}")
        col2.metric("Registros válidos",         f"{validas_linhas:,}")
        col3.metric("Registros descartados",     f"{total_linhas - validas_linhas:,}")

        logs.info("💾 Gerando arquivo final...")
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            df_validos.to_excel(writer, index=False, sheet_name="Inadimplências Válidas")
        progress.progress(100)
        logs.success("✅ Processo concluído com sucesso!")

        st.download_button(
            "⬇️ Baixar Inadimplências Válidas (.xlsx)",
            data=output.getvalue(),
            file_name="DER_Inadimplencias_Validas.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    except Exception as e:
        progress.empty()
        logs.error("❌ Erro geral no processo")
        st.exception(e)


# ═══════════════════════════════════════════════════════════════════
# PONTO DE ENTRADA DA PÁGINA
# ═══════════════════════════════════════════════════════════════════

def renderizar_validacao_extemporaneo() -> None:
    """
    Página principal de Validação de Extemporâneo.
    Chame esta função dentro de uma aba no main() do app_evasao.py.
    """
    st.markdown(
        """
        <div class="main-header">
            <h1>🕵️ Validação de Extemporâneo</h1>
            <p>Cruzamento com janela temporal de 30 dias e consolidação de inadimplências válidas</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    aba_cruzar, aba_uniao = st.tabs([
        "🔗 Cruzar Informações",
        "📂 Unificar Inadimplências",
    ])

    with aba_cruzar:
        _renderizar_cruzar_informacoes()

    with aba_uniao:
        _renderizar_unificar_inadimplencias()
