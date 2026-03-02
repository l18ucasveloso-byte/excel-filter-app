import streamlit as st
import pandas as pd
import io
import re

# ==========================
# Configuração básica
# ==========================

st.set_page_config(
    page_title="Processador de Planilhas Excel",
    page_icon="📊",
    layout="wide"
)


# ==========================
# Funções auxiliares
# ==========================

def parse_keywords(raw_text: str):
    """
    Recebe uma string com palavras-chave separadas por vírgula
    e devolve uma lista limpa (sem espaços vazios).
    """
    if not raw_text:
        return []
    keywords = [k.strip() for k in raw_text.split(",")]
    # Remove vazios
    keywords = [k for k in keywords if k]
    return keywords


def build_primary_mask(df: pd.DataFrame, keywords):
    """
    Cria uma máscara booleana para linhas que contêm
    QUALQUER uma das palavras-chave principais em QUALQUER coluna.
    """
    if df.empty or not keywords:
        return pd.Series([False] * len(df), index=df.index)

    # Converte todo o DataFrame para string para busca case-insensitive
    df_str = df.astype(str)

    # Monta um padrão de regex com todas as palavras-chave, escapando caracteres especiais
    pattern = "|".join(re.escape(k) for k in keywords)

    # Aplica .str.contains em cada coluna e faz OR entre colunas
    mask = df_str.apply(
        lambda col: col.str.contains(pattern, case=False, na=False)
    ).any(axis=1)

    return mask


def classify_rows(df: pd.DataFrame):
    """
    Recebe um DataFrame (já filtrado pela busca principal)
    e devolve três DataFrames:
      - df_certificados: linhas contendo "certificado"
      - df_logbook: linhas contendo "logbook"
      - df_resto: demais linhas
    A busca é feita em toda a linha (todas as colunas), case-insensitive.
    """
    if df.empty:
        empty_df = pd.DataFrame(columns=list(df.columns))
        return empty_df, empty_df, empty_df

    # Concatena todas as colunas em uma string por linha, em minúsculas
    combined = (
        df.astype(str)
        .apply(lambda row: " ".join(row.values.astype(str)), axis=1)
        .str.lower()
    )

    mask_cert = combined.str.contains("certificado", na=False)
    mask_log = combined.str.contains("logbook", na=False)

    df_certificados = df[mask_cert].copy()
    df_logbook = df[mask_log].copy()

    # Resto = linhas que não têm nem certificado nem logbook
    mask_resto = ~(mask_cert | mask_log)
    df_resto = df[mask_resto].copy()

    return df_certificados, df_logbook, df_resto


def process_files(uploaded_files, keywords):
    """
    Percorre todos os arquivos enviados,
    aplica a BUSCA PRINCIPAL e CLASSIFICAÇÃO,
    acumulando os resultados em três DataFrames finais.
    """
    all_certificados = []
    all_logbook = []
    all_resto = []

    total_files = len(uploaded_files)
    progress_bar = st.progress(0)
    status_text = st.empty()

    for i, uploaded_file in enumerate(uploaded_files, start=1):
        try:
            status_text.text(f"Processando arquivo {i}/{total_files}: {uploaded_file.name}")

            # Lê todas as abas (worksheets) do arquivo
            # sheet_name=None devolve um dict {nome_aba: DataFrame}
            xls = pd.read_excel(
                uploaded_file,
                sheet_name=None,
                engine="openpyxl"
            )

            if not xls:
                continue  # Arquivo sem abas

            for sheet_name, df in xls.items():
                if df.empty:
                    continue

                # Garante cópia para não alterar o original
                df = df.copy()

                # Adiciona colunas de rastreabilidade
                # Assumindo que a linha 1 é o cabeçalho no Excel,
                # então a primeira linha de dados é a linha 2.
                df["Numero_Linha_Original"] = df.index + 2
                df["Nome_Arquivo_Origem"] = uploaded_file.name

                # Etapa 1: busca principal
                mask_primary = build_primary_mask(df, keywords)
                df_filtered = df[mask_primary].copy()

                if df_filtered.empty:
                    continue

                # Etapa 2: classificação
                df_cert, df_log, df_rest = classify_rows(df_filtered)

                if not df_cert.empty:
                    all_certificados.append(df_cert)
                if not df_log.empty:
                    all_logbook.append(df_log)
                if not df_rest.empty:
                    all_resto.append(df_rest)

        except Exception as e:
            st.error(f"Erro ao processar o arquivo {uploaded_file.name}: {e}")

        progress_bar.progress(i / total_files)

    status_text.text("Processamento concluído.")

    # Concatena resultados de todos os arquivos
    if all_certificados:
        df_certificados_final = pd.concat(all_certificados, ignore_index=True)
    else:
        df_certificados_final = pd.DataFrame()

    if all_logbook:
        df_logbook_final = pd.concat(all_logbook, ignore_index=True)
    else:
        df_logbook_final = pd.DataFrame()

    if all_resto:
        df_resto_final = pd.concat(all_resto, ignore_index=True)
    else:
        df_resto_final = pd.DataFrame()

    return df_certificados_final, df_logbook_final, df_resto_final


def df_to_excel_bytes(df: pd.DataFrame, sheet_name: str = "Resultados"):
    """
    Converte um DataFrame para bytes de um arquivo .xlsx em memória,
    pronto para usar em st.download_button.
    """
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    output.seek(0)
    return output


# ==========================
# Interface Streamlit
# ==========================

def main():
    st.title("📊 Processador de Planilhas Excel")
    st.markdown(
        """
Ferramenta para:

- Fazer upload de **múltiplos arquivos Excel (.xlsx)**.
- Buscar por **palavras-chave principais** em todas as colunas.
- Classificar linhas contendo **“certificado”**, **“logbook”** ou **outros**.
- Gerar três arquivos: **Certificados.xlsx**, **Logbook.xlsx** e **Resto.xlsx** com rastreabilidade completa.
        """
    )

    st.sidebar.header("Configurações")

    uploaded_files = st.sidebar.file_uploader(
        "Envie seus arquivos .xlsx",
        type=["xlsx"],
        accept_multiple_files=True
    )

    raw_keywords = st.sidebar.text_area(
        "Palavras-chave principais (separadas por vírgula)",
        placeholder="Ex: CASE12, CAT998, TESTE123"
    )

    process_button = st.sidebar.button("🚀 Processar Arquivos")

    if process_button:
        if not uploaded_files:
            st.error("Por favor, envie pelo menos um arquivo .xlsx.")
            return

        keywords = parse_keywords(raw_keywords)
        if not keywords:
            st.error("Por favor, informe pelo menos uma palavra-chave principal.")
            return

        with st.spinner("Processando arquivos, por favor aguarde..."):
            df_certificados, df_logbook, df_resto = process_files(uploaded_files, keywords)

        st.subheader("Resultados")

        if df_certificados.empty and df_logbook.empty and df_resto.empty:
            st.warning("Nenhuma linha encontrada com as palavras-chave informadas.")
            return

        col1, col2, col3 = st.columns(3)

        # Certificados
        with col1:
            st.markdown("### 📄 Certificados.xlsx")
            if df_certificados.empty:
                st.info("Nenhuma linha classificada como 'certificado'.")
            else:
                st.dataframe(df_certificados.head(50))
                excel_bytes = df_to_excel_bytes(df_certificados, sheet_name="Certificados")
                st.download_button(
                    label="⬇️ Baixar Certificados.xlsx",
                    data=excel_bytes,
                    file_name="Certificados.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

        # Logbook
        with col2:
            st.markdown("### 📘 Logbook.xlsx")
            if df_logbook.empty:
                st.info("Nenhuma linha classificada como 'logbook'.")
            else:
                st.dataframe(df_logbook.head(50))
                excel_bytes = df_to_excel_bytes(df_logbook, sheet_name="Logbook")
                st.download_button(
                    label="⬇️ Baixar Logbook.xlsx",
                    data=excel_bytes,
                    file_name="Logbook.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

        # Resto
        with col3:
            st.markdown("### 📂 Resto.xlsx")
            if df_resto.empty:
                st.info("Nenhuma linha classificada como 'Resto'.")
            else:
                st.dataframe(df_resto.head(50))
                excel_bytes = df_to_excel_bytes(df_resto, sheet_name="Resto")
                st.download_button(
                    label="⬇️ Baixar Resto.xlsx",
                    data=excel_bytes,
                    file_name="Resto.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )


if __name__ == "__main__":
    main()
