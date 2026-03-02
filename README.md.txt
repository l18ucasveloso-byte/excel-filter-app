# 📊 Processador de Planilhas Excel (Streamlit)

Aplicação web em **Python + Streamlit** para:

- Fazer upload de **vários arquivos Excel (.xlsx)**.
- Buscar **palavras-chave principais** em todas as colunas de cada linha.
- Classificar automaticamente linhas que contenham:
  - `"certificado"` → **Certificados.xlsx**
  - `"logbook"` → **Logbook.xlsx**
  - Nenhum dos dois → **Resto.xlsx**
- Manter **rastreamento completo** da origem:
  - `Numero_Linha_Original`
  - `Nome_Arquivo_Origem`

---

## 🧠 Funcionamento

1. O usuário faz upload de múltiplos arquivos `.xlsx`.
2. Informa uma ou mais palavras-chave principais, separadas por vírgula (ex.: `CASE12, CAT998`).
3. O sistema:
   - Percorre **todas as abas** de todos os arquivos.
   - Busca as palavras-chave em todas as colunas de cada linha (case-insensitive).
   - Mantém apenas as linhas que contêm **pelo menos uma** palavra-chave principal.
4. Para cada linha filtrada:
   - Se contiver `"certificado"` (em qualquer coluna, case-insensitive) → vai para **Certificados.xlsx**.
   - Se contiver `"logbook"` → vai para **Logbook.xlsx**.
   - Se não contiver nenhuma dessas palavras → vai para **Resto.xlsx**.
   - Se contiver **ambas**, pode aparecer nos **dois** arquivos.
5. Cada linha exportada contém:
   - Todas as colunas originais.
   - `Numero_Linha_Original` (linha de dados no arquivo original, assumindo cabeçalho na linha 1).
   - `Nome_Arquivo_Origem` (nome do arquivo de onde veio a linha).

---

## 🛠 Requisitos

- Python 3.9 ou superior
- `pip` instalado

Bibliotecas Python (também listadas em `requirements.txt`):

- `streamlit`
- `pandas`
- `openpyxl`

---

## ▶️ Como rodar localmente

1. **Clonar o repositório** ou baixar os arquivos

   ```bash
   git clone https://github.com/SEU_USUARIO/SEU_REPOSITORIO.git
   cd SEU_REPOSITORIO
