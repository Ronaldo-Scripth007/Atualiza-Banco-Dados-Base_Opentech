import os
import pandas as pd

# =============================================================================
# CAMINHOS
# =============================================================================

PASTA_BRUTO = r"C:\Users\flavi\OneDrive - transking.com.br\PROCESSOS\02 - Bases\32 - Acompanhamento Operacional\06 - OpenTech\01 - Monitoramento\01 - ArquivoBruto"

PASTA_TRATADO = r"C:\Users\flavi\OneDrive - transking.com.br\PROCESSOS\02 - Bases\32 - Acompanhamento Operacional\06 - OpenTech\01 - Monitoramento\02 - ArquivoTratado"

os.makedirs(PASTA_TRATADO, exist_ok=True)


# =============================================================================
# FUNÇÃO PARA PEGAR O ARQUIVO MAIS RECENTE
# =============================================================================

def obter_arquivo_mais_recente(pasta):

    arquivos = [
        os.path.join(pasta, f)
        for f in os.listdir(pasta)
        if f.lower().endswith(".xlsx")
    ]

    if not arquivos:
        raise FileNotFoundError("Nenhum arquivo XLSX encontrado.")

    return max(arquivos, key=os.path.getctime)


# =============================================================================
# FORMATAR PLACA
# Remove:
# - espaços
# - tabs
# - quebras de linha
# - caracteres invisíveis
# =============================================================================

def formatar_placa(valor):

    if pd.isna(valor):
        return ""

    placa = str(valor)

    # Remove caracteres invisíveis
    placa = placa.replace("\xa0", "")

    # Remove tabs e quebras
    placa = placa.replace("\t", "")
    placa = placa.replace("\n", "")
    placa = placa.replace("\r", "")

    # Remove TODOS os espaços
    placa = "".join(placa.split())

    # Maiúsculo
    placa = placa.upper()

    return placa


# =============================================================================
# FORMATAR KM RESTANTE
# Exemplo:
# 1500.75 -> 1500,75
# =============================================================================

def formatar_km(valor):

    if pd.isna(valor):
        return ""

    try:

        valor = float(valor)

        return f"{valor:.2f}".replace(".", ",")

    except:

        return str(valor).replace(".", ",")


# =============================================================================
# FORMATAR % CONCLUSÃO
# Exemplo:
# 50 -> 0,5
# 75 -> 0,75
# =============================================================================

def formatar_percentual(valor):

    if pd.isna(valor):
        return ""

    try:

        valor = float(valor)

        # Caso venha como 50, 75 etc
        if valor > 1:
            valor = valor / 100

        # Remove zeros desnecessários
        valor_formatado = f"{valor:.4f}".rstrip("0").rstrip(".")

        return valor_formatado.replace(".", ",")

    except:

        return str(valor).replace(".", ",")


# =============================================================================
# FUNÇÃO PRINCIPAL
# =============================================================================

def main():

    try:

        # ======================================================
        # ARQUIVO MAIS RECENTE
        # ======================================================

        arquivo_origem = obter_arquivo_mais_recente(PASTA_BRUTO)

        print("📂 Arquivo utilizado:")
        print(arquivo_origem)

        # ======================================================
        # LEITURA XLSX
        # ======================================================

        df = pd.read_excel(arquivo_origem)

        print(f"📊 Linhas originais: {len(df)}")

        # ======================================================
        # COLUNAS DESEJADAS
        # ======================================================

        colunas_desejadas = [
            "Placa veículo",
            "Km Restante",
            "% conclusão da viagem",
            "Referência posição cidade"
        ]

        colunas_existentes = [
            col for col in colunas_desejadas
            if col in df.columns
        ]

        if not colunas_existentes:

            raise Exception(
                "Nenhuma das colunas desejadas foi encontrada."
            )

        df = df[colunas_existentes].copy()

        print("🧹 Colunas filtradas")

        # ======================================================
        # TRATAMENTO PLACA
        # ======================================================

        if "Placa veículo" in df.columns:

            df["Placa veículo"] = df[
                "Placa veículo"
            ].apply(formatar_placa)

        print("🚚 Placas tratadas")

        # ======================================================
        # FORMATAÇÃO KM
        # ======================================================

        if "Km Restante" in df.columns:

            df["Km Restante"] = df[
                "Km Restante"
            ].apply(formatar_km)

        # ======================================================
        # FORMATAÇÃO %
        # ======================================================

        if "% conclusão da viagem" in df.columns:

            df["% conclusão da viagem"] = df[
                "% conclusão da viagem"
            ].apply(formatar_percentual)

        print("🔧 Formatações ajustadas")

        # ======================================================
        # CAMINHO SAÍDA
        # ======================================================

        caminho_saida = os.path.join(
            PASTA_TRATADO,
            "OpenTech_Monitoramento.csv"
        )

        # Remove antigo
        if os.path.exists(caminho_saida):

            os.remove(caminho_saida)

        # ======================================================
        # SALVAR CSV
        # ======================================================

        df.to_csv(
            caminho_saida,
            sep=';',
            index=False,
            encoding='latin1'
        )

        print("✅ Arquivo salvo:")
        print(caminho_saida)

    except Exception as e:

        print("❌ Erro no processamento:")
        print(str(e))


# =============================================================================
# EXECUÇÃO
# =============================================================================

if __name__ == "__main__":
    main()