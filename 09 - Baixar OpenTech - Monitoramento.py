from datetime import datetime
import os
import shutil
import time
import zipfile

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException


# =============================================================================
# CONFIGURAÇÕES
# =============================================================================

USUARIO = "TKING47"
SENHA = "Op3nthT3ch#21"

PASTA_DOWNLOAD = r"C:\Users\flavi\OneDrive - transking.com.br\PROCESSOS\02 - Bases\32 - Acompanhamento Operacional\06 - OpenTech\01 - Monitoramento\00 - DownloadArquivo"

PASTA_DESTINO = r"C:\Users\flavi\OneDrive - transking.com.br\PROCESSOS\02 - Bases\32 - Acompanhamento Operacional\06 - OpenTech\01 - Monitoramento\01 - ArquivoBruto"

URL_LOGIN = "https://sil.opentechgr.com.br/Login.aspx"
URL_PAINEL = "https://sil.opentechgr.com.br/Forms/Operacao/PainelMonitoracao.aspx"


# =============================================================================
# CRIAR NAVEGADOR
# =============================================================================

def criar_navegador():

    os.makedirs(PASTA_DOWNLOAD, exist_ok=True)
    os.makedirs(PASTA_DESTINO, exist_ok=True)

    chrome_options = Options()

    # Se quiser rodar sem abrir a janela:
    # chrome_options.add_argument("--headless=new")

    chrome_options.add_argument("--window-size=1920,1080")

    prefs = {
        "download.default_directory": PASTA_DOWNLOAD,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    }

    chrome_options.add_experimental_option(
        "prefs",
        prefs
    )

    driver = webdriver.Chrome(
        options=chrome_options
    )

    return driver


# =============================================================================
# LIMPAR PASTA DOWNLOAD
# =============================================================================

def limpar_pasta_download():

    print("🧹 Limpando pasta de download...")

    os.makedirs(
        PASTA_DOWNLOAD,
        exist_ok=True
    )

    for arquivo in os.listdir(PASTA_DOWNLOAD):

        caminho = os.path.join(
            PASTA_DOWNLOAD,
            arquivo
        )

        try:

            if os.path.isfile(caminho):

                os.remove(caminho)

                print(
                    f"   🗑️ Removido: {arquivo}"
                )

        except Exception as e:

            print(
                f"   ⚠️ Não foi possível remover "
                f"{arquivo}: {e}"
            )


# =============================================================================
# AGUARDAR DOWNLOAD FINALIZAR
# =============================================================================

def aguardar_download(
    timeout=300,
    estabilidade_segundos=5
):

    print("\n⏳ Aguardando download do XLSX...")

    inicio = time.time()

    arquivo_crdownload_detectado = False
    ultimo_tamanho = None
    inicio_estabilidade = None

    while True:

        arquivos = os.listdir(
            PASTA_DOWNLOAD
        )

        # ---------------------------------------------------------
        # ARQUIVOS TEMPORÁRIOS DO CHROME
        # ---------------------------------------------------------

        arquivos_crdownload = [
            f
            for f in arquivos
            if f.lower().endswith(".crdownload")
        ]

        if arquivos_crdownload:

            if not arquivo_crdownload_detectado:

                print(
                    "⬇️ Download em andamento..."
                )

                print(
                    f"   Arquivo temporário: "
                    f"{arquivos_crdownload[0]}"
                )

                arquivo_crdownload_detectado = True

            # Enquanto existir .crdownload,
            # o download ainda está em andamento.

            ultimo_tamanho = None
            inicio_estabilidade = None

        # ---------------------------------------------------------
        # PROCURAR XLSX
        # ---------------------------------------------------------

        arquivos_xlsx = [
            f
            for f in arquivos
            if f.lower().endswith(".xlsx")
            and os.path.isfile(
                os.path.join(
                    PASTA_DOWNLOAD,
                    f
                )
            )
        ]

        # ---------------------------------------------------------
        # SE NÃO HOUVER CRDOWNLOAD E EXISTIR XLSX
        # ---------------------------------------------------------

        if (
            not arquivos_crdownload
            and arquivos_xlsx
        ):

            arquivo = max(
                [
                    os.path.join(
                        PASTA_DOWNLOAD,
                        f
                    )
                    for f in arquivos_xlsx
                ],
                key=os.path.getmtime
            )

            tamanho_atual = os.path.getsize(
                arquivo
            )

            # Arquivo vazio
            if tamanho_atual == 0:

                print(
                    "⚠️ XLSX encontrado, mas "
                    "arquivo está vazio."
                )

                ultimo_tamanho = None
                inicio_estabilidade = None

            else:

                # Primeira medição
                if ultimo_tamanho is None:

                    ultimo_tamanho = tamanho_atual
                    inicio_estabilidade = time.time()

                    print(
                        f"📄 XLSX encontrado: "
                        f"{os.path.basename(arquivo)}"
                    )

                    print(
                        f"   Tamanho atual: "
                        f"{tamanho_atual / 1024 / 1024:.2f} MB"
                    )

                # O tamanho mudou
                elif tamanho_atual != ultimo_tamanho:

                    print(
                        f"📈 Arquivo ainda sendo gravado..."
                    )

                    print(
                        f"   Tamanho anterior: "
                        f"{ultimo_tamanho / 1024 / 1024:.2f} MB"
                    )

                    print(
                        f"   Tamanho atual: "
                        f"{tamanho_atual / 1024 / 1024:.2f} MB"
                    )

                    ultimo_tamanho = tamanho_atual
                    inicio_estabilidade = time.time()

                # O tamanho não mudou
                else:

                    tempo_estavel = (
                        time.time()
                        - inicio_estabilidade
                    )

                    print(
                        f"⏱️ Arquivo estável há "
                        f"{tempo_estavel:.1f}s / "
                        f"{estabilidade_segundos}s"
                    )

                    # Arquivo permaneceu com o mesmo tamanho
                    # durante o período definido.
                    if tempo_estavel >= estabilidade_segundos:

                        print(
                            "\n✅ Download finalizado."
                        )

                        print(
                            f"📄 Arquivo: "
                            f"{os.path.basename(arquivo)}"
                        )

                        print(
                            f"📦 Tamanho final: "
                            f"{tamanho_atual / 1024 / 1024:.2f} MB"
                        )

                        return arquivo

        # ---------------------------------------------------------
        # TIMEOUT
        # ---------------------------------------------------------

        if time.time() - inicio > timeout:

            raise TimeoutException(
                "Tempo excedido aguardando "
                "download do arquivo XLSX."
            )

        time.sleep(1)


# =============================================================================
# VALIDAR XLSX
# =============================================================================

def validar_xlsx(caminho):

    print("\n🔍 Validando arquivo XLSX...")

    if not os.path.exists(caminho):

        raise Exception(
            "Arquivo XLSX não encontrado."
        )

    tamanho = os.path.getsize(
        caminho
    )

    if tamanho == 0:

        raise Exception(
            "O arquivo XLSX está vazio."
        )

    print(
        f"📦 Tamanho: "
        f"{tamanho / 1024 / 1024:.2f} MB"
    )

    # XLSX é um arquivo ZIP internamente.
    if not zipfile.is_zipfile(caminho):

        raise Exception(
            "O arquivo possui extensão .xlsx, "
            "mas não possui uma estrutura ZIP válida. "
            "O download pode estar corrompido."
        )

    try:

        with zipfile.ZipFile(
            caminho,
            "r"
        ) as arquivo_zip:

            arquivos = arquivo_zip.namelist()

            # Estruturas essenciais de um XLSX
            estruturas_obrigatorias = [
                "[Content_Types].xml",
                "xl/workbook.xml"
            ]

            for estrutura in estruturas_obrigatorias:

                if estrutura not in arquivos:

                    raise Exception(
                        f"O XLSX não possui "
                        f"a estrutura esperada: "
                        f"{estrutura}"
                    )

            # Testa a integridade do ZIP
            resultado = arquivo_zip.testzip()

            if resultado is not None:

                raise Exception(
                    f"Arquivo XLSX possui "
                    f"arquivo interno corrompido: "
                    f"{resultado}"
                )

    except zipfile.BadZipFile:

        raise Exception(
            "O arquivo XLSX está corrompido."
        )

    print(
        "✅ Estrutura XLSX validada com sucesso."
    )

    return True


# =============================================================================
# MOVER E RENOMEAR ARQUIVO
# =============================================================================

def mover_arquivo(arquivo_origem):

    print("\n📂 Preparando arquivo para destino...")

    # Confirma que realmente é XLSX
    extensao = os.path.splitext(
        arquivo_origem
    )[1].lower()

    if extensao != ".xlsx":

        raise Exception(
            f"O arquivo baixado não é XLSX. "
            f"Extensão encontrada: {extensao}"
        )

    timestamp = datetime.now().strftime(
        "%d_%m_%Y %H-%M"
    )

    nome_final = (
        f"OpenTech {timestamp}.xlsx"
    )

    caminho_destino = os.path.join(
        PASTA_DESTINO,
        nome_final
    )

    # Evita sobrescrever caso exista
    contador = 1

    while os.path.exists(caminho_destino):

        nome_final = (
            f"OpenTech {timestamp} "
            f"({contador}).xlsx"
        )

        caminho_destino = os.path.join(
            PASTA_DESTINO,
            nome_final
        )

        contador += 1

    print(
        f"📄 Origem: {arquivo_origem}"
    )

    print(
        f"📁 Destino: {caminho_destino}"
    )

    # Move somente depois da validação.
    shutil.move(
        arquivo_origem,
        caminho_destino
    )

    print(
        "\n✅ Arquivo salvo com sucesso:"
    )

    print(
        caminho_destino
    )


# =============================================================================
# LOGIN
# =============================================================================

def realizar_login(driver, wait):

    print("\n🔐 Realizando login...")

    driver.get(
        URL_LOGIN
    )

    # ---------------------------------------------------------
    # USUÁRIO
    # ---------------------------------------------------------

    campo_usuario = wait.until(
        EC.element_to_be_clickable(
            (
                By.ID,
                "txtUser"
            )
        )
    )

    campo_usuario.clear()

    campo_usuario.send_keys(
        USUARIO
    )

    # ---------------------------------------------------------
    # SENHA
    # ---------------------------------------------------------

    campo_senha = wait.until(
        EC.element_to_be_clickable(
            (
                By.ID,
                "txtPass"
            )
        )
    )

    campo_senha.clear()

    campo_senha.send_keys(
        SENHA
    )

    # ---------------------------------------------------------
    # LOGIN
    # ---------------------------------------------------------

    botao_login = wait.until(
        EC.element_to_be_clickable(
            (
                By.ID,
                "btnAutenticar"
            )
        )
    )

    botao_login.click()

    # Dá tempo para o servidor processar o login.
    time.sleep(3)

    print(
        "✅ Login realizado."
    )


# =============================================================================
# ACESSAR PAINEL
# =============================================================================

def acessar_painel(driver, wait):

    print("\n📊 Acessando painel...")

    driver.get(
        URL_PAINEL
    )

    wait.until(
        EC.presence_of_element_located(
            (
                By.TAG_NAME,
                "body"
            )
        )
    )

    time.sleep(3)

    print(
        "✅ Painel carregado."
    )


# =============================================================================
# BAIXAR BASE
# =============================================================================

def baixar_base(driver, wait):

    print(
        "\n⬇️ Iniciando exportação da base..."
    )

    # ---------------------------------------------------------
    # PRIMEIRO BOTÃO
    # ---------------------------------------------------------

    botao_exportar = wait.until(
        EC.element_to_be_clickable(
            (
                By.XPATH,
                "/html/body/form/div[5]/div[1]/span[2]/div[6]/div/div"
            )
        )
    )

    driver.execute_script(
        "arguments[0].click();",
        botao_exportar
    )

    print(
        "✅ Botão de exportação acionado."
    )

    time.sleep(2)

    # ---------------------------------------------------------
    # SEGUNDO BOTÃO
    # ---------------------------------------------------------

    botao_confirmar = wait.until(
        EC.element_to_be_clickable(
            (
                By.XPATH,
                "/html/body/form/div[5]/div[1]/span[2]/div[8]/div/div"
            )
        )
    )

    driver.execute_script(
        "arguments[0].click();",
        botao_confirmar
    )

    print(
        "✅ Exportação confirmada."
    )

    time.sleep(1)

    # ---------------------------------------------------------
    # POPUP
    # ---------------------------------------------------------

    botao_popup = wait.until(
        EC.element_to_be_clickable(
            (
                By.XPATH,
                "/html/body/div[3]/div[3]/div/button[1]"
            )
        )
    )

    driver.execute_script(
        "arguments[0].click();",
        botao_popup
    )

    print(
        "✅ Confirmação do popup realizada."
    )

    # Pequena espera para o Chrome iniciar o download.
    time.sleep(2)


# =============================================================================
# MAIN
# =============================================================================

def main():

    driver = None

    try:

        print("=" * 70)
        print("🚀 INICIANDO EXTRAÇÃO OPENTECH")
        print("=" * 70)

        # -----------------------------------------------------
        # LIMPAR DOWNLOADS ANTERIORES
        # -----------------------------------------------------

        limpar_pasta_download()

        # -----------------------------------------------------
        # CRIAR NAVEGADOR
        # -----------------------------------------------------

        driver = criar_navegador()

        wait = WebDriverWait(
            driver,
            60
        )

        # -----------------------------------------------------
        # LOGIN
        # -----------------------------------------------------

        realizar_login(
            driver,
            wait
        )

        # -----------------------------------------------------
        # PAINEL
        # -----------------------------------------------------

        acessar_painel(
            driver,
            wait
        )

        # -----------------------------------------------------
        # DOWNLOAD
        # -----------------------------------------------------

        baixar_base(
            driver,
            wait
        )

        # -----------------------------------------------------
        # AGUARDAR DOWNLOAD
        # -----------------------------------------------------

        arquivo_baixado = aguardar_download(
            timeout=300,
            estabilidade_segundos=5
        )

        # -----------------------------------------------------
        # VALIDAR XLSX
        # -----------------------------------------------------

        validar_xlsx(
            arquivo_baixado
        )

        # -----------------------------------------------------
        # MOVER ARQUIVO
        # -----------------------------------------------------

        mover_arquivo(
            arquivo_baixado
        )

        print("\n" + "=" * 70)
        print("✅ PROCESSO FINALIZADO COM SUCESSO")
        print("=" * 70)

    except Exception as e:

        print("\n" + "=" * 70)
        print("❌ ERRO DURANTE EXECUÇÃO")
        print("=" * 70)

        print(
            f"\n{type(e).__name__}: {e}"
        )

        # -----------------------------------------------------
        # DIAGNÓSTICO DA PASTA
        # -----------------------------------------------------

        print(
            "\n📂 Conteúdo atual da pasta de download:"
        )

        try:

            arquivos = os.listdir(
                PASTA_DOWNLOAD
            )

            if not arquivos:

                print(
                    "   Pasta vazia."
                )

            else:

                for arquivo in arquivos:

                    caminho = os.path.join(
                        PASTA_DOWNLOAD,
                        arquivo
                    )

                    if os.path.isfile(caminho):

                        tamanho = os.path.getsize(
                            caminho
                        )

                        print(
                            f"   - {arquivo} | "
                            f"{tamanho / 1024 / 1024:.2f} MB"
                        )

        except Exception as erro_pasta:

            print(
                f"   Não foi possível "
                f"listar a pasta: {erro_pasta}"
            )

    finally:

        if driver:

            print(
                "\n🌐 Encerrando navegador..."
            )

            driver.quit()

            print(
                "✅ Navegador encerrado."
            )


# =============================================================================
# EXECUÇÃO
# =============================================================================

if __name__ == "__main__":

    main()
