'''OpenTech - Monitoramento
Baixa a base do painel, pega só as colunas necessárias e sobe para o Supabase.
Feito para rodar de hora em hora no Agendador de Tarefas.

A cada execução:
  1) apaga os arquivos que estavam na pasta de download
  2) baixa o XLSX novo (o arquivo original fica inteiro na pasta, com todas as colunas)
  3) substitui o conteúdo da tabela public.opentech_monitoramento pela foto atual
     (tudo numa transação: se der erro, a tabela continua intacta)
'''

import os
import sys
import time
import unicodedata
import zipfile
from pathlib import Path
import warnings


import psycopg2
from dotenv import load_dotenv
from openpyxl import load_workbook
from psycopg2 import sql
from psycopg2.extras import execute_values
from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

warnings.filterwarnings("ignore", message="Workbook contains no default style")
# =====================================================================
# CONFIGURAÇÕES (.env na mesma pasta deste arquivo)
# =====================================================================
PASTA_BASE = Path(__file__).resolve().parent

load_dotenv(PASTA_BASE / ".env")

USUARIO = os.getenv("OPENTECH_USER", "").strip()
SENHA = os.getenv("OPENTECH_PASS", "").strip()
SUPABASE_DB_URL = os.getenv("SUPABASE_DB_URL", "").strip()

if not (USUARIO and SENHA and SUPABASE_DB_URL):
    raise RuntimeError("Faltam OPENTECH_USER, OPENTECH_PASS e/ou SUPABASE_DB_URL no arquivo .env.")

PASTA_DOWNLOAD = os.getenv("PASTA_DOWNLOAD", "").strip()

URL_LOGIN = "https://sil.opentechgr.com.br/Login.aspx"
URL_PAINEL = "https://sil.opentechgr.com.br/Forms/Operacao/PainelMonitoracao.aspx"

TABELA = "opentech_monitoramento"  # tabela no Supabase (schema public)
HEADLESS = True                    # True = Chrome oculto (ideal para o agendador)

# Colunas do relatório que vão para o banco (cabeçalho do XLSX -> campo interno)
COLUNAS = {
    "placa veiculo": "placa",
    "km restante": "km",
    "% conclusao da viagem": "pct",
    "referencia posicao cidade": "cidade",
}


# =====================================================================
# FUNÇÕES AUXILIARES - ARQUIVOS
# =====================================================================
def limpar_pasta_download():
    """Apaga o que sobrou do ciclo anterior (só no início da execução)."""
    os.makedirs(PASTA_DOWNLOAD, exist_ok=True)

    for nome in os.listdir(PASTA_DOWNLOAD):
        caminho = os.path.join(PASTA_DOWNLOAD, nome)

        try:
            if os.path.isfile(caminho):
                os.remove(caminho)
                print(f"[info] Removido: {nome}")
        except Exception as e:
            print(f"[aviso] Não foi possível remover {nome}: {e}")


def aguardar_download(timeout=300, estabilidade=5):
    """Espera o XLSX aparecer e parar de crescer."""
    print("[info] Aguardando download do XLSX...")

    inicio = time.time()
    ultimo = None
    desde = None

    while time.time() - inicio < timeout:
        arquivos = os.listdir(PASTA_DOWNLOAD)

        if any(f.lower().endswith(".crdownload") for f in arquivos):
            ultimo = None  # Chrome ainda está baixando
        else:
            xlsx = [os.path.join(PASTA_DOWNLOAD, f) for f in arquivos if f.lower().endswith(".xlsx")]

            if xlsx:
                arquivo = max(xlsx, key=os.path.getmtime)
                tamanho = os.path.getsize(arquivo)

                if tamanho > 0 and tamanho == ultimo:
                    if time.time() - desde >= estabilidade:
                        print(f"[ok] Download finalizado: {os.path.basename(arquivo)} ({tamanho / 1024 / 1024:.2f} MB)")
                        return arquivo
                else:
                    ultimo = tamanho
                    desde = time.time()

        time.sleep(1)

    raise TimeoutException("Tempo excedido aguardando o download do XLSX.")


def validar_xlsx(caminho):
    """Confere se o arquivo é mesmo um XLSX íntegro."""
    if not os.path.exists(caminho) or os.path.getsize(caminho) == 0:
        raise RuntimeError("O arquivo XLSX não existe ou está vazio.")

    if not zipfile.is_zipfile(caminho):
        raise RuntimeError("O arquivo baixado não é um XLSX válido (download corrompido?).")

    with zipfile.ZipFile(caminho) as z:
        nomes = z.namelist()

        for obrigatorio in ("[Content_Types].xml", "xl/workbook.xml"):
            if obrigatorio not in nomes:
                raise RuntimeError(f"O XLSX não tem a estrutura esperada: {obrigatorio}")

        if z.testzip() is not None:
            raise RuntimeError("O XLSX tem arquivo interno corrompido.")

    print("[ok] XLSX validado.")


# =====================================================================
# FUNÇÕES AUXILIARES - TRATAMENTO DOS DADOS
# =====================================================================
def normalizar_cabecalho(texto):
    """'Placa veículo' -> 'placa veiculo'"""
    texto = unicodedata.normalize("NFD", str(texto or ""))
    texto = "".join(c for c in texto if unicodedata.category(c) != "Mn")

    return " ".join(texto.lower().split())


def formatar_placa(valor):
    """Tira espaços, quebras e caracteres invisíveis; deixa maiúsculo."""
    if valor is None:
        return ""

    placa = str(valor).replace("\u200b", "")  # str.split() já remove espaços, tabs, quebras e \xa0

    return "".join(placa.split()).upper()


def para_numero(valor):
    """Converte para float (aceita número ou texto com vírgula). Vazio vira None."""
    if valor is None:
        return None

    if isinstance(valor, (int, float)):
        return float(valor)

    s = str(valor).strip().replace("%", "")

    if not s:
        return None

    if "," in s:
        s = s.replace(".", "").replace(",", ".")

    try:
        return float(s)
    except ValueError:
        return None


def para_percentual(valor):
    """50 -> 0.5 | 0.5 -> 0.5 (mesma regra do tratamento antigo)."""
    v = para_numero(valor)

    if v is None:
        return None

    if v > 1:
        v = v / 100

    return round(v, 4)


def para_texto(valor):
    if valor is None:
        return None

    return str(valor).strip() or None


def ler_planilha(caminho):
    """Lê o XLSX e devolve só as 4 colunas: (placa, km, pct, cidade). O arquivo original não é alterado."""
    planilha = load_workbook(caminho, read_only=True, data_only=True)
    linhas = planilha.active.iter_rows(values_only=True)

    cabecalho = [normalizar_cabecalho(c) for c in next(linhas, [])]

    pos = {}
    for nome, campo in COLUNAS.items():
        if nome not in cabecalho:
            raise RuntimeError(f"Coluna '{nome}' não encontrada no relatório. Colunas lidas: {cabecalho}")
        pos[campo] = cabecalho.index(nome)

    dados = []

    for linha in linhas:
        linha = list(linha) + [None] * len(cabecalho)

        placa = formatar_placa(linha[pos["placa"]])

        if not placa:
            continue

        dados.append(
            (
                placa,
                para_numero(linha[pos["km"]]),
                para_percentual(linha[pos["pct"]]),
                para_texto(linha[pos["cidade"]]),
            )
        )

    planilha.close()

    print(f"[info] Linhas lidas do relatório: {len(dados)}")

    return dados


# =====================================================================
# SELENIUM
# =====================================================================
def criar_navegador():
    opcoes = Options()

    if HEADLESS:
        opcoes.add_argument("--headless=new")

    opcoes.add_argument("--window-size=1920,1080")

    opcoes.add_experimental_option(
        "prefs",
        {
            "download.default_directory": PASTA_DOWNLOAD,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True,
        },
    )

    return webdriver.Chrome(options=opcoes)


def realizar_login(driver, wait):
    print("[info] Realizando login...")

    driver.get(URL_LOGIN)

    campo = wait.until(EC.element_to_be_clickable((By.ID, "txtUser")))
    campo.clear()
    campo.send_keys(USUARIO)

    campo = wait.until(EC.element_to_be_clickable((By.ID, "txtPass")))
    campo.clear()
    campo.send_keys(SENHA)

    wait.until(EC.element_to_be_clickable((By.ID, "btnAutenticar"))).click()

    time.sleep(3)  # tempo para o servidor processar o login

    print("[ok] Login realizado.")


def acessar_painel(driver, wait):
    print("[info] Acessando painel...")

    driver.get(URL_PAINEL)
    wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))

    time.sleep(3)

    print("[ok] Painel carregado.")


def clicar_js(driver, wait, xpath):
    elemento = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
    driver.execute_script("arguments[0].click();", elemento)


def baixar_base(driver, wait):
    print("[info] Iniciando exportação da base...")

    clicar_js(driver, wait, "/html/body/form/div[5]/div[1]/span[2]/div[6]/div/div")
    time.sleep(2)

    clicar_js(driver, wait, "/html/body/form/div[5]/div[1]/span[2]/div[8]/div/div")
    time.sleep(1)

    clicar_js(driver, wait, "/html/body/div[3]/div[3]/div/button[1]")  # popup de confirmação
    time.sleep(2)  # o Chrome inicia o download

    print("[ok] Exportação solicitada.")


# =====================================================================
# SUPABASE (substitui o conteúdo da tabela pela foto atual, numa transação)
# =====================================================================
def enviar_para_supabase(dados):
    if not dados:
        raise RuntimeError("O relatório veio sem linhas. Tabela mantida como estava.")

    conn = psycopg2.connect(SUPABASE_DB_URL)

    try:
        with conn, conn.cursor() as cur:
            tabela = sql.Identifier("public", TABELA)

            cur.execute(
                sql.SQL(
                    "create table if not exists {} ("
                    "id bigint generated always as identity primary key, "
                    "placa_veiculo text, "
                    "km_restante numeric(12,2), "
                    "pct_conclusao_viagem numeric(7,4), "
                    "referencia_posicao_cidade text, "
                    "atualizado_em timestamptz not null default now())"
                ).format(tabela)
            )

            cur.execute(sql.SQL("alter table {} enable row level security").format(tabela))

            cur.execute(sql.SQL("truncate table {} restart identity").format(tabela))

            comando = sql.SQL(
                "insert into {} (placa_veiculo, km_restante, pct_conclusao_viagem, "
                "referencia_posicao_cidade) values %s"
            ).format(tabela)

            execute_values(cur, comando.as_string(cur), dados, page_size=1000)

            cur.execute(sql.SQL("select count(*) from {}").format(tabela))
            total = cur.fetchone()[0]

        print(f"[ok] Tabela public.{TABELA} atualizada: {total} linhas.")

    finally:
        conn.close()


# =====================================================================
# MAIN
# =====================================================================
def main():
    driver = None

    try:
        print("=" * 60)
        print("INICIANDO EXTRAÇÃO OPENTECH")
        print("=" * 60)

        limpar_pasta_download()

        driver = criar_navegador()
        wait = WebDriverWait(driver, 60)

        realizar_login(driver, wait)
        acessar_painel(driver, wait)
        baixar_base(driver, wait)

        arquivo = aguardar_download()

        validar_xlsx(arquivo)

        dados = ler_planilha(arquivo)

        enviar_para_supabase(dados)

        print("=" * 60)
        print("PROCESSO FINALIZADO COM SUCESSO")
        print("=" * 60)

        return 0

    except Exception as e:
        print("=" * 60)
        print(f"ERRO DURANTE A EXECUÇÃO: {type(e).__name__}: {e}")
        print("=" * 60)

        try:
            print("[info] Conteúdo da pasta de download:")
            for nome in os.listdir(PASTA_DOWNLOAD) or ["(vazia)"]:
                print(f"   - {nome}")
        except Exception:
            pass

        return 1

    finally:
        if driver:
            driver.quit()
            print("[info] Navegador encerrado.")


if __name__ == "__main__":
    sys.exit(main())