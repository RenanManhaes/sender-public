"""
Sender - automação de disparo de WhatsApp a partir de uma planilha Excel.

Fluxo:
  1. Mostra a última atualização da planilha e quantos contatos tem.
  2. Pergunta se é para fazer um teste (envia para um número seu) ou disparar.
  3. Abre o WhatsApp Web e espera o login pelo QR code.
  4. Para cada contato: abre a conversa já com a mensagem no link, espera,
     aperta Enter e grava o resultado no Registros.md.

Dependências:  pip install -r requirements.txt
Uso:           python sender.py            (ou o atalho Sender.bat)
Documentação:  README.md e docs/
"""

import math
import msvcrt
import os
import random
import re
import shutil
import time
from collections import Counter
from datetime import datetime
from pathlib import Path
from urllib.parse import quote

import pandas as pd
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException, WebDriverException
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

# ============================== CONFIGURAÇÃO ==============================
# Ajuste estes valores para o seu caso. Os caminhos padrão ficam ao lado do
# script, em "dados/"; se preferir, aponte para qualquer outra pasta
# (por exemplo, um cofre do Obsidian).

PASTA = Path(__file__).resolve().parent

PLANILHA = str(PASTA / "dados" / "contatos.xlsx")
ABA = 0                               # aba lida (0 = primeira)
COLUNA_NOME = "Nome"                  # cabeçalho da coluna do nome usado na mensagem
COLUNA_NUMERO = "Telefone"            # cabeçalho da coluna do telefone

# Registro dos envios em Markdown: histórico para você e memória do robô
ARQUIVO_REGISTRO = str(PASTA / "dados" / "Registros.md")
PULAR_JA_ENVIADOS = True              # ignora números que já estão como "Enviado" no registro

# Nota com números que NUNCA entram em nenhuma lista (Sender e verificador).
# Um número por linha, em qualquer formato.
ARQUIVO_NUNCA_CONTATAR = str(Path(ARQUIVO_REGISTRO).with_name("Nunca contatar.md"))

# Pasta onde o Chrome do robô guarda o login do WhatsApp (separada do seu Chrome)
PERFIL_CHROME = str(Path.home() / ".sender-chrome")

# Timers (segundos)
TEMPO_LOGIN_QRCODE = 120              # espera para escanear o QR code (segue sozinho ao logar)
TEMPO_BUSCAR = 15                     # espera a conversa carregar
TEMPO_BUSCAR_EXTRA = 30               # se não carregou, espera até mais isso (internet lenta)
TEMPO_ANTES_ENTER = 5                 # espera antes de apertar Enter
TEMPO_TEXTO_EXTRA = 10                # se a mensagem não apareceu na caixa, espera até mais isso
TEMPO_APOS_ENTER = 5                  # garante que a mensagem saiu antes de trocar de página
TEMPO_ENTRE_CONTATOS = (20, 60)       # aleatório entre min e máx, antes do próximo contato

LIMITE_ERROS_SEGUIDOS = 3             # pausa o disparo e pergunta se deve continuar

# Mensagens disponíveis. Troque a ativa pelo menu do Sender ([3] Trocar a mensagem)
# ou mudando MENSAGEM_ATIVA. {Nome} vira o valor da coluna COLUNA_NOME da planilha.
# A memória de "quem já recebeu" é separada por mensagem.
MENSAGENS = {
    "convite": {
        "nome": "Convite (exemplo)",
        "texto": """Olá {Nome}! Tudo bem?

Estou passando para te convidar para o nosso próximo encontro.

Se quiser saber mais, é só acessar:
🔗 https://exemplo.com.br/evento

Espero encontrar você lá!""",
    },
    "grupo": {
        "nome": "Convite para um grupo (exemplo)",
        "texto": """Olá {Nome}! Tudo bem?

Criamos um grupo no WhatsApp para quem quer acompanhar as novidades.

Entre pelo link:
🔗 https://chat.whatsapp.com/EXEMPLO

Espero você lá!""",
    },
}
MENSAGEM_ATIVA = "convite"

# O wa.me abre uma página intermediária ("Continuar para a conversa") no
# computador; o web.whatsapp.com/send vai direto para a conversa no WhatsApp Web.
LINK_BASE = "https://web.whatsapp.com/send?phone={numero}&text={texto}"

# ==========================================================================

LINHA = "=" * 60


# ----------------------------- dados / mensagem ---------------------------

def nome_da_planilha(valor):
    """Nome da coluna B, do jeito que está na planilha (já vem tratado)."""
    return "" if pd.isna(valor) else str(valor).strip()


PARTICULAS = {"de", "da", "do", "das", "dos", "e"}   # usado no nome digitado no teste


def primeiro_nome(nome_completo):
    """'maria silva souza' -> 'Maria'; 'de souza de souza' -> 'Souza'."""
    if pd.isna(nome_completo):
        return ""
    palavras = [p for p in str(nome_completo).split() if p.lower() not in PARTICULAS]
    return palavras[0].capitalize() if palavras else ""


def formatar_numero(valor):
    """15991234567 -> 5515991234567. Retorna None se não parecer um número BR.

    Celular BR com 11 dígitos sempre tem 9 depois do DDD. Números de 11 dígitos
    sem esse 9 (ex.: 13055550123) são de fora do Brasil (EUA +1) e ficam de fora:
    com 55 na frente iriam para um número brasileiro errado.
    """
    if pd.isna(valor):
        return None
    digitos = re.sub(r"\D", "", str(valor).removesuffix(".0"))
    if digitos.startswith("55") and len(digitos) in (12, 13):
        digitos = digitos[2:]
    if len(digitos) == 11 and digitos[2] == "9":     # celular: DDD + 9xxxx-xxxx
        return "55" + digitos
    if len(digitos) == 10 and digitos[2] in "2345":   # fixo: DDD + 2xxx..5xxx-xxxx
        return "55" + digitos
    if len(digitos) == 10 and digitos[2] in "6789":   # celular antigo, sem o 9 na frente
        return "55" + digitos[:2] + "9" + digitos[2:]
    return None


def nunca_contatar():
    """Números da nota 'Nunca contatar.md' (qualquer sequência de 10+ dígitos)."""
    arquivo = Path(ARQUIVO_NUNCA_CONTATAR)
    if not arquivo.exists():
        return set()
    texto = arquivo.read_text(encoding="utf-8")
    candidatos = re.findall(r"\+?\d[\d\s().-]{8,}\d", texto)
    return {n for n in map(formatar_numero, candidatos) if n}


def bloqueado(numero):
    """True se o número está na nota 'Nunca contatar.md'."""
    return numero is not None and formatar_numero(numero) in nunca_contatar()


def montar_mensagem(nome):
    texto = MENSAGENS[MENSAGEM_ATIVA]["texto"]
    return texto.replace("{Nome}", nome) if nome else texto.replace(" {Nome}", "")


def nome_mensagem(chave=None):
    return MENSAGENS[chave or MENSAGEM_ATIVA]["nome"]


def mensagem_para_link(texto):
    """Converte a mensagem para o formato de URL (espaços, quebras de linha, emojis)."""
    return quote(texto, safe="")


def montar_link(numero, texto):
    return LINK_BASE.format(numero=numero, texto=mensagem_para_link(texto))


class ParadaEmergencia(Exception):
    """Disparada quando o usuário aperta ESC."""


RESUMO = Counter()                    # contagem de status desta execução

# --------------------------- registro na pasta de dados --------------------------
#
# Registros.md fica assim (uma seção por disparo, mais recente embaixo):
#
#   ## 🚀 Disparo oficial — 18/09/2026 às 18:30
#   | Nº | Nome | Número | Status | Horário |
#   |---:|---|---|---|---|
#   | 2 | Fulano | 5511999999999 | ✅ Enviado | 18:31:05 |
#   > **Resumo:** ✅ 1 enviado · ⏱ 1min · finalizado às 18:32

ICONES = {"Enviado": "✅", "Teste enviado": "🧪", "Número errado": "❌",
          "Número fora do padrão": "❌"}

# Falhas que aconteceram ANTES do Enter: a mensagem com certeza não saiu, então
# o contato volta a ser tentado no próximo disparo. Todo o resto é definitivo.
ERROS_RETENTAVEIS = {"Erro: conversa não carregou", "Erro: mensagem não apareceu",
                     "Erro: Chrome não respondeu", "Erro: texto dobrado na caixa"}
ICONES["Pulado por você"] = "⏭️"      # definitivo: nunca recebe


def escrever_registro(texto):
    with open(ARQUIVO_REGISTRO, "a", encoding="utf-8") as f:
        f.write(texto)


def iniciar_secao(tipo, quantidade):
    """Cabeçalho do arquivo (se novo) + título e tabela deste disparo."""
    arquivo = Path(ARQUIVO_REGISTRO)
    arquivo.parent.mkdir(parents=True, exist_ok=True)
    if not arquivo.exists() or not arquivo.read_text(encoding="utf-8").strip():
        arquivo.write_text(
            "# 📤 Sender — Registros\n\n"
            "> Registro automático dos disparos de WhatsApp do Sender.\n"
            "> Cada disparo ganha uma seção; o mais recente fica no final.\n"
            "> ⚠️ Não edite as tabelas: o Sender lê este arquivo para saber quem já recebeu.\n",
            encoding="utf-8")

    icone = "🧪" if tipo == "Teste" else "🚀"
    escrever_registro(
        f"\n---\n\n## {icone} {tipo} — {datetime.now():%d/%m/%Y às %H:%M}\n\n"
        f"- **Planilha:** `{Path(PLANILHA).name}`\n"
        f"- **Mensagem:** {nome_mensagem()} (`{MENSAGEM_ATIVA}`)\n"
        f"- **Contatos selecionados:** {quantidade}\n\n"
        "| Nº | Nome | Número | Status | Horário |\n"
        "|---:|---|---|---|---|\n")


def registrar(posicao, nome, numero, status):
    """Uma linha na tabela do disparo atual + eco no terminal."""
    nome_md = str(nome).replace("|", "/")
    icone = ICONES.get(status, "⚠️")
    escrever_registro(f"| {posicao} | {nome_md} | {numero} | {icone} {status} | "
                      f"{datetime.now():%d/%m %H:%M:%S} |\n")
    RESUMO[status] += 1
    print(f"  {icone} {status}: {nome} - {numero}")
    return status


def fechar_secao(inicio, interrompido):
    partes = [f"{ICONES.get(s, '⚠️')} {s}: **{qtd}**" for s, qtd in RESUMO.items()]
    partes.append(f"⏱ {formatar_duracao(time.time() - inicio)}")
    fim = "🛑 **interrompido** às" if interrompido else "finalizado às"
    partes.append(f"{fim} {datetime.now():%H:%M}")
    escrever_registro(f"\n> **Resumo:** {' · '.join(partes or ['nenhum envio'])}\n")


def ler_historico():
    """Lê o Registros.md — é a "memória" do Sender entre um disparo e outro.

    Retorna:
      enviados: números que já receberam a MENSAGEM ATIVA
      falhas:   números com falha definitiva (errado, fora do padrão, "verificar",
                pulado) — valem para qualquer mensagem; não são tentados de novo.
                Erros antes do Enter (ERROS_RETENTAVEIS) não entram aqui.
      ultimo:   (linha, nome, número, horário) do último envio da mensagem ativa

    Seções sem a linha "Mensagem" são dos primeiros disparos e contam como "evento".
    """
    enviados, falhas, ultimo = set(), set(), None
    if not Path(ARQUIVO_REGISTRO).exists():
        return enviados, falhas, ultimo
    mensagem_da_secao = "evento"
    for linha in Path(ARQUIVO_REGISTRO).read_text(encoding="utf-8").splitlines():
        if linha.startswith("## "):                        # começo de uma seção nova
            mensagem_da_secao = "evento"
        elif linha.startswith("- **Mensagem:**"):
            achou = re.search(r"\(`([^`]+)`\)", linha)
            mensagem_da_secao = achou.group(1) if achou else "evento"
        celulas = [c.strip() for c in linha.strip().strip("|").split("|")]
        if len(celulas) < 5 or not celulas[0].isdigit():   # ignora cabeçalho e testes
            continue
        posicao, nome, numero, status, horario = celulas[:5]
        status = status.split(" ", 1)[-1]                   # tira o ícone
        if status == "Enviado":
            if mensagem_da_secao == MENSAGEM_ATIVA:
                enviados.add(numero)
                ultimo = (int(posicao), nome, numero, horario)
        elif status not in ERROS_RETENTAVEIS:
            falhas.add(numero)
    return enviados, falhas, ultimo


def segundos_por_contato():
    return (TEMPO_BUSCAR + TEMPO_ANTES_ENTER + TEMPO_APOS_ENTER
            + sum(TEMPO_ENTRE_CONTATOS) / 2)


def formatar_duracao(segundos):
    minutos = int(segundos // 60)
    return f"{minutos // 60}h{minutos % 60:02d}min" if minutos >= 60 else f"{minutos}min"


def previsao_termino(segundos):
    return f"~{formatar_duracao(segundos)}, termina por volta das " \
           f"{datetime.fromtimestamp(time.time() + segundos):%H:%M}"


# --------------------------------- WhatsApp -------------------------------

def abrir_whatsapp():
    opcoes = webdriver.ChromeOptions()
    # Perfil próprio do robô: guarda o login do WhatsApp entre uma execução e outra
    opcoes.add_argument(f"--user-data-dir={PERFIL_CHROME}")
    # Mantém o WhatsApp funcionando mesmo com o Chrome minimizado ou atrás de outras janelas
    opcoes.add_argument("--disable-background-timer-throttling")
    opcoes.add_argument("--disable-backgrounding-occluded-windows")
    opcoes.add_argument("--disable-renderer-backgrounding")
    try:
        driver = webdriver.Chrome(options=opcoes)
    except WebDriverException:
        print("\n  ❌ Não consegui abrir o Chrome do robô.")
        print("     Se já existe uma janela do Chrome do Sender aberta, feche ela e tente de novo.")
        raise
    driver.maximize_window()
    driver.get("https://web.whatsapp.com")

    try:
        print(f"\n{LINHA}")
        print("  O WhatsApp já está logado nesse Chrome?")
        print("  (olhe o Chrome que abriu: se aparecem as suas conversas, está logado)")
        print("  [S] Sim, já está logado - pular o QR code")
        print("  [N] Não, vou escanear o QR code")
        logado = perguntar("  Escolha: ", ["s", "n"]) == "s"

        if logado:
            if whatsapp_carregou(driver):
                print("  ✅ Login confirmado, seguindo sem QR code.")
                return driver
            print("  ⚠️  Não encontrei as conversas na tela - o WhatsApp parece deslogado.")
            print("     Vamos para o QR code.")

        # Só segue quando as conversas aparecerem: disparar deslogado queimaria a lista
        while True:
            print(f"\n  Escaneie o QR code no Chrome que abriu.")
            print("  (segue sozinho quando as conversas aparecerem | ESC para parar)")
            esperar(TEMPO_LOGIN_QRCODE, "Aguardando login no QR code", pode_pular=True,
                    ate=lambda: bool(driver.find_elements(By.ID, "pane-side")))
            print("  Conferindo o login...")
            if whatsapp_carregou(driver):
                print("  ✅ Login confirmado.")
                return driver
            print("  ⚠️  As conversas ainda não apareceram - o WhatsApp não está logado.")
            print("     Vou esperar mais um pouco pelo QR code.")
    except ParadaEmergencia:
        driver.quit()             # sem login não há o que manter aberto
        raise


def whatsapp_carregou(driver, limite=20):
    """True se a lista de conversas aparecer (= logado) em até `limite` segundos."""
    fim = time.time() + limite
    while time.time() < fim:
        if driver.find_elements(By.ID, "pane-side"):
            return True
        time.sleep(1)
    return False


def manter_pc_acordado(ligar=True):
    """Impede o Windows de suspender enquanto o Sender roda (volta ao normal ao sair)."""
    import ctypes
    ES_CONTINUOUS, ES_SYSTEM_REQUIRED = 0x80000000, 0x00000001
    estado = ES_CONTINUOUS | (ES_SYSTEM_REQUIRED if ligar else 0)
    ctypes.windll.kernel32.SetThreadExecutionState(estado)


def esperar(segundos, etapa, pode_pular=False, pode_parar=True, ate=None):
    """Espera mostrando uma contagem regressiva com barra de progresso na mesma linha.

    ESC para o robô (ParadaEmergencia). Com pode_parar=False a tecla fica guardada
    e a parada acontece na próxima espera — usado logo após o Enter, para a
    mensagem que já saiu ser registrada no log antes de parar.
    `ate`: função checada 1x por segundo; se devolver True, a espera acaba antes.
    """
    # A linha que se atualiza precisa caber na janela e usar só caracteres simples:
    # se ela quebrar (ou tiver emoji/bloco de largura dupla), o \r não volta
    # ao começo e o cmd imprime uma linha nova a cada atualização.
    colunas = shutil.get_terminal_size().columns
    rotulo = etapa[:max(10, colunas - 30)]
    largura = max(5, min(20, colunas - len(rotulo) - 14))
    inicio = time.time()
    fim = inicio + segundos
    mostrado = None
    while (agora := time.time()) < fim:
        if pode_parar and msvcrt.kbhit():
            tecla = msvcrt.getwch()
            if tecla == "\x1b":
                print(f"\r  {rotulo}: ESC - parando...".ljust(colunas - 1), flush=True)
                raise ParadaEmergencia
            if pode_pular and tecla == "\r":
                break
        restante = math.ceil(fim - agora)
        if restante != mostrado:               # redesenha 1x por segundo
            if ate is not None and ate():
                break
            cheio = int((agora - inicio) / segundos * largura)
            barra = "#" * cheio + "-" * (largura - cheio)
            print(f"\r  {rotulo} [{barra}] {restante:>3}s ".ljust(colunas - 1),
                  end="", flush=True)
            mostrado = restante
        time.sleep(0.2)
    print(f"\r  {rotulo} [{'#' * largura}]  ok ".ljust(colunas - 1), flush=True)


def numero_invalido(driver):
    """True se o WhatsApp mostrou o aviso de número inválido (só o aviso, não a
    tela inteira — uma conversa com a palavra "inválido" não pode enganar)."""
    for aviso in driver.find_elements(By.CSS_SELECTOR, "div[role='dialog'], div[data-animate-modal-popup]"):
        texto = aviso.text.lower()
        if "inválido" in texto or "invalid" in texto:
            return True
    # plano B, caso o WhatsApp mude o layout do aviso: a frase exata dele
    # ("...compartilhado por url é inválido" / "...shared via url is invalid")
    texto = driver.find_element(By.TAG_NAME, "body").text.lower()
    return "url é inválido" in texto or "url is invalid" in texto


def caixa_de_mensagem(driver):
    try:
        return driver.find_element(By.CSS_SELECTOR, "footer div[contenteditable='true']")
    except NoSuchElementException:
        return None


def esperar_conversa(driver, limite):
    """Depois do tempo fixo, dá até `limite` segundos a mais para a internet lenta."""
    esperar(limite, "Internet lenta, esperando mais",
            ate=lambda: caixa_de_mensagem(driver) is not None or numero_invalido(driver))


def texto_da_caixa(driver):
    caixa = caixa_de_mensagem(driver)
    return caixa.text.strip() if caixa is not None else ""


def enviar(driver, posicao, nome, numero, status_ok="Enviado", tentativa=1):
    """Abre a conversa, envia, confere se saiu e registra. Retorna o status.

    ESC antes do Enter levanta ParadaEmergencia sem registrar nada: a mensagem
    ainda não saiu, então o contato pode ser retomado do zero com segurança.
    """
    try:
        driver.get(montar_link(numero, montar_mensagem(nome)))
        esperar(TEMPO_BUSCAR, "Carregando a conversa")
        if not caixa_de_mensagem(driver) and not numero_invalido(driver):
            esperar_conversa(driver, TEMPO_BUSCAR_EXTRA)

        if numero_invalido(driver):
            return registrar(posicao, nome, numero, "Número errado")
        if caixa_de_mensagem(driver) is None:
            return registrar(posicao, nome, numero, "Erro: conversa não carregou")

        esperar(TEMPO_ANTES_ENTER, "Aguardando para enviar")
        if not texto_da_caixa(driver):
            # às vezes o WhatsApp demora a colocar o texto na caixa
            esperar(TEMPO_TEXTO_EXTRA, "Esperando o texto aparecer",
                    ate=lambda: bool(texto_da_caixa(driver)))
        texto = texto_da_caixa(driver)
        if not texto:
            return registrar(posicao, nome, numero, "Erro: mensagem não apareceu")

        # Proteção contra mensagem dobrada: se já havia um rascunho nessa conversa,
        # a caixa pode ficar com o texto duas vezes. Limpa e recarrega uma vez.
        if len(texto) > len(montar_mensagem(nome)) + 20:
            if tentativa > 1:
                return registrar(posicao, nome, numero, "Erro: texto dobrado na caixa")
            print("  Havia um rascunho antigo na conversa - limpando e recarregando...")
            caixa = caixa_de_mensagem(driver)
            caixa.send_keys(Keys.CONTROL, "a")
            caixa.send_keys(Keys.DELETE)
            return enviar(driver, posicao, nome, numero, status_ok, tentativa + 1)

        caixa_de_mensagem(driver).send_keys(Keys.ENTER)
    except WebDriverException:
        return registrar(posicao, nome, numero, "Erro: Chrome não respondeu")

    # Daqui em diante o Enter já foi: nunca reenviar automaticamente.
    status = "Verificar no WhatsApp"
    try:
        esperar(TEMPO_APOS_ENTER, "Enviando a mensagem", pode_parar=False)
        caixa = caixa_de_mensagem(driver)
        if caixa is not None and not caixa.text.strip():   # caixa esvaziou = saiu
            status = status_ok
    except WebDriverException:
        pass
    finally:
        registrar(posicao, nome, numero, status)          # registra mesmo com Ctrl+C aqui
    return status


def contatos_da_planilha(df, enviados, falhas, quantidade):
    """Os próximos `quantidade` contatos pendentes, na ordem da planilha.

    Quem já recebeu ou já falhou fica de fora. Números inválidos novos que
    aparecem no caminho entram (só para irem pro registro, não gastam tempo)
    e não contam na quantidade.
    """
    selecionados, validos, vistos = [], 0, set()
    # start=2: mesma numeração de linha do Excel (a linha 1 é o cabeçalho)
    for posicao, (_, linha) in enumerate(df.iterrows(), start=2):
        if validos >= quantidade:
            break
        bruto = linha[COLUNA_NUMERO]
        numero = formatar_numero(bruto)
        if numero is None:
            if str(bruto).strip() not in falhas:
                selecionados.append((posicao, nome_da_planilha(linha[COLUNA_NOME]), bruto))
            continue
        if numero in enviados or numero in falhas or numero in vistos:   # vistos: repetido na planilha
            continue
        if bloqueado(numero):
            continue
        vistos.add(numero)
        selecionados.append((posicao, nome_da_planilha(linha[COLUNA_NOME]), bruto))
        validos += 1
    return selecionados


def disparar(driver, contatos, ja_enviados, titulo="Disparando...", status_ok="Enviado",
             tipo="Disparo oficial"):
    """Mesmo fluxo para o teste e para o oficial: aviso de ESC, progresso,
    contagens, pausa entre contatos e resumo no final (terminal e Markdown)."""
    print(f"\n{LINHA}")
    print(f"  {titulo}")
    print("  ⏸  ESC a qualquer momento: pausa e você escolhe continuar, pular ou parar")
    print("     (com a janela do cmd selecionada)")
    print(LINHA)

    pendentes = sum(1 for _, _, b in contatos
                    if (n := formatar_numero(b)) and n not in ja_enviados)
    inicio = time.time()
    iniciar_secao(tipo, pendentes)
    try:
        _disparar(driver, contatos, ja_enviados, status_ok, pendentes, inicio)
    except (ParadaEmergencia, KeyboardInterrupt, WebDriverException):
        fechar_secao(inicio, interrompido=True)
        raise
    fechar_secao(inicio, interrompido=False)
    print("\n  ✅ Envio concluído!")
    mostrar_resumo()


def menu_pausa(contato=None):
    """Menu do ESC. `contato` = (nome, número) de quem seria o próximo a receber.
    Retorna 'c' (continuar), 's' (pular esse contato) ou 'p' (parar)."""
    print(f"\n{LINHA}")
    print("  ⏸  PAUSADO - nada será enviado enquanto você decide.")
    opcoes = ["c", "p"]
    if contato:
        print(f"  Próximo a receber: {contato[0]} - {contato[1]}")
    print(LINHA)
    print("  [C] Continuar os envios")
    if contato:
        print(f"  [S] Pular {contato[0]} (não recebe, nem nos próximos disparos)")
        opcoes.insert(1, "s")
    print("  [P] Parar os envios")
    return perguntar("  Escolha: ", opcoes)


def _disparar(driver, contatos, ja_enviados, status_ok, pendentes, inicio):
    feitos = 0
    erros_seguidos = 0
    pular_proximo = False
    for posicao, nome, bruto in contatos:
        numero = formatar_numero(bruto)

        if numero is None:
            registrar(posicao, nome, bruto, "Número fora do padrão")
            continue
        if numero in ja_enviados:
            print(f"  {posicao}°, {nome}, {numero}, já enviado - pulando")
            continue

        # Estimativa pelos timers no início; depois, pela média real dos envios
        media = (time.time() - inicio) / feitos if feitos else segundos_por_contato()
        linha_planilha = f"   (linha {posicao} da planilha)" if isinstance(posicao, int) else ""
        print(f"\n  [{feitos + 1}/{pendentes}] {nome} - {numero}{linha_planilha}")
        if pendentes > 1:
            print(f"  ⏱  Falta {previsao_termino(media * (pendentes - feitos))}")

        if pular_proximo:
            status = registrar(posicao, nome, numero, "Pulado por você")
            pular_proximo = False
        else:
            while True:
                try:
                    status = enviar(driver, posicao, nome, numero, status_ok)
                    break
                except ParadaEmergencia:          # ESC antes do Enter: nada saiu ainda
                    escolha = menu_pausa((nome, numero))
                    if escolha == "p":
                        raise
                    if escolha == "s":
                        status = registrar(posicao, nome, numero, "Pulado por você")
                        break
                    print(f"  ▶  Retomando {nome} do começo (a mensagem ainda não tinha saído).")
        feitos += 1

        # Freio: vários erros seguidos = problema no WhatsApp/internet, não no contato
        erros_seguidos = erros_seguidos + 1 if status.startswith(("Erro", "Verificar")) else 0
        if erros_seguidos >= LIMITE_ERROS_SEGUIDOS and feitos < pendentes:
            print(f"\n{LINHA}")
            print(f"  ⚠️  {erros_seguidos} erros seguidos - o WhatsApp pode ter deslogado")
            print("     ou a internet caiu. Confira o Chrome do robô antes de seguir.")
            print(LINHA)
            if perguntar("  [C] Continuar o disparo   [P] Parar: ", ["c", "p"]) == "p":
                raise ParadaEmergencia
            erros_seguidos = 0

        if feitos < pendentes:
            proximo = proximo_contato(contatos, posicao, ja_enviados)
            if proximo:
                print(f"  Próximo: {proximo[0]} - {proximo[1]}  (ESC para pausar)")
            try:
                esperar(random.randint(*TEMPO_ENTRE_CONTATOS), "Pausa até o próximo contato")
            except ParadaEmergencia:
                escolha = menu_pausa(proximo)
                if escolha == "p":
                    raise
                pular_proximo = escolha == "s"
                print("  ▶  Retomando os envios.")


def proximo_contato(contatos, posicao_atual, ja_enviados):
    """(nome, número) do próximo contato que vai de fato receber mensagem."""
    passou = False
    for posicao, nome, bruto in contatos:
        if passou:
            numero = formatar_numero(bruto)
            if numero and numero not in ja_enviados:
                return nome, numero
        passou = passou or posicao == posicao_atual
    return None


# -------------------------------- interface -------------------------------

def perguntar(texto, opcoes):
    while True:
        resposta = input(texto).strip().lower()
        if resposta in opcoes:
            return resposta
        print(f"  Opção inválida. Digite: {' / '.join(opcoes)}")


def pedir_contato_teste():
    print(f"\n{LINHA}")
    print("  TESTE - para quem enviar?")
    print(LINHA)
    print("  Vou pedir uma coisa de cada vez. Digite e aperte Enter.")
    print(LINHA)

    print("\n  1/2 - Digite o NOME e aperte Enter.")
    print("        (só o primeiro nome é usado. Ex.: você)")
    nome = ""
    while not nome:
        nome = primeiro_nome(input("  Nome: "))

    print("\n  2/2 - Agora digite o NÚMERO e aperte Enter.")
    print("        (só números, com DDD, sem o 55, sem espaço,")
    print("         parênteses ou traço. Ex.: 15991234567)")
    while True:
        numero = formatar_numero(input("  Número: "))
        if numero:
            return nome, numero
        print("  Número inválido. Use DDD + número, ex.: 15991234567")


def resumo_planilha():
    atualizada = datetime.fromtimestamp(os.path.getmtime(PLANILHA))
    df = pd.read_excel(PLANILHA, sheet_name=ABA, dtype=str)
    numeros = df[COLUNA_NUMERO].map(formatar_numero)
    enviados, falhas, ultimo = ler_historico() if PULAR_JA_ENVIADOS else (set(), set(), None)
    validos = numeros.notna().sum()
    unicos = {n for n in numeros.dropna() if not bloqueado(n)}   # repetido conta 1x
    ja_enviados = len(unicos & enviados)
    ja_falharam = len((unicos & falhas) - enviados)
    pendentes = len(unicos - enviados - falhas)

    print(f"  💬 Mensagem ativa: {nome_mensagem()}")
    print(f"  Planilha: {PLANILHA}")
    print(f"  Última atualização: {atualizada:%d/%m/%Y às %H:%M}")
    print(f"  Contatos na planilha: {len(df)}")
    print(f"    - números válidos:   {validos}")
    print(f"    - números inválidos: {len(df) - validos}")
    if PULAR_JA_ENVIADOS:
        print(f"    - já receberam esta mensagem: {ja_enviados}")
        if ja_falharam:
            print(f"    - falharam antes:    {ja_falharam}  (não são tentados de novo)")
        print(f"    - pendentes:         {pendentes}")
    if ultimo:
        linha, nome, numero, horario = ultimo
        print(f"  Último enviado: linha {linha} - {nome} ({numero}) em {horario}")
    print(f"  Tempo para enviar a todos os pendentes: "
          f"~{formatar_duracao(pendentes * segundos_por_contato())}")
    return df, (enviados, falhas), pendentes


def escolher_quantidade(pendentes, df, enviados, falhas):
    """Pergunta quantos leads disparar, mostra de onde até onde vai e confirma.
    Retorna a lista de contatos selecionados ou None."""
    if pendentes == 0:
        print("\n  Não há leads pendentes: todos da planilha já receberam.")
        input("  Aperte Enter para voltar ao menu inicial...")
        return None

    print(f"\n{LINHA}")
    print("  Quantos contatos vamos disparar agora?")
    print(f"  Pendentes: {pendentes}  |  cada contato leva ~{segundos_por_contato():.0f}s")
    print(f"  Todos os {pendentes}: {previsao_termino(pendentes * segundos_por_contato())}")
    print(LINHA)
    while True:
        resposta = input(f"  Quantidade (1 a {pendentes}, Enter = todos, 0 = voltar): ").strip()
        if not resposta:
            quantidade = pendentes
            break
        if resposta.isdigit() and 0 <= int(resposta) <= pendentes:
            quantidade = int(resposta)
            break
        print(f"  Digite um número de 1 a {pendentes}.")
    if quantidade == 0:
        return None

    contatos = contatos_da_planilha(df, enviados, falhas, quantidade)
    validos = [c for c in contatos if formatar_numero(c[2])]
    primeiro, ultimo = validos[0], validos[-1]
    print(f"\n  💬 Mensagem: {nome_mensagem()}")
    print(f"  Vou disparar para os próximos {quantidade} leads pendentes da planilha:")
    print(f"    começa na linha {primeiro[0]} - {primeiro[1]}")
    print(f"    termina na linha {ultimo[0]} - {ultimo[1]}")
    print(f"  Tempo estimado: {previsao_termino(quantidade * segundos_por_contato())}.")
    confirma = perguntar("  Confirma? Digite SIM para começar ou NAO para voltar: ", ["sim", "nao"])
    return contatos if confirma == "sim" else None


def chrome_aberto(driver):
    if driver is None:
        return False
    try:
        driver.title          # falha se o Chrome foi fechado na mão
        return True
    except Exception:
        return False


def menu_inicial(chrome_ligado):
    os.system("cls")
    print(LINHA)
    print("  Oi! Eu sou o Sender 👋")
    print(LINHA)
    df, log, pendentes = resumo_planilha()

    print(f"\n{LINHA}")
    print("  O que vamos fazer?")
    print("  [1] Fazer um teste antes (envia a mensagem para um número seu)")
    print("  [2] Disparar para os leads da planilha (você escolhe quantos)")
    print(f"  [3] Trocar a mensagem (ativa: {nome_mensagem()})")
    print("  [0] Sair" + (" e fechar o Chrome" if chrome_ligado else ""))
    escolha = perguntar("  Escolha: ", ["1", "2", "3", "0"])
    return escolha, df, log, pendentes


def trocar_mensagem():
    """Mostra as mensagens disponíveis com uma prévia e troca a ativa."""
    global MENSAGEM_ATIVA
    chaves = list(MENSAGENS)
    print(f"\n{LINHA}")
    print("  Qual mensagem vamos usar?")
    print(LINHA)
    for i, chave in enumerate(chaves, start=1):
        ativa = "  ← ativa" if chave == MENSAGEM_ATIVA else ""
        print(f"\n  [{i}] {nome_mensagem(chave)}{ativa}")
        for linha in MENSAGENS[chave]["texto"].replace("{Nome}", "Fulano").splitlines():
            print(f"      │ {linha}")
    print("\n  [0] Voltar sem trocar")
    escolha = perguntar("  Escolha: ", [str(i) for i in range(len(chaves) + 1)])
    if escolha != "0":
        MENSAGEM_ATIVA = chaves[int(escolha) - 1]
        print(f"\n  ✅ Mensagem ativa agora: {nome_mensagem()}")
        print("     Dica: faça um teste [1] para ver no celular antes de disparar.")
        input("  Aperte Enter para voltar ao menu...")


def main():
    driver = None
    manter_pc_acordado(True)
    try:
        while True:
            escolha, df, (enviados, falhas), pendentes = menu_inicial(chrome_aberto(driver))
            if escolha == "0":
                break
            if escolha == "3":
                trocar_mensagem()
                continue
            contatos = None
            if escolha == "2":
                contatos = escolher_quantidade(pendentes, df, enviados, falhas)
                if contatos is None:
                    continue

            teste = pedir_contato_teste() if escolha == "1" else None

            try:
                if not chrome_aberto(driver):
                    if driver is not None:           # sobra de um Chrome que travou
                        try:
                            driver.quit()
                        except WebDriverException:
                            pass
                    driver = None
                    driver = abrir_whatsapp()

                while teste:
                    try:
                        disparar(driver, [("Teste", *teste)], set(),
                                 titulo="TESTE - mesmo fluxo do disparo oficial",
                                 status_ok="Teste enviado", tipo="Teste")
                        print("\n  Confira no celular se a mensagem chegou certa (nome, emojis, link).")
                    except ParadaEmergencia:
                        aviso_parada()
                        print("  ✅ A parada de emergência funcionou. No oficial será igual.")
                    print("  [1] Fazer outro teste")
                    print("  [2] Está tudo certo, disparar para os leads")
                    print("  [0] Voltar ao menu inicial (o Chrome continua aberto)")
                    escolha = perguntar("  Escolha: ", ["1", "2", "0"])
                    if escolha == "1":
                        teste = pedir_contato_teste()
                    elif escolha == "2" and (
                            contatos := escolher_quantidade(pendentes, df, enviados, falhas)):
                        teste = None
                    else:
                        break
                else:
                    disparar(driver, contatos, enviados)
                    input("\n  Aperte Enter para voltar ao menu inicial...")

            except ParadaEmergencia:
                aviso_parada()
                input("\n  Aperte Enter para voltar ao menu inicial...")
            except WebDriverException as erro:
                # Chrome fechado/travado no meio: não derruba o Sender, volta pro menu
                print(f"\n{LINHA}")
                print("  ❌ O Chrome do robô parou de responder (foi fechado ou travou).")
                print(f"     Detalhe: {str(erro).splitlines()[0][:100]}")
                print("  Nada mais será enviado. No próximo disparo ele abre o Chrome de novo")
                print("  e continua de onde parou.")
                print(LINHA)
                mostrar_resumo()
                input("\n  Aperte Enter para voltar ao menu inicial...")

    except KeyboardInterrupt:
        print("\n\n  🛑 Ctrl+C - encerrando o Sender.")
    finally:
        if driver is not None:
            try:
                driver.quit()
            except WebDriverException:
                pass
        manter_pc_acordado(False)
        mostrar_resumo()
        print("  Até mais!")


def aviso_parada():
    print(f"\n{LINHA}")
    print("  🛑 ENVIOS PARADOS - nada mais será enviado.")
    print("  O Chrome continua aberto. Quem já recebeu é pulado no próximo disparo.")
    print(LINHA)
    mostrar_resumo()


def mostrar_resumo():
    """Mostra o que foi feito desde o último resumo e zera a contagem."""
    if not RESUMO:
        return
    print("\n  Resumo:")
    for status, qtd in RESUMO.items():
        print(f"    - {status}: {qtd}")
    print(f"  Registro completo na pasta de dados: Sender/{Path(ARQUIVO_REGISTRO).name}")
    RESUMO.clear()


if __name__ == "__main__":
    main()
