"""
Sender - automação de disparo de WhatsApp a partir de uma planilha Excel.

Fluxo:
  1. Mostra a campanha ativa (planilha + mensagem + regras) e quantos pendentes há.
  2. Teste (envia para um número seu) ou disparo (você escolhe quantos contatos).
  3. Abre o WhatsApp Web e espera o login pelo QR code.
  4. Para cada contato: abre a conversa já com a mensagem, envia (texto ou
     foto com legenda), confere se saiu e grava o resultado no Registros.md.

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
import unicodedata
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
# Ajuste estes valores para o seu caso. Por padrão tudo fica em "dados/", ao lado
# do script; aponte para qualquer outra pasta se preferir (ex.: um cofre do Obsidian).

PASTA = Path(__file__).resolve().parent

# Pasta com as planilhas (o menu "Personalizada" lista tudo o que está aqui)
PASTA_PLANILHAS = str(PASTA / "dados")

# Arquivo (editável) com as campanhas e mensagens. Se existir, ele MANDA:
# o que estiver nele substitui as campanhas e mensagens escritas aqui embaixo.
# O menu [3] relê o arquivo, então dá para criar campanha com o Sender aberto.
ARQUIVO_CAMPANHAS = str(Path(PASTA_PLANILHAS) / "Campanhas.md")

# Planilha usada se nenhuma campanha for ativada (as campanhas definem a sua)
PLANILHA = str(Path(PASTA_PLANILHAS) / "contatos.xlsx")
ABA = 0                               # aba lida (0 = primeira)
LINHA_CABECALHO = 0                   # linha do cabeçalho (0 = primeira). Detectada sozinha nas campanhas
COLUNA_NOME = "Nome"                  # cabeçalho da coluna do nome usado na mensagem
COLUNA_NUMERO = "Telefone"            # cabeçalho da coluna do telefone
# Coluna opcional de controle manual (ex.: "Mensagem enviada?"): só recebe quem está
# com "Não"; após cada envio confirmado vira "Sim" e a planilha é salva.
COLUNA_JA_ENVIADA = None

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
TEMPO_TEXTO_EXTRA = 10                # se a mensagem não apareceu na caixa, espera até mais isso
TEMPO_ANTES_ENTER = 5                 # espera antes de enviar
TEMPO_APOS_ENTER = 5                  # garante que a mensagem saiu antes de trocar de página
TEMPO_ENTRE_CONTATOS = (20, 60)       # aleatório entre min e máx, antes do próximo contato

LIMITE_ERROS_SEGUIDOS = 3             # pausa o disparo e pergunta se deve continuar

# Mensagens disponíveis. {Nome} vira o primeiro nome da coluna de nome.
# Com "imagem", a mensagem vai como FOTO + LEGENDA (uma mensagem só). Use JPG/PNG.
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
    "convite_imagem": {
        "nome": "Convite com imagem (exemplo)",
        "imagem": str(PASTA / "dados" / "imagem-exemplo.jpg"),
        "texto": """Oi, {Nome}! Tudo bem?

Preparamos uma novidade para você: o site oficial do evento já está no ar.

https://exemplo.com.br/evento

Nos vemos lá! 🚀""",
    },
}
MENSAGEM_ATIVA = "convite"

# Campanhas = planilha + mensagem + regras. É isso que aparece no menu do Sender.
# "planilha" é o nome do arquivo dentro de PASTA_PLANILHAS (cabeçalho e colunas são
# detectados sozinhos). "usar_controle": só recebe quem está "Não" na coluna de
# controle (ex.: "Mensagem enviada?") e, após o envio, a linha vira "Sim".
CAMPANHAS = {
    # "excluir": planilhas (na mesma pasta) cujos números NUNCA recebem a campanha.
    # Ex.: "excluir": ["clientes.xlsx"] para não mandar convite a quem já comprou.
    "convite": {"nome": "Convite", "publico": "lista de contatos",
                "planilha": "contatos.xlsx", "mensagem": "convite"},
    "convite_imagem": {"nome": "Convite com imagem", "publico": "lista de contatos",
                       "planilha": "contatos.xlsx", "mensagem": "convite_imagem", "usar_controle": True},
}
CAMPANHA_INICIAL = "convite"          # campanha que já vem escolhida ao abrir o Sender

# O wa.me abre uma página intermediária ("Continuar para a conversa") no
# computador; o web.whatsapp.com/send vai direto para a conversa no WhatsApp Web.
LINK_BASE = "https://web.whatsapp.com/send?phone={numero}&text={texto}"
LINK_CONVERSA = "https://web.whatsapp.com/send?phone={numero}"   # conversa sem texto

# ==========================================================================

LINHA = "=" * 60


# ----------------------------- dados / mensagem ---------------------------

def ler_planilha(caminho=None, cabecalho=None):
    """Lê a planilha ativa (ou outra) a partir da linha de cabeçalho, como texto."""
    df = pd.read_excel(caminho or PLANILHA, sheet_name=ABA,
                       header=LINHA_CABECALHO if cabecalho is None else cabecalho, dtype=str)
    return df.dropna(how="all")          # linhas em branco saem; o índice original fica


def planilhas_da_pasta():
    """Todas as planilhas de PASTA_PLANILHAS e suas subpastas (ex.: Campanhas/),
    da mais recente para a mais antiga. Ignora backups e temporários do Excel."""
    pasta = Path(PASTA_PLANILHAS)
    arquivos = [a for a in pasta.rglob("*.xls*")
                if "backup" not in a.name.lower() and not a.name.startswith("~$")]
    return sorted(arquivos, key=lambda a: a.stat().st_mtime, reverse=True)


def caminho_planilha(nome):
    """Nome do arquivo (ou caminho relativo) -> caminho completo. Procura nas subpastas."""
    if Path(nome).is_absolute():
        return Path(nome)
    direto = Path(PASTA_PLANILHAS) / nome
    if direto.exists():
        return direto
    for arquivo in planilhas_da_pasta():
        if arquivo.name.lower() == Path(nome).name.lower():
            return arquivo
    return direto                      # não achou: devolve o caminho esperado (vira aviso)


def linha_excel(indice):
    """Índice do pandas -> número da linha no Excel (considerando o cabeçalho)."""
    return int(indice) + LINHA_CABECALHO + 2


# Palavras usadas para reconhecer as colunas, em ordem de preferência
PISTAS_TELEFONE = ["whatsapp", "whats", "celular", "telefone", "fone", "atualiza", "phone",
                   "numero", "num.", "contato", "cel"]


def sem_acento(texto):
    return "".join(c for c in unicodedata.normalize("NFKD", str(texto))
                   if not unicodedata.combining(c)).lower()


def detectar_planilha(caminho):
    """Acha a linha de cabeçalho e as colunas de nome e telefone de uma planilha.
    Retorna (linha_cabecalho, coluna_nome, coluna_numero, colunas) ou None."""
    bruto = pd.read_excel(caminho, sheet_name=ABA, header=None, dtype=str, nrows=40)
    for i in range(len(bruto)):
        colunas = [str(c).strip() for c in bruto.iloc[i].tolist() if isinstance(c, str) and c.strip()]
        baixas = [c.lower() for c in colunas]
        baixas = [sem_acento(c) for c in baixas]
        tem_nome = any("nome" in c for c in baixas)
        tem_tel = any(p in c for c in baixas for p in PISTAS_TELEFONE)
        if tem_nome and tem_tel:
            df = ler_planilha(caminho, cabecalho=i)
            return (i, escolher_coluna_nome(df), escolher_coluna_telefone(df), list(df.columns))
    return None


def escolher_coluna_controle(df):
    """Coluna de controle manual, tipo 'Mensagem enviada?' / 'Enviado' (ou None)."""
    for c in df.columns:
        if "enviad" in str(c).lower():
            return str(c)
    return None


def marcado_como_enviado(linha):
    """Com coluna de controle ativa, só recebe quem está marcado "Não".
    True = pular esta linha ("Sim", vazio ou qualquer outra coisa)."""
    if not COLUNA_JA_ENVIADA or COLUNA_JA_ENVIADA not in linha:
        return False
    valor = linha[COLUNA_JA_ENVIADA]
    return pd.isna(valor) or str(valor).strip().lower() not in {"não", "nao", "n"}


def marcar_sim_na_planilha(linha_do_excel):
    """Depois de um envio confirmado, troca a coluna de controle para "Sim" e salva.
    Se a planilha estiver aberta no Excel, avisa e segue (o Registros.md garante
    que a pessoa não recebe de novo)."""
    if not COLUNA_JA_ENVIADA or not isinstance(linha_do_excel, int):
        return
    import openpyxl
    try:
        livro = openpyxl.load_workbook(PLANILHA)
        aba = livro.worksheets[ABA]
        cabecalho = [str(c.value).strip() if c.value is not None else "" for c in aba[LINHA_CABECALHO + 1]]
        coluna = cabecalho.index(COLUNA_JA_ENVIADA) + 1
        aba.cell(row=linha_do_excel, column=coluna, value="Sim")
        livro.save(PLANILHA)
        print(f"  📝 Planilha: linha {linha_do_excel} marcada como \"Sim\"")
    except PermissionError:
        print("  ⚠️  Não consegui salvar \"Sim\" na planilha (está aberta no Excel?). "
              "O envio está no registro e não se repete.")
    except (ValueError, OSError) as erro:
        print(f"  ⚠️  Não consegui marcar \"Sim\" na planilha: {erro}")


def escolher_coluna_nome(df):
    colunas = [str(c) for c in df.columns]
    for c in colunas:                                   # 1º: exatamente "Nome"
        if c.strip().lower() == "nome":
            return c
    for c in colunas:                                   # 2º: contém "nome", menos sobrenome
        if "nome" in c.lower() and "sobrenome" not in c.lower():
            return c
    return None


def escolher_coluna_telefone(df):
    """Entre as colunas com cara de telefone, fica com a que tem mais números válidos."""
    candidatas = [str(c) for c in df.columns
                  if any(p in sem_acento(c) for p in PISTAS_TELEFONE)]
    if not candidatas:
        return None
    return max(candidatas, key=lambda c: df[c].map(formatar_numero).notna().sum())


def nome_da_planilha(valor):
    """Nome da planilha. Nome já tratado ('Maria') vai como está; nome cru
    ('MARIA APARECIDA', 'joão silva') vira só o primeiro nome ('Maria', 'João')."""
    if pd.isna(valor) or not str(valor).strip():
        return ""
    texto = str(valor).strip()
    if " " in texto or texto.islower() or texto.isupper():
        return primeiro_nome(texto)
    return texto


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
    digitos = re.sub(r"\D", "", str(valor).removesuffix(".0")).lstrip("0")   # 011... -> 11...
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
    if nome:
        return texto.replace("{Nome}", nome)
    return texto.replace(", {Nome}", "").replace(" {Nome}", "").replace("{Nome}", "")


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

# --------------------------- campanhas do arquivo --------------------------
#
# Formato do Campanhas.md (uma campanha por bloco "## "):
#
#   ## Convite para o evento
#   id: evento                  <- opcional; guarda a memória de quem já recebeu
#   publico: leads
#   planilha: contatos.xlsx
#   imagem: banner.jpg          <- opcional (foto com legenda)
#   excluir: clientes.xlsx  <- opcional, separe por vírgula
#   controle: sim               <- opcional: só envia para quem está "Não" na coluna de controle
#   segmentar: É membro?        <- opcional: o Sender pergunta para qual público enviar
#   mensagem:
#   Olá {Nome}! Tudo bem?
#   ...

CAMPOS_CAMPANHA = {"id", "publico", "planilha", "imagem", "excluir", "controle", "segmentar", "mensagem"}


def _slug(texto):
    limpo = re.sub(r"[^a-z0-9]+", "_", texto.lower().strip())
    return limpo.strip("_") or "campanha"


def ler_arquivo_campanhas(caminho=None):
    """Lê o Campanhas.md. Devolve (mensagens, campanhas, erros)."""
    arquivo = Path(caminho or ARQUIVO_CAMPANHAS)
    if not arquivo.exists():
        return {}, {}, []
    mensagens, campanhas, erros = {}, {}, []
    blocos = re.split(r"^##\s+", arquivo.read_text(encoding="utf-8"), flags=re.M)[1:]
    for bloco in blocos:
        linhas = bloco.splitlines()
        nome = linhas[0].strip()
        dados, texto, lendo_texto = {}, [], False
        for linha in linhas[1:]:
            if lendo_texto:
                if linha.strip() in {"---", "***", "___"}:   # separador entre blocos
                    break
                texto.append(linha)
                continue
            achou = re.match(r"\s*([a-zA-Zçã]+)\s*:\s*(.*)$", linha)
            if achou and achou.group(1).lower() in CAMPOS_CAMPANHA:
                campo, valor = achou.group(1).lower(), achou.group(2).strip()
                if campo == "mensagem":
                    lendo_texto = True
                    if valor:
                        texto.append(valor)
                else:
                    dados[campo] = valor
            elif linha.strip():
                erros.append(f"{nome}: linha ignorada -> {linha.strip()[:40]}")
        texto = "\n".join(texto).strip("\n")
        if not dados.get("planilha"):
            erros.append(f"{nome}: falta a linha \"planilha:\"")
            continue
        if not texto:
            erros.append(f"{nome}: falta a \"mensagem:\"")
            continue
        # "id" identifica a MENSAGEM (é o que guarda quem já recebeu); a campanha
        # é identificada por nome + público, então duas campanhas podem usar a
        # mesma mensagem para públicos diferentes.
        msg_id = dados.get("id") or _slug(nome)
        publico = dados.get("publico", "")
        chave = _slug(f"{nome} {publico}")
        while chave in campanhas:
            chave += "_2"
        imagem = dados.get("imagem", "")
        mensagens[msg_id] = {"nome": nome, "texto": texto}
        if imagem:
            mensagens[msg_id]["imagem"] = imagem if Path(imagem).is_absolute() else str(caminho_imagem(imagem))
        campanhas[chave] = {"nome": nome, "publico": publico,
                            "planilha": dados["planilha"], "mensagem": msg_id}
        if dados.get("controle", "").lower() in {"sim", "s", "1", "true"}:
            campanhas[chave]["usar_controle"] = True
        if dados.get("excluir"):
            campanhas[chave]["excluir"] = [x.strip() for x in dados["excluir"].split(",") if x.strip()]
        if dados.get("segmentar"):
            campanhas[chave]["segmentar"] = dados["segmentar"]
    return mensagens, campanhas, erros


def caminho_imagem(nome):
    """Imagem na pasta das planilhas ou em qualquer subpasta dela."""
    direto = Path(PASTA_PLANILHAS) / nome
    if direto.exists():
        return direto
    for achado in Path(PASTA_PLANILHAS).rglob(Path(nome).name):
        return achado
    return direto


def carregar_campanhas(avisar=True):
    """Se o Campanhas.md existir, ele substitui as campanhas/mensagens do código."""
    global MENSAGENS, CAMPANHAS
    mensagens, campanhas, erros = ler_arquivo_campanhas()
    if avisar:
        for erro in erros:
            print(f"  ⚠️  Campanhas.md - {erro}")
    if campanhas:
        MENSAGENS = mensagens
        CAMPANHAS = campanhas
    return bool(campanhas)


# --------------------------- registro no Markdown --------------------------
#
# Registros.md fica assim (uma seção por disparo, mais recente embaixo):
#
#   ## 🚀 Disparo oficial — 18/09/2026 às 18:30
#   | Nº | Nome | Número | Status | Horário |
#   |---:|---|---|---|---|
#   | 2 | Fulano | 5511900001111 | ✅ Enviado | 18:31:05 |
#   > **Resumo:** ✅ 1 enviado · ⏱ 1min · finalizado às 18:32

ICONES = {"Enviado": "✅", "Teste enviado": "🧪", "Número errado": "❌",
          "Número fora do padrão": "❌"}

# Falhas que aconteceram ANTES do Enter: a mensagem com certeza não saiu, então
# o contato volta a ser tentado no próximo disparo. Todo o resto é definitivo.
ERROS_RETENTAVEIS = {"Erro: conversa não carregou", "Erro: mensagem não apareceu",
                     "Erro: Chrome não respondeu", "Erro: texto dobrado na caixa",
                     "Erro: botão de anexo não encontrado", "Erro: imagem não abriu",
                     "Erro: legenda não apareceu", "Erro: opção Fotos e vídeos não encontrada"}
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
        + (f"- **Público:** {SEGMENTAR_COLUNA} = {SEGMENTO if SEGMENTO is not None else 'todos'}\n"
           if SEGMENTAR_COLUNA else "")
        + f"- **Contatos selecionados:** {quantidade}\n\n"
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
        print("  [1] Sim, já está logado - pular o QR code")
        print("  [0] Não, vou escanear o QR code")
        logado = perguntar("  Escolha: ", ["1", "0"]) == "1"

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
    """Caixa de texto da conversa (no rodapé; com prévia de link o layout muda)."""
    for seletor in ("footer div[contenteditable='true']",
                    "div[contenteditable='true'][aria-label^='Digite uma mensagem']",
                    "div[contenteditable='true'][aria-placeholder^='Digite uma mensagem']",
                    "div[contenteditable='true'][aria-label^='Type a message']"):
        achados = driver.find_elements(By.CSS_SELECTOR, seletor)
        if achados:
            return achados[0]
    return None


def esperar_conversa(driver, limite):
    """Depois do tempo fixo, dá até `limite` segundos a mais para a internet lenta."""
    esperar(limite, "Internet lenta, esperando mais",
            ate=lambda: caixa_de_mensagem(driver) is not None or numero_invalido(driver))


def texto_da_caixa(driver):
    caixa = caixa_de_mensagem(driver)
    return caixa.text.strip() if caixa is not None else ""


# ------------------------------ envio com imagem ---------------------------
#
# Lições do WhatsApp Web real:
# - O menu de anexo tem "Fotos e vídeos" E "Nova figurinha". O campo de arquivo que
#   já existe na página é o da FIGURINHA (accept="image/*") -> a imagem virava figurinha.
#   O certo é clicar em "Fotos e vídeos" (campo com accept "image/*,video/...").
# - Clicar em "Fotos e vídeos" abriria a janela de arquivos do Windows; o Sender
#   intercepta esse clique e entrega o arquivo direto ao WhatsApp.
# - O ESC não fecha a prévia: é preciso clicar em "Fechar".

SELETORES_ANEXAR = ["[aria-label='Anexar']", "[title='Anexar']", "[aria-label='Attach']",
                    "[title='Attach']", "span[data-icon='plus-rounded']", "span[data-icon='plus']",
                    "span[data-icon='attach-menu-plus']", "span[data-icon='clip']"]

# Intercepta o clique que abriria a janela de arquivos e guarda o campo de arquivo
JS_INTERCEPTAR_ARQUIVO = """
window.__inputFoto = null;
if (!window.__senderPatch) {
  const guardar = el => { window.__inputFoto = el;
    if (!el.isConnected) { el.style.display = 'none'; document.body.appendChild(el); } };
  const clickOrig = HTMLInputElement.prototype.click;
  HTMLInputElement.prototype.click = function () {
    if (this.type === 'file') { guardar(this); return; }
    return clickOrig.apply(this, arguments);
  };
  if (HTMLInputElement.prototype.showPicker) {
    const pickOrig = HTMLInputElement.prototype.showPicker;
    HTMLInputElement.prototype.showPicker = function () {
      if (this.type === 'file') { guardar(this); return; }
      return pickOrig.apply(this, arguments);
    };
  }
  window.__senderPatch = true;
}
"""

# Item "Fotos e vídeos" do menu de anexo
JS_ITEM_FOTOS = """
const alvo = [...document.querySelectorAll('li, [role="button"], [role="menuitem"], button, div, span')]
  .filter(e => e.offsetParent !== null && /^(fotos e vídeos|fotos e videos|photos & videos|photos and videos)$/i
               .test((e.innerText || e.getAttribute('aria-label') || '').trim()));
if (!alvo.length) return null;
const e = alvo[alvo.length - 1];
return e.closest('li, [role="button"], [role="menuitem"], button') || e;
"""

# Caixa de legenda da prévia da foto (fora do rodapé, da busca e da caixa da conversa)
JS_LEGENDA = """
return [...document.querySelectorAll('div[contenteditable="true"]')]
  .find(e => e.offsetParent !== null && !e.closest('footer') && !e.closest('#side')
             && !/^(digite uma mensagem|type a message)/i.test(e.getAttribute('aria-label') || '')) || null;
"""

# Cola o texto como se fosse Ctrl+V, sem usar a área de transferência do Windows
JS_COLAR = """
const alvo = arguments[0], texto = arguments[1];
alvo.focus();
const dados = new DataTransfer();
dados.setData('text/plain', texto);
alvo.dispatchEvent(new ClipboardEvent('paste', {clipboardData: dados, bubbles: true, cancelable: true}));
"""

# Botão de enviar da prévia ("Enviar 1 item selecionado"), fora do rodapé
JS_BOTAO_ENVIAR = """
const cands = [...document.querySelectorAll(
  '[aria-label^="Enviar"], [aria-label^="Send"], span[data-icon="send"], span[data-icon="wds-ic-send-filled"]')]
  .filter(e => e.offsetParent !== null && !e.closest('footer'));
if (!cands.length) return null;
return cands[0].closest('button, [role="button"]') || cands[0];
"""

# Quantas mensagens JÁ ENVIADAS na conversa contêm o trecho (ignora a caixa de texto)
JS_CONTAR_MENSAGENS = """
const trecho = arguments[0];
return [...document.querySelectorAll('#main [role="row"]')]
  .filter(r => !r.closest('footer') && !r.querySelector('[contenteditable="true"]')
               && (r.innerText || '').includes(trecho)).length;
"""


def trecho_marcador(texto):
    """Um pedaço que identifica ESTA mensagem: o começo da linha mais longa.
    (O link não serve: outras mensagens da campanha usam o mesmo link.)"""
    linhas = [l.strip() for l in texto.splitlines() if l.strip() and "http" not in l]
    return max(linhas, key=len)[:40]


def contar_mensagens(driver, marcador):
    try:
        return driver.execute_script(JS_CONTAR_MENSAGENS, marcador)
    except WebDriverException:
        return -1


def clicar_anexar(driver):
    for escopo in ("footer ", ""):
        for seletor in SELETORES_ANEXAR:
            for botao in driver.find_elements(By.CSS_SELECTOR, escopo + seletor):
                if botao.is_displayed():
                    botao.click()
                    time.sleep(1.5)
                    return True
    return False


def campo_de_fotos(driver):
    """Campo de arquivo de FOTOS E VÍDEOS (nunca o da figurinha)."""
    for campo in driver.find_elements(By.CSS_SELECTOR, "input[type='file']"):
        if "video" in (campo.get_attribute("accept") or ""):
            return campo
    return None


def anexar_foto(driver, imagem):
    """Abre o anexo > Fotos e vídeos e entrega o arquivo. Devolve None ou o erro."""
    driver.execute_script(JS_INTERCEPTAR_ARQUIVO)
    if campo_de_fotos(driver) is None:
        if not clicar_anexar(driver):
            return "Erro: botão de anexo não encontrado"
        item = driver.execute_script(JS_ITEM_FOTOS)
        if item is None:
            return "Erro: opção Fotos e vídeos não encontrada"
        item.click()
        time.sleep(1)
    campo = campo_de_fotos(driver) or driver.execute_script("return window.__inputFoto")
    if campo is None or "video" not in (campo.get_attribute("accept") or ""):
        return "Erro: opção Fotos e vídeos não encontrada"      # nunca usar o de figurinha
    campo.send_keys(str(Path(imagem).resolve()))
    return None


def fechar_previa(driver):
    """Fecha a pré-visualização da imagem sem enviar (usado na pausa por ESC e em erros).
    No WhatsApp real o ESC não fecha: é preciso clicar no "Fechar" (ícone x-alt)."""
    try:
        # Na prévia de FOTO há dois "X": o "Fechar" (canto superior) e o x-alt da
        # caixa de legenda (que só apaga a legenda). Sempre o "Fechar" primeiro.
        botao = driver.execute_script("""
          const vis = e => e.offsetParent !== null && !e.closest('footer') && !e.closest('#side');
          const fechar = [...document.querySelectorAll('button[aria-label="Fechar"], button[aria-label="Close"]')].filter(vis);
          if (fechar.length) return fechar[0];
          const x = [...document.querySelectorAll('span[data-icon="x-alt"]')].filter(vis);
          return x.length ? (x[0].closest('button, [role="button"]') || x[0]) : null;""")
        if botao is not None:
            botao.click()
            time.sleep(1)
        else:
            driver.find_element(By.TAG_NAME, "body").send_keys(Keys.ESCAPE)
        # a prévia de foto pergunta "Deseja descartar a seleção?" -> Descartar
        descartar = driver.execute_script("""
          const b = [...document.querySelectorAll('[role="dialog"] button, [role="dialog"] [role="button"]')]
            .filter(e => /^(descartar|discard)$/i.test((e.innerText || '').trim()));
          return b.length ? b[0] : null;""")
        if descartar is not None:
            descartar.click()
            time.sleep(1)
    except WebDriverException:
        pass


def enviar_com_imagem(driver, posicao, nome, numero, status_ok, imagem):
    """Foto + legenda: abre a conversa com o texto na caixa, anexa por "Fotos e
    vídeos", garante o texto na legenda e envia. Só conta como enviado se uma
    mensagem NOVA com o texto aparecer na conversa."""
    texto = montar_mensagem(nome)
    marcador = trecho_marcador(texto)
    try:
        if not Path(imagem).exists():
            print(f"  ❌ Imagem não encontrada: {imagem}")
            raise ParadaEmergencia
        driver.get(montar_link(numero, texto))
        esperar(TEMPO_BUSCAR, "Carregando a conversa")
        if not caixa_de_mensagem(driver) and not numero_invalido(driver):
            esperar_conversa(driver, TEMPO_BUSCAR_EXTRA)
        if numero_invalido(driver):
            return registrar(posicao, nome, numero, "Número errado")
        if caixa_de_mensagem(driver) is None:
            return registrar(posicao, nome, numero, "Erro: conversa não carregou")

        # 1) a mensagem na caixa (se a prévia não tiver legenda própria, vai daqui)
        if marcador not in texto_da_caixa(driver):
            esperar(TEMPO_TEXTO_EXTRA, "Esperando o texto aparecer",
                    ate=lambda: marcador in texto_da_caixa(driver))
        escrito = texto_da_caixa(driver)
        if marcador not in escrito:
            return registrar(posicao, nome, numero, "Erro: mensagem não apareceu")
        if tamanho_sem_espacos(escrito) > tamanho_sem_espacos(texto) + 20:   # rascunho antigo + mensagem
            limpar_caixa(driver)
            return registrar(posicao, nome, numero, "Erro: texto dobrado na caixa")
        antes = contar_mensagens(driver, marcador)

        # 2) anexa como FOTO (Fotos e vídeos)
        erro = anexar_foto(driver, imagem)
        if erro:
            fechar_previa(driver)
            return registrar(posicao, nome, numero, erro)
        esperar(15, "Abrindo a imagem", ate=lambda: driver.execute_script(JS_BOTAO_ENVIAR) is not None)
        if driver.execute_script(JS_BOTAO_ENVIAR) is None:
            fechar_previa(driver)
            return registrar(posicao, nome, numero, "Erro: imagem não abriu")

        # 3) legenda: se a prévia tem caixa própria, garante o texto nela
        legenda = driver.execute_script(JS_LEGENDA)
        if legenda is not None and marcador not in legenda.text:
            driver.execute_script(JS_COLAR, legenda, texto)
            time.sleep(1)
            legenda = driver.execute_script(JS_LEGENDA)
            if legenda is None or marcador not in legenda.text:
                fechar_previa(driver)
                return registrar(posicao, nome, numero, "Erro: legenda não apareceu")

        # 4) envia
        esperar(TEMPO_ANTES_ENTER, "Aguardando para enviar")
        driver.execute_script(JS_BOTAO_ENVIAR).click()
    except ParadaEmergencia:
        fechar_previa(driver)                     # nada foi enviado: fecha a prévia e pausa
        raise
    except WebDriverException:
        fechar_previa(driver)
        return registrar(posicao, nome, numero, "Erro: Chrome não respondeu")

    # Daqui em diante o envio já foi pedido: nunca reenviar a imagem automaticamente.
    status = "Verificar no WhatsApp"
    try:
        esperar(TEMPO_APOS_ENTER + 10, "Enviando a imagem", pode_parar=False,
                ate=lambda: contar_mensagens(driver, marcador) > antes)
        if contar_mensagens(driver, marcador) > antes:
            status = status_ok
            if marcador in texto_da_caixa(driver):
                limpar_caixa(driver)              # o texto foi na legenda: tira a cópia da caixa
        elif marcador in texto_da_caixa(driver) and driver.execute_script(JS_BOTAO_ENVIAR) is None:
            # a foto saiu sem legenda e o texto ficou na caixa: envia o texto logo depois
            print("  ⚠️  O texto não foi junto com a imagem - enviando o texto em seguida.")
            caixa_de_mensagem(driver).send_keys(Keys.ENTER)
            esperar(TEMPO_APOS_ENTER, "Enviando o texto", pode_parar=False,
                    ate=lambda: contar_mensagens(driver, marcador) > antes)
            if contar_mensagens(driver, marcador) > antes:
                status = status_ok
    except WebDriverException:
        pass
    finally:
        registrar(posicao, nome, numero, status)
    return status


def tamanho_sem_espacos(texto):
    """Tamanho ignorando espaços e quebras (o WhatsApp dobra as linhas em branco)."""
    return len(re.sub(r"\s", "", texto))


def limpar_caixa(driver):
    caixa = caixa_de_mensagem(driver)
    if caixa is not None:
        caixa.send_keys(Keys.CONTROL, "a")
        caixa.send_keys(Keys.DELETE)


def enviar(driver, posicao, nome, numero, status_ok="Enviado", tentativa=1):
    """Abre a conversa, envia, confere se saiu e registra. Retorna o status.

    ESC antes do Enter levanta ParadaEmergencia sem registrar nada: a mensagem
    ainda não saiu, então o contato pode ser retomado do zero com segurança.
    """
    imagem = MENSAGENS[MENSAGEM_ATIVA].get("imagem")
    if imagem:
        return enviar_com_imagem(driver, posicao, nome, numero, status_ok, imagem)
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
        if tamanho_sem_espacos(texto) > tamanho_sem_espacos(montar_mensagem(nome)) + 20:
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
    for indice, linha in df.iterrows():
        posicao = linha_excel(indice)                # mesma numeração de linha do Excel
        if validos >= quantidade:
            break
        if marcado_como_enviado(linha):
            continue                                 # coluna de controle diz "Sim"
        bruto = linha[COLUNA_NUMERO]
        numero = formatar_numero(bruto)
        if numero is None:
            if pd.isna(bruto) and not nome_da_planilha(linha[COLUNA_NOME]):
                continue                             # linha sem nome e sem telefone (ex.: totais)
            if str(bruto).strip() not in falhas:
                selecionados.append((posicao, nome_da_planilha(linha[COLUNA_NOME]), bruto))
            continue
        if numero in enviados or numero in falhas or numero in vistos:   # vistos: repetido na planilha
            continue
        if bloqueado(numero) or numero in NUMEROS_EXCLUIDOS:
            continue
        if not no_segmento(linha):                # público escolhido na campanha
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
    Retorna 'c' (continuar), 's' (pular esse contato) ou 'p' (parar).
    Na tela as opções são números: 1 continuar, 2 pular, 0 parar."""
    print(f"\n{LINHA}")
    print("  ⏸  PAUSADO - nada será enviado enquanto você decide.")
    opcoes = ["1", "0"]
    if contato:
        print(f"  Próximo a receber: {contato[0]} - {contato[1]}")
    print(LINHA)
    print("  [1] Continuar os envios")
    if contato:
        print(f"  [2] Pular {contato[0]} (não recebe, nem nos próximos disparos)")
        opcoes.insert(1, "2")
    print("  [0] Parar os envios")
    return {"1": "c", "2": "s", "0": "p"}[perguntar("  Escolha: ", opcoes)]


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
                    if status == "Enviado":
                        marcar_sim_na_planilha(posicao)   # coluna de controle -> "Sim"
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
            if perguntar("  [1] Continuar o disparo   [0] Parar: ", ["1", "0"]) == "0":
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

CAMPANHA_ATIVA = "Personalizada"      # nome da campanha em uso (preenchido ao ativar)
CAMPANHA_CHAVE = None                 # chave em CAMPANHAS (None = personalizada)
NUMEROS_EXCLUIDOS = set()             # números das planilhas de "excluir" da campanha
SEGMENTAR_COLUNA = None               # coluna que divide o público (ex.: "É membro?")
SEGMENTO = None                       # valor escolhido (ex.: "Não" = não membros)
ROTULO_EXCLUIDOS = ""                 # ex.: "clientes.xlsx"


def perguntar(texto, opcoes):
    while True:
        resposta = input(texto).strip().lower()
        if resposta in opcoes:
            return resposta
        print(f"  Opção inválida. Digite: {' / '.join(opcoes)}")


def voltar_ao_menu():
    input("\n  Aperte Enter para voltar ao menu...")


# ------------------------------- campanhas --------------------------------

def ativar_campanha(chave):
    """Liga planilha + mensagem + regras da campanha. Devolve None ou o erro."""
    global PLANILHA, LINHA_CABECALHO, COLUNA_NOME, COLUNA_NUMERO, COLUNA_JA_ENVIADA
    global MENSAGEM_ATIVA, CAMPANHA_ATIVA, CAMPANHA_CHAVE, NUMEROS_EXCLUIDOS, ROTULO_EXCLUIDOS
    global SEGMENTAR_COLUNA, SEGMENTO
    camp = CAMPANHAS[chave]
    excluidos = set()
    for nome_arquivo in camp.get("excluir", []):
        arquivo = caminho_planilha(nome_arquivo)
        if not arquivo.exists():
            return f"planilha de exclusão não encontrada: {arquivo}"
        det = detectar_planilha(arquivo)
        if not det:
            return f"não achei a coluna de telefone em {arquivo.name}"
        excluidos |= set(ler_planilha(arquivo, cabecalho=det[0])[det[2]].map(formatar_numero).dropna())
    caminho = caminho_planilha(camp["planilha"])
    if not caminho.exists():
        return f"planilha não encontrada: {caminho}"
    try:
        detectado = detectar_planilha(caminho)
    except PermissionError:
        return f"a planilha {caminho.name} está aberta no Excel"
    if not detectado:
        return f"não achei colunas de nome e telefone em {caminho.name}"
    cabecalho, col_nome, col_tel, _ = detectado
    controle = None
    if camp.get("usar_controle"):
        controle = escolher_coluna_controle(ler_planilha(caminho, cabecalho=cabecalho))
    PLANILHA, LINHA_CABECALHO, COLUNA_NOME, COLUNA_NUMERO = str(caminho), cabecalho, col_nome, col_tel
    COLUNA_JA_ENVIADA = controle
    MENSAGEM_ATIVA = camp["mensagem"]
    CAMPANHA_ATIVA = f"{camp['nome']} ({camp['publico']})"
    CAMPANHA_CHAVE = chave
    NUMEROS_EXCLUIDOS = excluidos
    SEGMENTAR_COLUNA = camp.get("segmentar") or None
    SEGMENTO = None                   # o público é escolhido depois, no menu
    ROTULO_EXCLUIDOS = ", ".join(camp.get("excluir", []))
    return None


def estado_atual():
    return (PLANILHA, LINHA_CABECALHO, COLUNA_NOME, COLUNA_NUMERO, COLUNA_JA_ENVIADA,
            MENSAGEM_ATIVA, CAMPANHA_ATIVA, CAMPANHA_CHAVE, NUMEROS_EXCLUIDOS, ROTULO_EXCLUIDOS,
            SEGMENTAR_COLUNA, SEGMENTO)


def restaurar(estado):
    global PLANILHA, LINHA_CABECALHO, COLUNA_NOME, COLUNA_NUMERO, COLUNA_JA_ENVIADA
    global MENSAGEM_ATIVA, CAMPANHA_ATIVA, CAMPANHA_CHAVE, NUMEROS_EXCLUIDOS, ROTULO_EXCLUIDOS
    global SEGMENTAR_COLUNA, SEGMENTO
    (PLANILHA, LINHA_CABECALHO, COLUNA_NOME, COLUNA_NUMERO, COLUNA_JA_ENVIADA,
     MENSAGEM_ATIVA, CAMPANHA_ATIVA, CAMPANHA_CHAVE, NUMEROS_EXCLUIDOS, ROTULO_EXCLUIDOS,
     SEGMENTAR_COLUNA, SEGMENTO) = estado


def no_segmento(linha):
    """True se a linha pertence ao público escolhido (ou se não há segmentação)."""
    if not SEGMENTAR_COLUNA or SEGMENTO is None or SEGMENTAR_COLUNA not in linha:
        return True
    valor = linha[SEGMENTAR_COLUNA]
    return not pd.isna(valor) and str(valor).strip().lower() == str(SEGMENTO).strip().lower()


def valores_do_segmento(df=None):
    """Valores da coluna de segmentação e quantos contatos cada um tem de verdade
    (número válido, sem repetidos, sem bloqueados e sem os excluídos da campanha)."""
    df = ler_planilha() if df is None else df
    if not SEGMENTAR_COLUNA or SEGMENTAR_COLUNA not in df.columns:
        return []
    contagem, vistos = Counter(), set()
    for _, linha in df.iterrows():
        numero = formatar_numero(linha[COLUNA_NUMERO])
        if not numero or numero in vistos or bloqueado(numero) or numero in NUMEROS_EXCLUIDOS:
            continue
        vistos.add(numero)
        valor = linha[SEGMENTAR_COLUNA]
        contagem["(vazio)" if pd.isna(valor) else str(valor).strip()] += 1
    return contagem.most_common()


def numeros_da_campanha():
    """Contas da campanha ativa (sem imprimir nada)."""
    df = ler_planilha()
    if SEGMENTAR_COLUNA and SEGMENTO is not None:
        df = df[df.apply(no_segmento, axis=1)]
    numeros = df[COLUNA_NUMERO].map(formatar_numero)
    enviados, falhas, ultimo = ler_historico() if PULAR_JA_ENVIADOS else (set(), set(), None)
    fora = df.apply(marcado_como_enviado, axis=1) if COLUNA_JA_ENVIADA else pd.Series(False, index=df.index)
    todos = {n for n in numeros[~fora].dropna() if not bloqueado(n)}       # repetido conta 1x
    unicos = todos - NUMEROS_EXCLUIDOS
    return {
        "df": df, "enviados": enviados, "falhas": falhas, "ultimo": ultimo,
        "total": len(df), "invalidos": int(numeros.isna().sum()), "fora_controle": int(fora.sum()),
        "ja_receberam": len(unicos & enviados), "pendentes": len(unicos - enviados - falhas),
        "falharam": len((unicos & falhas) - enviados),
        "excluidos": len(todos & NUMEROS_EXCLUIDOS),
    }


def descricao_mensagem():
    msg = MENSAGENS[MENSAGEM_ATIVA]
    return msg["nome"] + ("  🖼  com imagem" if msg.get("imagem") else "")


MARCA_RASCUNHO = "((escrever"          # a mensagem ainda é um rascunho no Campanhas.md


def mensagem_e_rascunho():
    """True quando a campanha existe, mas o texto da mensagem ainda não foi escrito."""
    return MARCA_RASCUNHO in MENSAGENS[MENSAGEM_ATIVA]["texto"].lower()


def aviso_rascunho():
    print(f"\n{LINHA}")
    print("  ✋ A mensagem desta campanha ainda não foi escrita.")
    print(f"     Abra {ARQUIVO_CAMPANHAS}, troque o trecho")
    print(f"     \"{MARCA_RASCUNHO}...\" pelo texto e volte aqui.")
    print("     Nada é enviado enquanto o rascunho estiver lá.")
    print(LINHA)
    voltar_ao_menu()


def cartao_campanha(n):
    """Resumo da campanha ativa, mostrado no topo do menu."""
    atualizada = datetime.fromtimestamp(os.path.getmtime(PLANILHA))
    print(f"  📣 Campanha: {CAMPANHA_ATIVA}")
    print(f"     Planilha: {Path(PLANILHA).name}  (atualizada em {atualizada:%d/%m às %H:%M})")
    print(f"     Mensagem: {descricao_mensagem()}")
    if mensagem_e_rascunho():
        print("     ✋ MENSAGEM AINDA NÃO ESCRITA - esta campanha não dispara (veja o Campanhas.md)")
    imagem = MENSAGENS[MENSAGEM_ATIVA].get("imagem")
    if imagem and not Path(imagem).exists():
        print(f"     ⚠️  IMAGEM NÃO ENCONTRADA: coloque {Path(imagem).name} em {Path(imagem).parent}")
    if SEGMENTAR_COLUNA:
        publico = SEGMENTO if SEGMENTO is not None else "todos"
        print(f"     Público:  \"{SEGMENTAR_COLUNA}\" = {publico}   (troque em [3])")
    if COLUNA_JA_ENVIADA:
        print(f"     Regra:    só recebe quem está \"Não\" em \"{COLUNA_JA_ENVIADA}\" (vira \"Sim\" após o envio)")
    print()
    print(f"     ⏳ Pendentes: {n['pendentes']}     ✅ Já receberam: {n['ja_receberam']}"
          f"     ❌ Números inválidos: {n['invalidos']}")
    extras = []
    if COLUNA_JA_ENVIADA:
        extras.append(f"{n['fora_controle']} já marcados \"Sim\" na planilha")
    if n.get("excluidos"):
        extras.append(f"{n['excluidos']} ficam de fora por estarem em {ROTULO_EXCLUIDOS}")
    if n["falharam"]:
        extras.append(f"{n['falharam']} com falha antes (não são tentados de novo)")
    if extras:
        print("     (" + " · ".join(extras) + ")")
    if n["pendentes"]:
        print(f"     Tempo para todos os pendentes: ~{formatar_duracao(n['pendentes'] * segundos_por_contato())}")
    if n["ultimo"]:
        linha, nome, numero, horario = n["ultimo"]
        print(f"     Último enviado: linha {linha} - {nome} em {horario}")


def menu_inicial(chrome_ligado):
    os.system("cls")
    print(LINHA)
    print("  Oi! Eu sou o Sender 👋")
    print(LINHA)
    try:
        n = numeros_da_campanha()
        cartao_campanha(n)
    except PermissionError:
        n = None
        print(f"  ⚠️  A planilha {Path(PLANILHA).name} está aberta no Excel. Feche e volte ao menu.")
    except (FileNotFoundError, KeyError, ValueError) as erro:
        n = None
        print(f"  ⚠️  Não consegui ler a planilha da campanha ({Path(PLANILHA).name}): {erro}")
        print(f"     Coloque a planilha em {PASTA_PLANILHAS} ou escolha outra campanha em [3].")
    print(f"\n{LINHA}")
    print("  [1] Testar a mensagem (envia só para o seu número)")
    print("  [2] Disparar para os pendentes")
    print("  [3] Trocar de campanha")
    print("  [4] Ver a mensagem completa")
    print("  [0] Sair" + (" e fechar o Chrome" if chrome_ligado else ""))
    print(LINHA)
    return perguntar("  Escolha: ", ["1", "2", "3", "4", "0"]), n


def escolher_segmento():
    """Pergunta para qual público da planilha a campanha vai (ex.: membro / não membro)."""
    global SEGMENTO
    valores = valores_do_segmento()
    if not valores:
        print(f"  ⚠️  A planilha não tem a coluna \"{SEGMENTAR_COLUNA}\" - a campanha vai para todos.")
        voltar_ao_menu()
        return
    print(f"\n{LINHA}")
    print(f"  Para quem vai esta campanha?  (coluna \"{SEGMENTAR_COLUNA}\")")
    print(LINHA)
    for i, (valor, quantos) in enumerate(valores, start=1):
        print(f"  [{i}] {valor}  ({quantos} contatos)")
    print(f"  [{len(valores) + 1}] Todos")
    escolha = perguntar("  Escolha: ", [str(i) for i in range(1, len(valores) + 2)])
    SEGMENTO = None if int(escolha) > len(valores) else valores[int(escolha) - 1][0]
    print(f"  ✅ Público: {SEGMENTO if SEGMENTO is not None else 'todos'}")


def escolher_campanha():
    """Lista as campanhas prontas (com pendentes de cada uma) + a opção personalizada."""
    carregar_campanhas()          # relê o Campanhas.md: campanha nova aparece na hora
    estado = estado_atual()
    chaves = list(CAMPANHAS)
    print(f"\n{LINHA}")
    print("  Qual campanha vamos usar?")
    print(LINHA)
    for i, chave in enumerate(chaves, start=1):
        camp = CAMPANHAS[chave]
        erro = ativar_campanha(chave)
        if erro:
            situacao = f"⚠️  {erro}"
        else:
            n = numeros_da_campanha()
            imagem = "  🖼 com imagem" if MENSAGENS[camp["mensagem"]].get("imagem") else ""
            situacao = f"{n['pendentes']} pendentes · planilha {camp['planilha']}{imagem}"
        ativa = "  ← ativa" if chave == estado[7] else ""
        print(f"\n  [{i}] {camp['nome']} — {camp['publico']}{ativa}")
        print(f"      {situacao}")
        restaurar(estado)
    print("\n  [9] Personalizada (escolher planilha e mensagem à mão)")
    print("  [0] Voltar sem trocar")
    opcoes = [str(i) for i in range(len(chaves) + 1)] + ["9"]
    escolha = perguntar("  Escolha: ", opcoes)
    if escolha == "0":
        return
    if escolha == "9":
        campanha_personalizada()
        return
    erro = ativar_campanha(chaves[int(escolha) - 1])
    if erro:
        print(f"\n  ❌ Não deu para ativar: {erro}")
        restaurar(estado)
        voltar_ao_menu()
        return
    if SEGMENTAR_COLUNA:
        escolher_segmento()


def campanha_personalizada():
    """Monta uma campanha escolhendo a planilha e a mensagem à mão."""
    global CAMPANHA_ATIVA, CAMPANHA_CHAVE, NUMEROS_EXCLUIDOS, ROTULO_EXCLUIDOS
    estado = estado_atual()
    CAMPANHA_CHAVE, NUMEROS_EXCLUIDOS, ROTULO_EXCLUIDOS = None, set(), ""
    if not trocar_planilha():
        restaurar(estado)
        return
    if not trocar_mensagem():
        restaurar(estado)
        return
    CAMPANHA_ATIVA = f"Personalizada ({Path(PLANILHA).name})"


def ver_mensagem():
    msg = MENSAGENS[MENSAGEM_ATIVA]
    print(f"\n{LINHA}")
    print(f"  💬 {msg['nome']}")
    print(LINHA)
    if msg.get("imagem"):
        existe = "✅" if Path(msg["imagem"]).exists() else "❌ ARQUIVO NÃO ENCONTRADO"
        print(f"  🖼  Imagem: {msg['imagem']}  {existe}\n")
    for linha in msg["texto"].replace("{Nome}", "Fulano").splitlines():
        print(f"  │ {linha}")
    voltar_ao_menu()


# ------------------------- planilha / mensagem à mão ------------------------

def trocar_planilha():
    """Lista as planilhas da pasta, detecta cabeçalho e colunas e troca a ativa.
    Devolve True se trocou."""
    global PLANILHA, LINHA_CABECALHO, COLUNA_NOME, COLUNA_NUMERO, COLUNA_JA_ENVIADA
    pasta = Path(PASTA_PLANILHAS)
    arquivos = planilhas_da_pasta()
    print(f"\n{LINHA}")
    print(f"  Passo 1/2 - Qual planilha?  (pasta: {pasta})")
    print(LINHA)
    if not arquivos:
        print("  Nenhuma planilha encontrada. Coloque o arquivo .xlsx nessa pasta.")
        voltar_ao_menu()
        return False
    for i, a in enumerate(arquivos, start=1):
        subpasta = "" if a.parent == pasta else f"  [{a.parent.name}]"
        print(f"  [{i}] {a.name}{subpasta}  ({datetime.fromtimestamp(a.stat().st_mtime):%d/%m %H:%M})")
    print("  [0] Cancelar")
    escolha = perguntar("  Escolha: ", [str(i) for i in range(len(arquivos) + 1)])
    if escolha == "0":
        return False
    caminho = arquivos[int(escolha) - 1]

    try:
        detectado = detectar_planilha(caminho)
    except PermissionError:
        print("  ❌ A planilha está aberta no Excel. Feche e tente de novo.")
        voltar_ao_menu()
        return False
    if not detectado:
        print("  ❌ Não achei colunas de nome e telefone nessa planilha.")
        voltar_ao_menu()
        return False
    cabecalho, col_nome, col_tel, colunas = detectado
    df = ler_planilha(caminho, cabecalho=cabecalho)
    col_controle = escolher_coluna_controle(df)
    if col_controle:
        nao = df[col_controle].astype(str).str.strip().str.lower().isin({"não", "nao", "n"}).sum()
        print(f"\n  A planilha tem a coluna \"{col_controle}\" ({nao} com \"Não\").")
        print("  Enviar só para quem está \"Não\" (e marcar \"Sim\" depois)?")
        resp = perguntar("  [1] Sim   [0] Não, ignorar a coluna: ", ["1", "0"])
        if resp == "0":
            col_controle = None

    while True:
        numeros = df[col_tel].map(formatar_numero) if col_tel else pd.Series(dtype=str)
        print(f"\n  Cabeçalho na linha {cabecalho + 1} do Excel | {len(df)} linhas")
        print(f"  Nome: {col_nome}  |  Telefone: {col_tel} ({numeros.notna().sum()} válidos)")
        for _, linha in df.head(3).iterrows():
            nome = nome_da_planilha(linha[col_nome]) if col_nome else "?"
            print(f"    - {nome or '(sem nome)'} | {formatar_numero(linha[col_tel]) or 'número inválido'}")
        resposta = perguntar("  [1] Está certo   [2] Escolher as colunas   [0] Cancelar: ", ["1", "2", "0"])
        if resposta == "0":
            return False
        if resposta == "1" and col_nome and col_tel:
            break
        for i, c in enumerate(colunas, start=1):
            print(f"    [{i:>2}] {c}")
        opcoes = [str(i) for i in range(1, len(colunas) + 1)]
        col_nome = colunas[int(perguntar("  Número da coluna do NOME: ", opcoes)) - 1]
        col_tel = colunas[int(perguntar("  Número da coluna do TELEFONE: ", opcoes)) - 1]

    PLANILHA, LINHA_CABECALHO, COLUNA_NOME, COLUNA_NUMERO = str(caminho), cabecalho, col_nome, col_tel
    COLUNA_JA_ENVIADA = col_controle
    return True


def trocar_mensagem():
    """Mostra as mensagens disponíveis com prévia e troca a ativa. Devolve True se trocou."""
    global MENSAGEM_ATIVA
    chaves = list(MENSAGENS)
    print(f"\n{LINHA}")
    print("  Passo 2/2 - Qual mensagem?")
    print(LINHA)
    for i, chave in enumerate(chaves, start=1):
        msg = MENSAGENS[chave]
        imagem = "  🖼 com imagem" if msg.get("imagem") else ""
        primeira = next((l for l in msg["texto"].splitlines()[1:] if l.strip()), "")
        print(f"  [{i}] {msg['nome']}{imagem}")
        print(f"      \"{primeira[:70]}...\"")
    print("  [0] Cancelar")
    escolha = perguntar("  Escolha: ", [str(i) for i in range(len(chaves) + 1)])
    if escolha == "0":
        return False
    MENSAGEM_ATIVA = chaves[int(escolha) - 1]
    return True


# ------------------------------ teste e disparo ------------------------------

def pedir_contato_teste():
    """Pede o número (e o nome, só se a mensagem usa {Nome})."""
    usa_nome = "{Nome}" in MENSAGENS[MENSAGEM_ATIVA]["texto"]
    print(f"\n{LINHA}")
    print(f"  🧪 TESTE da campanha: {CAMPANHA_ATIVA}")
    print(f"     Mensagem: {descricao_mensagem()}")
    print(LINHA)
    nome = ""
    if usa_nome:
        print("  Digite o NOME que vai aparecer na mensagem e aperte Enter (ex.: você)")
        while not nome:
            nome = primeiro_nome(input("  Nome: "))
    print("  Digite o NÚMERO que vai receber o teste e aperte Enter")
    print("  (DDD + número, sem 55, sem espaço ou traço. Ex.: 15991234567)")
    while True:
        numero = formatar_numero(input("  Número: "))
        if numero:
            return nome, numero
        print("  Número inválido. Use DDD + número, ex.: 15991234567")


def escolher_quantidade(n):
    """Pergunta quantos disparar, mostra de onde até onde vai e confirma.
    Devolve a lista de contatos ou None."""
    pendentes = n["pendentes"]
    if pendentes == 0:
        print("\n  Não há pendentes nesta campanha: todos já receberam.")
        voltar_ao_menu()
        return None
    print(f"\n{LINHA}")
    print(f"  📣 {CAMPANHA_ATIVA}")
    print(f"  💬 {descricao_mensagem()}")
    print(f"  Pendentes: {pendentes}  |  todos: {previsao_termino(pendentes * segundos_por_contato())}")
    print(LINHA)
    while True:
        resposta = input(f"  Quantos enviar agora? (1 a {pendentes}, Enter = todos, 0 = voltar): ").strip()
        if not resposta:
            quantidade = pendentes
            break
        if resposta.isdigit() and 0 <= int(resposta) <= pendentes:
            quantidade = int(resposta)
            break
        print(f"  Digite um número de 1 a {pendentes}.")
    if quantidade == 0:
        return None

    contatos = contatos_da_planilha(n["df"], n["enviados"], n["falhas"], quantidade)
    validos = [c for c in contatos if formatar_numero(c[2])]
    primeiro, ultimo = validos[0], validos[-1]
    print(f"\n  Vou enviar para {quantidade} contatos: da linha {primeiro[0]} ({primeiro[1]})"
          f" até a linha {ultimo[0]} ({ultimo[1]}).")
    print(f"  Tempo estimado: {previsao_termino(quantidade * segundos_por_contato())}.")
    confirma = perguntar("  [1] Começar o disparo   [0] Voltar: ", ["1", "0"])
    return contatos if confirma == "1" else None


def chrome_aberto(driver):
    if driver is None:
        return False
    try:
        driver.title          # falha se o Chrome foi fechado na mão
        return True
    except Exception:
        return False


def garantir_chrome(driver):
    """Devolve um Chrome aberto e logado (reaproveita o que já está aberto)."""
    if chrome_aberto(driver):
        return driver
    if driver is not None:                   # sobra de um Chrome que travou
        try:
            driver.quit()
        except WebDriverException:
            pass
    return abrir_whatsapp()


def main():
    driver = None
    manter_pc_acordado(True)
    if carregar_campanhas():
        print(f"  Campanhas lidas de {ARQUIVO_CAMPANHAS}")
    inicial = CAMPANHA_INICIAL if CAMPANHA_INICIAL in CAMPANHAS else next(iter(CAMPANHAS), None)
    erro = ativar_campanha(inicial) if inicial else "nenhuma campanha configurada"
    if erro:
        print(f"  ⚠️  Campanha inicial não ativada ({erro}). Usando a configuração padrão.")
        voltar_ao_menu()
    try:
        while True:
            escolha, n = menu_inicial(chrome_aberto(driver))
            if escolha == "0":
                break
            if escolha == "3":
                escolher_campanha()
                continue
            if escolha == "4":
                ver_mensagem()
                continue
            if n is None:
                print("\n  ⚠️  A planilha desta campanha não está disponível. Use [3] para trocar de campanha.")
                voltar_ao_menu()
                continue

            if mensagem_e_rascunho():
                aviso_rascunho()
                continue

            contatos = escolher_quantidade(n) if escolha == "2" else None
            if escolha == "2" and contatos is None:
                continue
            teste = pedir_contato_teste() if escolha == "1" else None

            try:
                driver = garantir_chrome(driver)
                while teste:
                    try:
                        disparar(driver, [("Teste", *teste)], set(),
                                 titulo=f"🧪 TESTE - {CAMPANHA_ATIVA}",
                                 status_ok="Teste enviado", tipo="Teste")
                        print("\n  Confira no celular se a mensagem chegou certa.")
                    except ParadaEmergencia:
                        aviso_parada()
                    print("\n  [1] Testar de novo")
                    print("  [2] Está tudo certo: disparar esta campanha")
                    print("  [0] Voltar ao menu")
                    escolha = perguntar("  Escolha: ", ["1", "2", "0"])
                    if escolha == "1":
                        teste = pedir_contato_teste()
                    elif escolha == "2" and (contatos := escolher_quantidade(numeros_da_campanha())):
                        teste = None
                    else:
                        break
                else:
                    disparar(driver, contatos, n["enviados"])
                    voltar_ao_menu()

            except ParadaEmergencia:
                aviso_parada()
                voltar_ao_menu()
            except WebDriverException as falha:
                # Chrome fechado/travado no meio: não derruba o Sender, volta pro menu
                print(f"\n{LINHA}")
                print("  ❌ O Chrome do robô parou de responder (foi fechado ou travou).")
                print(f"     Detalhe: {str(falha).splitlines()[0][:100]}")
                print("  Nada mais será enviado. No próximo disparo ele abre o Chrome de novo")
                print("  e continua de onde parou.")
                print(LINHA)
                mostrar_resumo()
                voltar_ao_menu()

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
    print(f"  Registro completo em: dados/{Path(ARQUIVO_REGISTRO).name}")
    RESUMO.clear()


if __name__ == "__main__":
    main()
