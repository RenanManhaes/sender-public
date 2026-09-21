"""
Relatório de campanhas do Sender (para apresentar a clientes).

Lê o Registros.md, as planilhas das campanhas e as listas de proteção e gera
um relatório em Markdown (ótimo no Obsidian), só com números agregados — sem
nomes e sem telefones. Rode de novo a qualquer momento para atualizar.

Uso:  python relatorio.py      (ou o atalho Relatorio.bat)
"""

import re
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path

import pandas as pd

import sender

# ------------------------------ configuração ------------------------------
TITULO = "Minha campanha"             # aparece no título do relatório
DATA_EVENTO = ""                      # opcional, ex.: "25/09/2026"
ARQUIVO_RELATORIO = str(Path(sender.ARQUIVO_REGISTRO).with_name("Relatório de campanhas.md"))
# ---------------------------------------------------------------------------

GRUPOS = {
    "Enviado": "✅ Entregue ao WhatsApp",
    "Número errado": "❌ Número sem WhatsApp",
    "Número fora do padrão": "❌ Número inválido / fora do Brasil",
    "Pulado por você": "⏭️ Retirado manualmente",
    "Verificar no WhatsApp": "⚠️ Conferência manual",
}
FALHA_TECNICA = "🔁 Falha técnica (volta para a fila)"


def grupo(status):
    if status in GRUPOS:
        return GRUPOS[status]
    return FALHA_TECNICA if status.startswith("Erro") else status


def ler_secoes():
    """Seções do Registros.md: tipo, data, planilha, mensagem e linhas."""
    secoes, atual = [], None
    arquivo = Path(sender.ARQUIVO_REGISTRO)
    if not arquivo.exists():
        return secoes
    for linha in arquivo.read_text(encoding="utf-8").splitlines():
        if linha.startswith("## "):
            data = re.search(r"(\d{2}/\d{2}/\d{4}) às (\d{2}:\d{2})", linha)
            atual = {"teste": "Teste" in linha, "mensagem": sender.MENSAGEM_ATIVA, "planilha": "",
                     "data": datetime.strptime(" ".join(data.groups()), "%d/%m/%Y %H:%M") if data else None,
                     "linhas": []}
            secoes.append(atual)
        elif atual is None:
            continue
        elif linha.startswith("- **Mensagem:**"):
            achou = re.search(r"\(`([^`]+)`\)", linha)
            if achou:
                atual["mensagem"] = achou.group(1)
        elif linha.startswith("- **Planilha:**") and "`" in linha:
            atual["planilha"] = linha.split("`")[1]
        else:
            celulas = [c.strip() for c in linha.strip().strip("|").split("|")]
            if len(celulas) >= 5 and (celulas[0].isdigit() or celulas[0] == "Teste"):
                atual["linhas"].append({"posicao": celulas[0], "numero": celulas[2],
                                        "status": celulas[3].split(" ", 1)[-1]})
    return secoes


def resumo_por_mensagem(secoes):
    """Por mensagem: resultado final de cada pessoa, lotes e período."""
    dados = defaultdict(lambda: {"disparos": 0, "datas": [], "ultimo_status": {}})
    for s in secoes:
        if s["teste"] or not s["linhas"]:
            continue
        d = dados[s["mensagem"]]
        d["disparos"] += 1
        d["datas"].append(s["data"])
        for l in s["linhas"]:
            chave = (f"{s['planilha']}:{l['posicao']}" if l["status"] == "Número fora do padrão"
                     else l["numero"])
            if d["ultimo_status"].get(chave) != "Enviado":
                d["ultimo_status"][chave] = l["status"]
    for d in dados.values():
        d["resultado"] = Counter(grupo(st) for st in d["ultimo_status"].values())
        d["entregues"] = sum(1 for st in d["ultimo_status"].values() if st == "Enviado")
    return dados


def motivo_invalido(valor):
    if pd.isna(valor):
        return "Sem telefone"
    digitos = re.sub(r"\D", "", str(valor))
    if not digitos:
        return "Sem telefone"
    if len(digitos) == 11 and digitos[0] == "1" and digitos[2] != "9":
        return "Número dos EUA (sem +1)"
    if len(digitos) >= 12 and not digitos.startswith("55"):
        return "Número de outro país"
    if len(digitos) < 10:
        return "Número incompleto"
    return "Formato inválido"


def raio_x_planilha(caminho):
    det = sender.detectar_planilha(caminho)
    if not det:
        return None
    df = sender.ler_planilha(caminho, cabecalho=det[0])
    numeros = df[det[2]].map(sender.formatar_numero)
    validos = numeros.dropna()
    try:
        nao_enviar = len(pd.read_excel(caminho, sheet_name="Não enviar", dtype=str).dropna(how="all"))
    except ValueError:
        nao_enviar = 0
    return {"linhas": len(df), "unicos": int(validos.nunique()),
            "duplicados": int(validos.size - validos.nunique()),
            "invalidos": Counter(motivo_invalido(v) for v, n in zip(df[det[2]], numeros) if pd.isna(n)),
            "nao_enviar": nao_enviar}


def tabela(cabecalho, linhas):
    saida = ["| " + " | ".join(cabecalho) + " |", "|" + "|".join("---" for _ in cabecalho) + "|"]
    saida += ["| " + " | ".join(str(c) for c in l) + " |" for l in linhas]
    return "\n".join(saida)


def fmt_periodo(datas):
    datas = sorted(d for d in datas if d)
    if not datas:
        return "—"
    ini, fim = datas[0], datas[-1]
    return f"{ini:%d/%m}" if ini.date() == fim.date() else f"{ini:%d/%m} a {fim:%d/%m}"


def gerar():
    agora = datetime.now()
    secoes = ler_secoes()
    por_msg = resumo_por_mensagem(secoes)
    testes = sum(1 for s in secoes if s["teste"] and s["linhas"])
    nomes_msg = {k: v["nome"] for k, v in sender.MENSAGENS.items()}

    # bases = todas as planilhas usadas pelas campanhas
    bases = {}
    for camp in sender.CAMPANHAS.values():
        caminho = Path(sender.PASTA_PLANILHAS) / camp["planilha"]
        if caminho.exists() and camp["planilha"] not in bases:
            bases[camp["planilha"]] = raio_x_planilha(caminho)
    bases = {k: v for k, v in bases.items() if v}

    agora_nao_md = Path(sender.ARQUIVO_REGISTRO).with_name("Agora não.md")
    agora_nao = (len(re.findall(r"^\| \+?\d", agora_nao_md.read_text(encoding="utf-8"), re.M))
                 if agora_nao_md.exists() else 0)
    nunca = len(sender.nunca_contatar())

    alcancados = set()
    for d in por_msg.values():
        alcancados |= {n for n, st in d["ultimo_status"].items() if st == "Enviado"}
    entregas = sum(d["entregues"] for d in por_msg.values())
    tentativas = sum(d["entregues"] + d["resultado"].get(GRUPOS["Número errado"], 0)
                     + d["resultado"].get(FALHA_TECNICA, 0) for d in por_msg.values())

    estado = sender.estado_atual()
    situacao = []
    for chave, camp in sender.CAMPANHAS.items():
        erro = sender.ativar_campanha(chave)
        if erro:
            situacao.append((camp["nome"], camp["publico"], "—", "—", f"⚠️ {erro}"))
            continue
        n = sender.numeros_da_campanha()
        feitas = por_msg.get(camp["mensagem"], {}).get("entregues", 0)
        status = ("✅ Concluída" if n["pendentes"] == 0 and feitas
                  else "🟡 Em andamento" if feitas else "⏳ Não iniciada")
        situacao.append((camp["nome"], camp["publico"], n["ja_receberam"], n["pendentes"], status))
    sender.restaurar(estado)

    L = [f"# 📊 Relatório de campanhas de WhatsApp — {TITULO}\n"]
    evento = f"Evento em **{DATA_EVENTO}**. " if DATA_EVENTO else ""
    L.append(f"> [!info] Sobre este relatório\n> {evento}Dados consolidados em **{agora:%d/%m/%Y às %H:%M}** "
             "a partir do registro automático de envios. Apenas números agregados: nenhum nome ou telefone é exibido.\n")

    L.append("## Resumo executivo\n")
    total_base = sum(b["linhas"] for b in bases.values())
    L.append(tabela(["Indicador", "Resultado"], [
        ("Contatos nas planilhas das campanhas", total_base),
        ("Mensagens entregues ao WhatsApp", entregas),
        ("Pessoas alcançadas (sem repetição)", len(alcancados)),
        ("Taxa de entrega (entregues ÷ tentativas com número válido)",
         f"{entregas / tentativas:.0%}" if tentativas else "—"),
        ("Campanhas com envios", len([d for d in por_msg.values() if d["entregues"]])),
        ("Contatos protegidos (fora das listas)", agora_nao + nunca),
    ]))

    L.append("\n## Resultado por campanha\n")
    linhas = []
    for chave, d in sorted(por_msg.items(), key=lambda x: min(x[1]["datas"])):
        r = d["resultado"]
        linhas.append((nomes_msg.get(chave, chave), fmt_periodo(d["datas"]), d["disparos"], len(d["ultimo_status"]),
                       d["entregues"], r.get(GRUPOS["Número fora do padrão"], 0) + r.get(GRUPOS["Número errado"], 0),
                       r.get(FALHA_TECNICA, 0)))
    L.append(tabela(["Campanha", "Período", "Lotes", "Contatos trabalhados", "✅ Entregues",
                     "❌ Inválidos", "🔁 Falha técnica*"], linhas))
    L.append("\n\\* Falhas técnicas não contam como envio e voltam para a fila no lote seguinte.\n")
    if entregas:
        L.append("```mermaid\npie showData title Mensagens entregues por campanha")
        L += [f'    "{nomes_msg.get(k, k)}" : {d["entregues"]}' for k, d in por_msg.items() if d["entregues"]]
        L.append("```\n")

    L.append("## Situação atual das campanhas\n")
    L.append(tabela(["Campanha", "Público", "Já receberam", "Pendentes", "Situação"], situacao))

    L.append("\n## Qualidade das bases de contatos\n")
    L.append(tabela(["Planilha", "Contatos", "WhatsApp válido", "Repetidos", "Inválidos"],
                    [(nome, b["linhas"], b["unicos"], b["duplicados"], sum(b["invalidos"].values()))
                     for nome, b in bases.items()]))
    motivos = sum((b["invalidos"] for b in bases.values()), Counter())
    if motivos:
        L.append("\n**Por que um número é considerado inválido:**\n")
        L.append(tabela(["Motivo", "Contatos"], motivos.most_common()))
    L.append("\nNúmeros inválidos **não são abertos no WhatsApp**, o que protege a reputação do número.\n")

    L.append("## Cuidados com a reputação do número\n")
    L.append(tabela(["Proteção", "Contatos"], [
        ("Responderam \"Agora não\" a convites anteriores (nunca recebem)", agora_nao),
        ("… retirados das planilhas (aba \"Não enviar\")", sum(b["nao_enviar"] for b in bases.values())),
        ("Lista \"Nunca contatar\"", nunca),
        ("Números repetidos (recebem uma vez só)", sum(b["duplicados"] for b in bases.values())),
    ]))
    L.append("""
- **Uma mensagem por pessoa em cada campanha**: o sistema lembra quem já recebeu.
- **Intervalo aleatório** entre um envio e outro, e disparos em lotes.
- **Confirmação de entrega**: só conta como enviado quando a mensagem aparece na conversa.
- **Freio automático** depois de falhas seguidas.
- **Teste antes de cada campanha** para um número interno.
""")

    L.append("## Linha do tempo\n")
    tempo = [(f"{s['data']:%d/%m %H:%M}", nomes_msg.get(s["mensagem"], s["mensagem"]),
              sum(1 for l in s["linhas"] if l["status"] == "Enviado"))
             for s in sorted((s for s in secoes if not s["teste"] and s["linhas"]), key=lambda s: s["data"])]
    L.append(tabela(["Data", "Campanha", "Entregues no lote"], tempo))
    L.append(f"\nAlém dos lotes oficiais, foram feitos **{testes} envios de teste** para números internos.\n")

    L.append("## Mensagens utilizadas\n")
    for chave in por_msg:
        msg = sender.MENSAGENS.get(chave)
        if msg:
            L.append(f"### {msg['nome']}{' (com imagem)' if msg.get('imagem') else ''}\n")
            L.append("\n".join("> " + l for l in msg["texto"].replace("{Nome}", "[primeiro nome]").splitlines()))
            L.append("")

    L.append(f"---\n*Relatório gerado automaticamente em {agora:%d/%m/%Y às %H:%M} pelo Sender "
             "(`python relatorio.py`).*\n")
    Path(ARQUIVO_RELATORIO).write_text("\n".join(L), encoding="utf-8")
    return ARQUIVO_RELATORIO


if __name__ == "__main__":
    print(f"  ✅ Relatório atualizado: {gerar()}")
