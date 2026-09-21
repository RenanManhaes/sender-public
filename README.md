# 📤 Sender — disparo de WhatsApp a partir de uma planilha

Automação em **Python + Selenium** que lê uma lista de contatos no Excel e envia uma mensagem personalizada para cada um pelo **WhatsApp Web**, com texto ou **foto com legenda**. Tem timers entre as etapas, pausa por ESC, memória de quem já recebeu e registro de tudo em Markdown.

Não tem IA no processo: é automação de navegador, feita para rodar no seu computador enquanto você acompanha.

> ⚠️ **Leia antes de usar:** automatizar o WhatsApp Web vai contra os Termos de Uso do WhatsApp e pode levar ao **banimento do número**. Use com lista própria (pessoas que forneceram o contato), em lotes pequenos e por sua conta e risco. Veja [docs/seguranca.md](docs/seguranca.md).

---

## O que ele faz

- **Campanhas:** cada campanha junta **planilha + mensagem + regras**. No menu você só escolhe qual usar.
- **Texto ou foto com legenda:** a mensagem pode levar uma imagem. Ela vai por *Anexo › Fotos e vídeos*, nunca como figurinha, com o texto na legenda, numa mensagem só.
- **Nome na mensagem:** `{Nome}` vira o primeiro nome da planilha ("MARIA SILVA" → "Maria").
- **Qualquer planilha:** descobre sozinho a linha do cabeçalho e as colunas de nome e telefone. Funciona até com exportações que têm relatório antes da tabela.
- **Coluna de controle (opcional):** só recebe quem está com **"Não"** numa coluna como "Mensagem enviada?". Depois de cada envio confirmado, a linha vira **"Sim"** e a planilha é salva.
- **Lista de exclusão por campanha:** por exemplo, não mandar o convite para quem já comprou (`"excluir": ["clientes.xlsx"]`).
- **Nunca envia duas vezes** para a mesma pessoa. A memória é o próprio arquivo de registro, separada por mensagem.
- **Confere se a mensagem saiu:** só conta como enviado quando aparece uma mensagem **nova** com o texto na conversa.
- **Valida números brasileiros:** celular com 9, fixo e celular antigo sem o 9. O que não for telefone válido fica de fora.
- **Pausa por ESC** a qualquer momento: `[1]` continuar, `[2]` pular o contato, `[0]` parar.
- **Freio automático** depois de 3 erros seguidos, por exemplo se o WhatsApp deslogar ou a internet cair.
- **Respostas sempre por número:** `1` = sim/continuar, `2` = segunda opção, `0` = não/voltar/parar.
- **Chrome minimizado:** não usa mouse, teclado nem área de transferência, então o computador continua livre durante o envio.

Acompanha o **verificador "Agora não"** ([docs](docs/verificador-agora-nao.md)). Ele pesquisa uma resposta no WhatsApp Web, por exemplo quem respondeu "Agora não" a um convite anterior, cruza com a planilha e move essas pessoas para uma aba "Não enviar". Faz backup antes.

---

## Instalação

Requisitos: **Windows**, **Python 3.12+** e **Google Chrome**.

```bash
git clone https://github.com/<seu-usuario>/sender-whatsapp-bot.git
cd sender-whatsapp-bot
pip install -r requirements.txt
```

O `chromedriver` é resolvido automaticamente pelo Selenium 4.

### Prepare os dados

1. Copie `dados/contatos-exemplo.xlsx` para `dados/contatos.xlsx` e troque pelos seus contatos. As colunas **Nome** e **Telefone** são obrigatórias. **Mensagem enviada?** é opcional.
2. Para a campanha com imagem, troque `dados/imagem-exemplo.jpg` pela sua imagem, em JPG ou PNG. WEBP pode virar figurinha.
3. Ajuste mensagens e campanhas no topo do `sender.py`. Os detalhes estão em [docs/configuracao.md](docs/configuracao.md).

---

## Uso

```bash
python sender.py
```
Também dá para usar o atalho `Sender.bat`.

```
============================================================
  Oi! Eu sou o Sender 👋
============================================================
  📣 Campanha: Convite com imagem (lista de contatos)
     Planilha: contatos.xlsx  (atualizada em 01/01 às 10:00)
     Mensagem: Convite com imagem (exemplo)  🖼  com imagem
     Regra:    só recebe quem está "Não" em "Mensagem enviada?" (vira "Sim" após o envio)

     ⏳ Pendentes: 3     ✅ Já receberam: 0     ❌ Números inválidos: 1
     Tempo para todos os pendentes: ~3min
============================================================
  [1] Testar a mensagem (envia só para o seu número)
  [2] Disparar para os pendentes
  [3] Trocar de campanha
  [4] Ver a mensagem completa
  [0] Sair
============================================================
```

| Opção | O que faz |
|---|---|
| **[1] Testar** | Envia a mensagem da campanha ativa, com a imagem se houver, para o seu número. Só pede o nome se a mensagem usar `{Nome}`. |
| **[2] Disparar** | Pergunta **quantos** enviar e mostra de qual linha até qual linha vai e o tempo estimado. Só começa depois de **[1] Começar o disparo**. |
| **[3] Trocar de campanha** | Lista as campanhas com os pendentes de cada uma. **[9] Personalizada** deixa escolher à mão qualquer planilha da pasta `dados/` e qualquer mensagem. |
| **[4] Ver a mensagem** | Mostra o texto completo e confere se a imagem existe. |

Na primeira vez, escaneie o QR code. O perfil próprio do Chrome guarda o login para as próximas.

### Pausa (ESC)

```
  ⏸  PAUSADO - nada será enviado enquanto você decide.
  Próximo a receber: Fulano - 5515999999999
  [1] Continuar os envios
  [2] Pular Fulano (não recebe, nem nos próximos disparos)
  [0] Parar os envios
```

---

## Relatório para clientes

```bash
python relatorio.py
```
Também dá para usar o atalho `Relatorio.bat`.

Gera `dados/Relatório de campanhas.md`, que fica ótimo no Obsidian. O relatório traz:
- resumo executivo;
- resultado por campanha, com gráfico;
- situação atual de cada campanha;
- qualidade das bases, com os motivos de número inválido;
- proteções da reputação do número;
- linha do tempo e as mensagens usadas.

Tem **só números agregados**, sem nomes nem telefones. Rode de novo depois de cada disparo para atualizar. O título e a data do evento ficam no topo do `relatorio.py`.

## Registro

Cada disparo vira uma seção em `dados/Registros.md`:

```markdown
## 🚀 Disparo oficial — 01/01/2026 às 10:00

- **Planilha:** `contatos.xlsx`
- **Mensagem:** Convite (exemplo) (`convite`)
- **Contatos selecionados:** 30

| Nº | Nome | Número | Status | Horário |
|---:|---|---|---|---|
| 2 | Maria | 5515991234567 | ✅ Enviado | 01/01 10:01:12 |

> **Resumo:** ✅ Enviado: **1** · ⏱ 1min · finalizado às 10:02
```

Esse arquivo é a memória do robô: é por ele que o Sender sabe quem já recebeu e continua de onde parou. O formato e as regras estão em [docs/registros.md](docs/registros.md).

---

## Estrutura

```
sender-whatsapp-bot/
├── sender.py               # o robô (configuração no topo do arquivo)
├── verificar_agora_nao.py  # verificador "Agora não" (somente leitura)
├── relatorio.py           # relatório agregado para clientes
├── Sender.bat / AgoraNao.bat / Relatorio.bat
├── requirements.txt
├── dados/                  # planilhas, imagens e registros (ignorado pelo git)
│   ├── contatos-exemplo.xlsx
│   └── imagem-exemplo.jpg
└── docs/
    ├── funcionamento.md    # fluxo, status, memória, pausa, validações
    ├── configuracao.md     # todas as opções, mensagens e campanhas
    ├── registros.md        # formato do registro em Markdown
    ├── verificador-agora-nao.md
    └── seguranca.md        # risco de banimento e boas práticas
```

A pasta `dados/` é ignorada pelo git, exceto os dois exemplos: **nomes e telefones nunca devem ir para o repositório**.

---

## Licença

[MIT](LICENSE).
