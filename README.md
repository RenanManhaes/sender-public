# 📤 Sender — disparo de WhatsApp a partir de uma planilha

Automação em **Python + Selenium** que lê uma lista de contatos no Excel e envia uma mensagem personalizada para cada um pelo **WhatsApp Web**, com timers entre as etapas, pausa por ESC, memória de quem já recebeu e registro de tudo em Markdown.

Sem IA no meio do caminho: é automação de navegador, feita para rodar no seu computador, com você acompanhando.

> ⚠️ **Leia antes de usar:** automatizar o WhatsApp Web vai contra os Termos de Uso do WhatsApp e pode levar ao **banimento do número**. Use com lista própria (pessoas que forneceram o contato), em lotes pequenos e por sua conta e risco. Veja [docs/seguranca.md](docs/seguranca.md).

---

## O que ele faz

- **Lê a planilha** (nome + telefone) e monta a fila dos próximos contatos.
- **Abre a conversa já com a mensagem escrita** (`web.whatsapp.com/send?phone=...&text=...`), espera, aperta Enter e **confere se a mensagem saiu**.
- **Nunca envia duas vezes** para a mesma pessoa: a memória é o próprio arquivo de registro.
- **Valida números brasileiros** (celular com 9, fixo, celular antigo sem o 9) e descarta o que não for telefone válido.
- **Pausa por ESC** a qualquer momento: continuar, pular o contato ou parar.
- **Freio automático** depois de 3 erros seguidos (WhatsApp deslogado, internet caiu).
- **Registra cada envio** em Markdown, com status, horário e resumo por disparo.
- **Várias mensagens** no mesmo projeto, com memória separada por mensagem.
- Funciona com o **Chrome minimizado**; não usa mouse, teclado nem área de transferência, então o computador continua livre.

Também acompanha o **verificador "Agora não"** ([docs](docs/verificador-agora-nao.md)): pesquisa uma resposta no WhatsApp Web (ex.: quem respondeu "Agora não" a um convite anterior), cruza com a planilha e move essas pessoas para uma aba "Não enviar", com backup.

---

## Instalação

Requisitos: **Windows**, **Python 3.12+** e **Google Chrome**.

```bash
git clone https://github.com/<seu-usuario>/sender-whatsapp.git
cd sender-whatsapp
pip install -r requirements.txt
```

O `chromedriver` é resolvido automaticamente pelo Selenium 4.

### Prepare a planilha

Uma planilha `.xlsx` com uma coluna de **nome** e uma de **telefone**:

| Nome | Telefone |
|---|---|
| Maria | 15991234567 |
| João | 11987654321 |

Veja `dados/contatos-exemplo.xlsx`. Salve a sua como `dados/contatos.xlsx` ou aponte `PLANILHA` para onde quiser.

### Ajuste a configuração

No topo do `sender.py`: caminhos, nomes das colunas, timers e as mensagens (`MENSAGENS`). Detalhes em [docs/configuracao.md](docs/configuracao.md).

---

## Uso

```bash
python sender.py
```
(ou o atalho `Sender.bat`)

```
============================================================
  Oi! Eu sou o Sender 👋
============================================================
  💬 Mensagem ativa: Convite (exemplo)
  Contatos na planilha: 4
    - números válidos:   3
    - pendentes:         3

  O que vamos fazer?
  [1] Fazer um teste antes (envia a mensagem para um número seu)
  [2] Disparar para os contatos da planilha (você escolhe quantos)
  [3] Trocar a mensagem (ativa: Convite (exemplo))
  [0] Sair
```

1. **Teste:** envia a mensagem para um número seu, pelo mesmo caminho do disparo oficial.
2. **Disparo:** você escolhe **quantos** contatos, e ele mostra de qual linha até qual linha vai, o tempo estimado e a hora prevista de término. Só começa depois do `SIM`.
3. **Login:** na primeira vez, escaneie o QR code; o perfil próprio do Chrome guarda o login para as próximas.

### Pausa (ESC)

```
  ⏸  PAUSADO - nada será enviado enquanto você decide.
  Próximo a receber: Fulano - 5515999999999
  [C] Continuar os envios
  [S] Pular Fulano (não recebe, nem nos próximos disparos)
  [P] Parar os envios
```

---

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

Esse arquivo é a memória do robô: ele lê os números já enviados para continuar de onde parou. Formato e regras em [docs/registros.md](docs/registros.md).

---

## Estrutura

```
sender-whatsapp/
├── sender.py               # o robô (configuração no topo do arquivo)
├── verificar_agora_nao.py  # verificador "Agora não" (somente leitura)
├── Sender.bat / AgoraNao.bat
├── requirements.txt
├── dados/                  # planilha, registros e listas (ignorado pelo git)
└── docs/
    ├── funcionamento.md    # fluxo, status, memória, pausa, validações
    ├── configuracao.md     # todas as opções
    ├── registros.md        # formato do registro em Markdown
    ├── verificador-agora-nao.md
    └── seguranca.md        # risco de banimento e boas práticas
```

A pasta `dados/` é ignorada pelo git: **nomes e telefones nunca devem ir para o repositório**.

---

## Licença

[MIT](LICENSE).
