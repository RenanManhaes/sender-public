# Como o Sender funciona

## Visão geral do fluxo

```
Sender (cmd)
  └─ Menu inicial: lê planilha + Registros.md → mostra totais e tempo
       ├─ [1] Teste ──┐
       └─ [2] Disparo ┤  (escolhe quantidade → mostra linha inicial/final → SIM)
                      ▼
            Abre o Chrome do robô (perfil próprio)
            Pergunta se já está logado → confere / espera QR code
                      ▼
            Para cada contato:
              1. Abre web.whatsapp.com/send?phone=...&text=...   (mensagem já no link)
              2. Carregando a conversa ............ 15s (+ até 30s se a internet estiver lenta)
              3. Aguardando para enviar ........... 5s  (+ até 10s se o texto não apareceu)
              4. Aperta Enter
              5. Enviando a mensagem .............. 5s  → confere se a caixa esvaziou
              6. Registra no Registros.md
              7. Pausa até o próximo contato ...... aleatório entre 20 e 60s
                      ▼
            Resumo no terminal + na pasta `dados/` → volta ao menu (Chrome continua aberto)
```

### Por que `web.whatsapp.com/send` e não `wa.me`

No computador, o `wa.me` abre uma página intermediária ("Continuar para a conversa") que exigiria cliques extras. O `web.whatsapp.com/send?phone=...&text=...` recebe o mesmo número e a mesma mensagem e abre direto a conversa com o texto já escrito. Por isso o robô **não usa Ctrl+V**: a mensagem vai dentro do link (convertida por `mensagem_para_link()`), o que também evita o limite do ChromeDriver com emojis.

---

## A mensagem

Definida em `MENSAGENS` no topo do `sender.py` (a ativa aparece no menu e pode ser trocada em **[3] Trocar a mensagem**). Cada seção do `Registros.md` registra qual mensagem foi usada. O marcador `{Nome}` é trocado pelo nome da **coluna de nome** da planilha, exatamente como está escrito lá.

No **teste**, o nome digitado passa por um tratamento: usa só o primeiro nome, com inicial maiúscula, ignorando partículas (`de`, `da`, `do`, `dos`, `das`, `e`).

> ⚠️ Textos com datas relativas ("faltam 7 dias") precisam ser revistos a cada dia de disparo.

---

## Validação dos números

A coluna **de telefone** é lida como texto e normalizada:

| Formato na planilha | Exemplo | Resultado |
|---|---|---|
| Celular BR: DDD + 9 + 8 dígitos | `15991234567` | `5515991234567` ✅ |
| Já com 55 | `5515991234567` | `5515991234567` ✅ |
| Fixo BR: DDD + 2..5 + 7 dígitos | `1532345678` | `551532345678` ✅ |
| Celular antigo sem o 9 | `1698765432` | `5516998765432` ✅ (o 9 é adicionado) |
| 11 dígitos **sem 9** depois do DDD | `13055550123` | ❌ fora do padrão (é um número dos EUA sem `+1`) |
| 9 dígitos, 12 dígitos de outro país, etc. | `159912345` | ❌ fora do padrão |

Números fora do padrão **não são abertos no WhatsApp**: vão direto para o registro e não contam na quantidade escolhida.

---

## Memória: como ele continua de onde parou

A memória do Sender é o próprio `Registros.md`. Antes de cada disparo ele lê todas as tabelas e monta:

- **Enviados** → nunca mais recebem.
- **Falhas definitivas** → nunca são tentadas de novo.
- **Erros retentáveis** → voltam para a fila no próximo disparo (a mensagem comprovadamente **não saiu**).

A seleção pega os **próximos N pendentes na ordem da planilha**. Exemplo: um disparo pausado na linha 19 continua, no disparo seguinte, pelas linhas que tinham falhado antes do envio e depois segue da linha 20.

As linhas mostradas (`linha 20 da planilha`) usam a **mesma numeração do Excel** (linha 1 = cabeçalho).

---

## Status possíveis

| Status | Quando | Tenta de novo? |
|---|---|---|
| ✅ Enviado | Enter apertado e a caixa esvaziou | Não (já recebeu) |
| 🧪 Teste enviado | Envio de teste para o seu número | — (não conta como lead) |
| ❌ Número errado | O WhatsApp mostrou "número compartilhado por url é inválido" | Não |
| ❌ Número fora do padrão | O número não é um telefone brasileiro válido | Não |
| ⏭️ Pulado por você | Você escolheu **[2] Pular** na pausa | Não |
| ⚠️ Verificar no WhatsApp | Enter foi apertado, mas não deu para confirmar se saiu | **Não** (evita duplicar — confira manualmente) |
| ⚠️ Erro: conversa não carregou | A conversa não abriu nem com o tempo extra | Sim |
| ⚠️ Erro: mensagem não apareceu | A conversa abriu, mas o texto não apareceu na caixa | Sim |
| ⚠️ Erro: texto dobrado na caixa | Havia rascunho antigo e o texto ficou duplicado mesmo após limpar | Sim |
| ⚠️ Erro: Chrome não respondeu | O Chrome travou antes do Enter | Sim |

Regra de ouro: **só é retentado o que falhou antes do Enter.** Qualquer coisa depois do Enter nunca é reenviada automaticamente.

---

## Pausa (ESC)

| Quando você aperta ESC | O que acontece |
|---|---|
| Aguardando login | Fecha o Chrome e volta ao menu |
| Carregando / aguardando para enviar | Pausa. **[1]** recomeça esse contato do zero (nada tinha saído). **[2]** pula esse contato. **[0]** para. |
| Logo depois do Enter | Termina de registrar o envio e pausa na etapa seguinte |
| Pausa entre contatos | Pausa. **[2]** pula o **próximo** contato (mostrado na tela). |

Depois de **[0] Parar**, o Chrome continua aberto e logado; o próximo disparo não pede QR code.

---

## Proteções

- **Rascunho antigo:** se a caixa tiver o texto duas vezes (rascunho anterior + mensagem nova), o robô limpa e recarrega a conversa uma vez antes de enviar.
- **Freio de erros:** 3 erros seguidos → pausa com `[1] Continuar / [0] Parar`.
- **Login conferido:** o disparo só começa quando a lista de conversas (`#pane-side`) aparece.
- **Aviso de número inválido:** detectado pelo pop-up do WhatsApp (com plano B pela frase exata), para que conversas que contenham a palavra "inválido" não enganem o robô.
- **PC acordado:** `SetThreadExecutionState` impede a suspensão do Windows só enquanto o Sender roda.
- **Chrome em segundo plano:** flags que impedem o Chrome de "congelar" a aba quando minimizado.

---

## Tempo estimado

Cada contato leva em média `15 + 5 + 5 + (20+60)/2 = 65s`. A previsão inicial usa esses timers; a partir do 2º contato, usa a **média real** do disparo. Exemplo: 30 contatos ≈ 32 min.
