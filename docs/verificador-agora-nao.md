# Verificador "Agora não"

Ferramenta **paralela** ao Sender (`verificar_agora_nao.py`, comando `AgoraNao`). Ela não envia mensagens. Serve para tirar da planilha quem já respondeu **"Agora não"** a um convite anterior, antes de um disparo.

## Por que existe

Campanhas enviadas pela API oficial do WhatsApp costumam usar botões de resposta rápida, por exemplo **"Quero receber"** e **"Agora não"**. Quem respondeu "Agora não" tem mais chance de denunciar o número se receber outra mensagem. O verificador faz o levantamento pela busca do WhatsApp Web.

## Como usar

1. **Feche o Sender** (opção `[0] Sair`). Os dois usam o mesmo perfil do Chrome e o Chrome não deixa abrir o mesmo perfil duas vezes.
2. No cmd:
   ```
   AgoraNao
   ```
3. Aperte Enter e responda se o WhatsApp já está logado (**[1]** sim / **[0]** não).
4. O robô:
   1. pesquisa `agora não` na barra de pesquisa;
   2. detecta sozinho o painel de resultados que rola e **desce até o fim**, insistindo algumas vezes no final (a busca carrega em lotes);
   3. guarda só quem mandou **exatamente** "Agora não" (mensagens que apenas contêm as palavras ficam de fora);
   4. cruza com a planilha **pelo telefone**;
   5. atualiza o registro na pasta `dados/`;
   6. pergunta se pode mover os encontrados para a aba **"Não enviar"**.
5. **ESC** interrompe a coleta e segue com o que já foi encontrado.

## O que ele garante

- **Somente leitura no WhatsApp:** não abre conversas, não clica em contatos, não envia nada.
- **Backup antes de mexer na planilha:** `contatos - backup AAAA-MM-DD HHMM.xlsx`, na mesma pasta.
- **Conferência linha a linha:** antes de tirar uma linha, confere se o telefone dela é o esperado.
- **Nada é apagado:** as linhas vão para a aba "Não enviar" com as colunas originais, mais **Motivo** e **Movido em**. O Sender só lê a primeira aba.
- **Planilha aberta no Excel:** ele avisa e não mexe.

## Registro na pasta `dados/` — `dados/Agora não.md`

Registro **acumulado**: quem já foi registrado nunca sai da lista; cada verificação só acrescenta os novos.

```markdown
| Contato no WhatsApp | Número | Na planilha? | Visto em |
|---|---|---|---|
| +55 15 99123-4567 | 5515991234567 | ✅ linha 98 (Fulano) | 19/09/2026 |
| +55 11 98765-4321 | 5511987654321 | — | 19/09/2026 |
| Contato Salvo | — | ❓ conferir | 19/09/2026 |
```

- O cruzamento com a planilha usa **a lista acumulada inteira**. Por isso, quem apareceu numa verificação antiga continua sendo tirado de planilhas novas.
- **Contatos salvos na agenda** aparecem pelo nome, não pelo número, e ficam como `❓ conferir`.

## Nunca contatar — `dados/Nunca contatar.md`

Números nessa nota ficam fora de **tudo**: o Sender nunca envia para eles e o verificador não os registra. Um número por linha, em qualquer formato, com comentário opcional depois.

## Limite importante: histórico do WhatsApp Web

A busca do WhatsApp Web só enxerga mensagens **sincronizadas para o navegador**. Um aparelho conectado recentemente recebe apenas o histórico recente: em uso real, a busca alcançou pouco mais de duas semanas de respostas, e nem a rolagem com a rodinha do mouse carregou mais do que isso.

Consequências:

- Respostas "Agora não" **anteriores ao histórico sincronizado** não aparecem.
- Para esse período, a fonte completa é a **plataforma que enviou a campanha**: o histórico de execução do workflow ou a API de conversas.
- Para conferir se há respostas antigas: pesquise "agora não" no **celular** e role até o fim.

## Diagnóstico

Se aparecer `Não encontrei a lista de resultados na tela` ou `Não encontrei o painel de resultados que rola`, o WhatsApp mudou o layout da busca. Tire um print com o F12 aberto (aba Elements, sobre um resultado) para ajustar os seletores em `JS_PAINEL` / `JS_LER_RESULTADOS`.
