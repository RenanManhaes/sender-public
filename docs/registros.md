# Registros na pasta `dados/`

O Sender grava cada envio em `dados/Registros.md`. Esse arquivo é ao mesmo tempo o **histórico** para você consultar e a **memória** do robô.

## Formato

Uma seção por teste ou disparo, a mais recente no final:

```markdown
---

## 🚀 Disparo oficial — 01/01/2026 às 10:00

- **Planilha:** `contatos.xlsx`
- **Mensagem:** Convite (exemplo) (`convite`)
- **Contatos selecionados:** 30

| Nº | Nome | Número | Status | Horário |
|---:|---|---|---|---|
| 2 | Fulano | 5511900001111 | ✅ Enviado | 01/01 10:01:12 |
| 5 | Ciclano | 13055550123 | ❌ Número fora do padrão | 01/01 10:02:40 |

> **Resumo:** ✅ Enviado: **1** · ❌ Número fora do padrão: **1** · ⏱ 17min · 🛑 **interrompido** às 19:09
```

- **Nº** = linha da planilha no Excel (testes aparecem como `Teste`).
- **Mensagem** = qual mensagem foi usada nesse disparo. A memória de quem já recebeu é separada por mensagem; seções sem essa linha contam como a mensagem padrão.
- Cada linha é gravada **na hora** do envio: se o PC desligar no meio, o que foi feito fica registrado.
- O resumo diz `finalizado` ou `interrompido` (ESC → Parar, Ctrl+C ou Chrome travado).
- A nota pode ficar aberta na pasta `dados/` durante o disparo; ela atualiza sozinha.

## Regras

- ⚠️ **Não edite as linhas das tabelas** (principalmente Número e Status). O Sender lê essas colunas para decidir quem já recebeu. Anotações fora das tabelas não atrapalham.
- Linhas de **teste** são ignoradas na memória (seu número nunca conta como lead).
- Plugins que reformatam tabelas (ex.: Advanced Tables) podem mudar o alinhamento sem problema, desde que o conteúdo das células continue igual.

## Como a memória lê o arquivo

Para cada linha de tabela cuja 1ª coluna é um número:

- Status `Enviado` → número vai para **enviados** (e vira o "Último enviado" do menu).
- Status em `ERROS_RETENTAVEIS` → ignorado (o contato volta para a fila).
- Qualquer outro status → número vai para **falhas definitivas** (nunca mais é tentado).

## Outras notas do Sender na pasta `dados/`

| Nota | Para que serve |
|---|---|
| `dados/Agora não.md` | Registro acumulado de quem respondeu "Agora não" (gerado pelo verificador; ver [verificador-agora-nao.md](verificador-agora-nao.md)) |
| `dados/Nunca contatar.md` | Números que nenhum robô usa. Editável à mão: um número por linha |
