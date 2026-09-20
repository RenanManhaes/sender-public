# Segurança e boas práticas

## Risco de banimento do WhatsApp

Automatizar o WhatsApp Web **vai contra os Termos de Uso do WhatsApp**. Disparos em volume podem levar ao bloqueio do número. O Sender reduz o risco, mas não elimina:

- Pausa **aleatória** de 20–60s entre contatos (comportamento menos robótico).
- Você escolhe **quantos** contatos por disparo — prefira lotes (ex.: 30–80 por vez) a centenas seguidas.
- Mensagem personalizada com o nome de cada pessoa.

Recomendações:

- Envie só para quem tem relação com o evento/comunidade (leads que forneceram o contato).
- Se possível, use um número dedicado, não o principal.
- Se muitas pessoas bloquearem ou denunciarem, pare e reavalie a lista/mensagem.

## Dados pessoais

A planilha e o `Registros.md` contêm **nomes e telefones de leads**. Por isso:

- Ficam **fora do repositório**; o `.gitignore` bloqueia `*.xlsx`, `*.xls`, `*.csv`, `*.txt` e `Registros*.md`.
- O repositório é **privado**.
- Nunca cole telefones reais em issues, commits ou na documentação.

## Uso do computador durante o disparo

- O Chrome do robô pode ficar **minimizado**; o computador continua livre (o robô não usa mouse, teclado nem área de transferência).
- **Não clique dentro** do Chrome do robô durante o disparo. Para navegar, use outro Chrome.
- Bloquear a tela (Win+L) não atrapalha. Desligar/reiniciar interrompe — rode `Sender` de novo e ele continua de onde parou.
- Notebook: deixe na tomada.

## Checklist antes de disparar

- [ ] Rodou o `AgoraNao` (com o Sender fechado) para tirar quem respondeu "Agora não"?
- [ ] A **mensagem ativa** (topo do menu) é a certa?
- [ ] A mensagem está atualizada? (evite datas relativas como "faltam 7 dias")
- [ ] O link da mensagem abre?
- [ ] Fez um **teste** para o seu número e conferiu nome, emojis e link no celular?
- [ ] Revisou quem vem na fila? (o menu mostra linha inicial e final; na pausa aparece o próximo)
