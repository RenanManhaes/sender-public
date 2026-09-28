# 📣 Campanhas do Sender

> Renomeie este arquivo para **`Campanhas.md`** (na mesma pasta) para ele valer.
> O que estiver aqui substitui as campanhas e mensagens escritas no `sender.py`.
> O menu **[3] Trocar de campanha** relê o arquivo, então dá para editar com o Sender aberto.

**Como preencher cada bloco:**

| Linha | Obrigatória? | Para que serve |
|---|---|---|
| `## Nome da campanha` | sim | Nome que aparece no menu |
| `id:` | recomendado | Identificador da mensagem. É por ele que o Sender lembra quem já recebeu — **não mude depois de disparar**. Duas campanhas podem usar o mesmo `id` (mesma mensagem, públicos diferentes) |
| `publico:` | não | Para quem é (aparece ao lado do nome) |
| `planilha:` | **sim** | Arquivo desta pasta. Cabeçalho e colunas são detectados sozinhos |
| `imagem:` | não | Foto enviada junto, com o texto como legenda (JPG ou PNG) |
| `excluir:` | não | Planilhas desta pasta cujos números **nunca** recebem esta campanha, ex.: `excluir: clientes.xlsx` (separe por vírgula) |
| `controle:` | não | `sim` = só envia para quem está "Não" na coluna de controle, e marca "Sim" depois |
| `mensagem:` | **sim** | Tudo o que vier depois vira o texto. `{Nome}` = primeiro nome |

Negrito do WhatsApp: `*assim*`.

---

## Convite
id: convite
publico: lista de contatos
planilha: contatos.xlsx
mensagem:
Olá {Nome}! Tudo bem?

Estou passando para te convidar para o nosso próximo encontro.

Se quiser saber mais, é só acessar:
🔗 https://exemplo.com.br/evento

Espero encontrar você lá!

---

## Convite com imagem
id: convite_imagem
publico: quem ainda não confirmou
planilha: contatos.xlsx
imagem: imagem-exemplo.jpg
controle: sim
mensagem:
Oi, {Nome}! Tudo bem?

Preparamos uma novidade para você: o site oficial do evento já está no ar.

*Garanta sua vaga:*
👉 https://exemplo.com.br/evento

Nos vemos lá! 🚀
