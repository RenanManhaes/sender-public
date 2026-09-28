# Configuração

Tudo fica no bloco `CONFIGURAÇÃO`, no topo do `sender.py`.

## Arquivos

| Variável | Padrão | Para que serve |
|---|---|---|
| `ARQUIVO_CAMPANHAS` | `dados/Campanhas.md` | Arquivo editável com as campanhas e mensagens (se existir, manda) |
| `PASTA_PLANILHAS` | `dados/` | Pasta listada no menu **[4] Trocar a planilha** |
| `PLANILHA` | `dados/contatos.xlsx` | Planilha usada ao abrir o Sender |
| `LINHA_CABECALHO` | `0` | Linha do cabeçalho (0 = primeira). Detectada sozinha ao trocar pelo menu |
| `COLUNA_JA_ENVIADA` | `None` | Coluna de controle manual (ex.: "Mensagem enviada?"). **Só recebe quem está com "Não"**; depois de cada envio confirmado o Sender troca para **"Sim" e salva a planilha** (se ela estiver aberta no Excel, avisa e segue; o registro garante que não repete). Ao trocar pelo menu [4], o Sender detecta a coluna e pergunta se deve usá-la |
| `ABA` | `0` | Aba lida (0 = primeira) |
| `COLUNA_NOME` | `"Nome"` | Cabeçalho da coluna do nome usado na mensagem |
| `COLUNA_NUMERO` | `"Telefone"` | Cabeçalho da coluna do telefone |
| `ARQUIVO_REGISTRO` | `dados/Registros.md` | Registro e memória dos envios |
| `PULAR_JA_ENVIADOS` | `True` | Pula quem já está como Enviado no registro |
| `PERFIL_CHROME` | `~/.sender-chrome` | Perfil próprio do Chrome do robô (guarda o login do WhatsApp) |

> As colunas são encontradas **pelo nome do cabeçalho**, não pela letra. Ao trocar de planilha pelo menu [4], o Sender acha sozinho a linha do cabeçalho, a coluna de nome (prefere "Nome", nunca "Sobrenome") e a de telefone (entre as que parecem telefone — Whatsapp, Celular, Telefone… — fica com a que tem **mais números válidos**). Você confirma na prévia ou escolhe as colunas à mão.

> **Nome:** nome já tratado (uma palavra, ex.: "Maria") vai como está; nome cru ("MARIA APARECIDA", "joão silva") vira só o primeiro nome com inicial maiúscula.

## Timers (segundos)

| Variável | Padrão | Etapa |
|---|---|---|
| `TEMPO_LOGIN_QRCODE` | `120` | Espera pelo QR code (segue sozinho quando loga) |
| `TEMPO_BUSCAR` | `15` | Carregando a conversa |
| `TEMPO_BUSCAR_EXTRA` | `30` | Tempo extra se a conversa não carregou (internet lenta) |
| `TEMPO_ANTES_ENTER` | `5` | Aguardando para enviar |
| `TEMPO_TEXTO_EXTRA` | `10` | Tempo extra se o texto não apareceu na caixa |
| `TEMPO_APOS_ENTER` | `5` | Enviando a mensagem (antes de conferir e trocar de página) |
| `TEMPO_ENTRE_CONTATOS` | `(60, 180)` | Pausa aleatória entre um contato e outro (1 a 3 minutos) |
| `LIMITE_ERROS_SEGUIDOS` | `3` | Erros seguidos até o freio automático |

Os tempos "extra" só são usados quando necessário e terminam antes se a condição for atendida.

## Mensagem

`MENSAGENS` — as mensagens disponíveis (exemplos: `convite` e `convite_imagem`). Cada uma tem `nome` e `texto`; use `{Nome}` onde entra o nome. Pode ter emojis e quebras de linha.

`MENSAGEM_ATIVA` — a mensagem usada ao abrir o Sender. Durante o uso, troque pelo menu **[3] Trocar a mensagem** (mostra a prévia de cada uma).

A memória de quem já recebeu é **separada por mensagem**: quem recebeu uma mensagem ainda pode receber outra. Falhas definitivas (número errado, fora do padrão, pulado, verificar) valem para todas as mensagens.

**Mensagem com imagem:** acrescente `"imagem": "caminho\\da\\imagem.jpg"` na mensagem. O Sender abre a conversa com o texto na caixa, espera o texto aparecer, anexa a imagem por **Anexo > Fotos e vídeos** (nunca pelo campo de figurinha) e clica em "Enviar 1 item selecionado": o texto da caixa vai junto como legenda. Use JPG/PNG (WEBP pode virar figurinha). A confirmação do envio procura na conversa um trecho exclusivo do texto (a linha mais longa sem link).

`LINK_BASE` — formato do link do WhatsApp Web. Não precisa mudar.

## Campanhas

> ⚠️ Se existir o arquivo **`Campanhas.md`** na pasta das planilhas, ele substitui `MENSAGENS` e `CAMPANHAS` do código (veja o README). O que está aqui embaixo vale como referência e como valor padrão quando o arquivo não existe.

`CAMPANHAS` — o que aparece no menu **[3] Trocar de campanha**. Cada campanha tem:

| Campo | Exemplo | Para que serve |
|---|---|---|
| `nome` | `"Convite com imagem"` | Nome no menu |
| `publico` | `"lista de contatos"` | Para quem é (aparece ao lado do nome) |
| `planilha` | `"contatos.xlsx"` | Arquivo dentro de `PASTA_PLANILHAS`; cabeçalho e colunas são detectados sozinhos |
| `mensagem` | `"convite_imagem"` | Chave em `MENSAGENS` |
| `excluir` | `["clientes.xlsx"]` | Planilhas (na mesma pasta) cujos números **nunca** recebem esta campanha |
| `usar_controle` | `True` | Só recebe quem está "Não" na coluna de controle e marca "Sim" após o envio |
| `segmentar` | `"É membro?"` | Coluna que divide o público. Ao ativar a campanha, o Sender lista os valores da coluna com a contagem de contatos e pergunta para qual enviar (ou "Todos"). O público escolhido aparece no cartão do menu e no `Registros.md` |

`CAMPANHA_INICIAL` — a campanha que já vem escolhida ao abrir o Sender.

## Trocar de campanha

1. Acrescente a nova mensagem em `MENSAGENS` (com uma chave nova) e ative pelo menu [3].
2. Aponte `PLANILHA` para a nova lista (ou substitua o arquivo).
3. **Registro:** o Sender pula todo número que já aparece como Enviado no `Registros.md`. Para uma campanha nova em que as mesmas pessoas devem receber de novo, aponte `ARQUIVO_REGISTRO` para um arquivo novo (ex.: `Registros - Campanha X.md`).
