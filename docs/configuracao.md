# Configuração

Tudo fica no bloco `CONFIGURAÇÃO`, no topo do `sender.py`.

## Arquivos

| Variável | Padrão | Para que serve |
|---|---|---|
| `PLANILHA` | `dados/contatos.xlsx` | Planilha de contatos |
| `ABA` | `0` | Aba lida (0 = primeira) |
| `COLUNA_NOME` | `"Nome"` | Cabeçalho da coluna do nome usado na mensagem (vai como está) |
| `COLUNA_NUMERO` | `"Telefone"` | Cabeçalho da coluna do telefone |
| `ARQUIVO_REGISTRO` | `...\dados\Registros.md` | Registro e memória dos envios |
| `PULAR_JA_ENVIADOS` | `True` | Pula quem já está como Enviado no registro |
| `PERFIL_CHROME` | `~/.sender-chrome` | Perfil próprio do Chrome do robô (guarda o login do WhatsApp) |

> As colunas são encontradas **pelo nome do cabeçalho**, não pela letra. Se renomear o cabeçalho na planilha, atualize aqui.

## Timers (segundos)

| Variável | Padrão | Etapa |
|---|---|---|
| `TEMPO_LOGIN_QRCODE` | `120` | Espera pelo QR code (segue sozinho quando loga) |
| `TEMPO_BUSCAR` | `15` | Carregando a conversa |
| `TEMPO_BUSCAR_EXTRA` | `30` | Tempo extra se a conversa não carregou (internet lenta) |
| `TEMPO_ANTES_ENTER` | `5` | Aguardando para enviar |
| `TEMPO_TEXTO_EXTRA` | `10` | Tempo extra se o texto não apareceu na caixa |
| `TEMPO_APOS_ENTER` | `5` | Enviando a mensagem (antes de conferir e trocar de página) |
| `TEMPO_ENTRE_CONTATOS` | `(20, 60)` | Pausa aleatória entre um contato e outro |
| `LIMITE_ERROS_SEGUIDOS` | `3` | Erros seguidos até o freio automático |

Os tempos "extra" só são usados quando necessário e terminam antes se a condição for atendida.

## Mensagem

`MENSAGENS` — as mensagens disponíveis (hoje: `evento` = convite para o evento e `grupo` = convite para o grupo do WhatsApp). Cada uma tem `nome` e `texto`; use `{Nome}` onde entra o nome. Pode ter emojis e quebras de linha.

`MENSAGEM_ATIVA` — a mensagem usada ao abrir o Sender. Durante o uso, troque pelo menu **[3] Trocar a mensagem** (mostra a prévia de cada uma).

A memória de quem já recebeu é **separada por mensagem**: quem recebeu o convite do evento ainda pode receber o do grupo. Falhas definitivas (número errado, fora do padrão, pulado, verificar) valem para todas as mensagens.

`LINK_BASE` — formato do link do WhatsApp Web. Não precisa mudar.

## Trocar de campanha

1. Acrescente a nova mensagem em `MENSAGENS` (com uma chave nova) e ative pelo menu [3].
2. Aponte `PLANILHA` para a nova lista (ou substitua o arquivo).
3. **Registro:** o Sender pula todo número que já aparece como Enviado no `Registros.md`. Para uma campanha nova em que as mesmas pessoas devem receber de novo, aponte `ARQUIVO_REGISTRO` para um arquivo novo (ex.: `Registros - Campanha X.md`).
