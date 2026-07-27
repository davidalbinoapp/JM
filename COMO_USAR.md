# Jornal Mural (AGA) — sistema completo, pronto pra usar

## O que você tem em mãos

Um conjunto de arquivos que juntos automatizam a montagem do JM:

- **`auto_jm.py`** — o atalho do dia a dia. Recebe **só a pasta da semana** e
  monta tudo (acha o briefing e casa os QR codes sozinho). É o que você roda.
- **`run_jm.py`** — o script "manual" por trás, usado quando você precisa
  informar um QR ou o nome de saída na mão.
- **`jm_engine.py`** — o motor (diagramação, corte e enquadramento de foto,
  boxes, página final).
- **`briefing_parser.py`** — lê o Word e entende a estrutura.
- **`template_compactado.pptx`** — o modelo visual (capa, matérias, cartazes, página final).
- **`tags_editorias/`** — as 11 tags de editoria + a logo de fechamento.
- **`fonts/`** — a Liberation Sans (reserva da Arial).
- **`models/`** — o modelo de detecção de rosto (enquadramento inteligente).
- **`requirements.txt`** + **`setup_mac.command`** / **`setup_windows.bat`** — a instalação automática.

Esses arquivos não mudam toda semana — só o briefing e as fotos mudam.

## Instalação (uma vez só, com dois cliques)

Não precisa mexer no terminal. Escolha o do seu sistema:

- **Mac** → dois cliques em **`setup_mac.command`**
  - Se o Mac avisar "desenvolvedor não identificado": clique com o botão
    direito no arquivo → **Abrir** → **Abrir**.
- **Windows** → dois cliques em **`setup_windows.bat`**

O instalador confere e instala tudo sozinho: as bibliotecas do Python
(`python-pptx`, `python-docx`, `Pillow`, `pymupdf`), o **LibreOffice**
(usado só pra montar a página final de miniaturas) e o **OpenCV** (opcional,
pro enquadramento inteligente de fotos). No fim ele mostra um
✅ **TUDO PRONTO** — ou aponta exatamente o que faltou.

Sobre a **fonte Arial**: você quase nunca vai precisar instalar à mão. Mac
e Windows já vêm com Arial, e o sistema ainda traz a Liberation Sans (mesmas
medidas da Arial) na pasta `fonts/` como rede de segurança. Se um dia a Arial
faltar num computador, é só instalar a fonte por lá — mas na prática isso
raramente acontece.

Se o LibreOffice não for instalado, o JM ainda é gerado normalmente — só a
última página (as miniaturas) fica com a versão genérica do template, e
aparece um aviso dizendo como resolver.

## Passo 1 — Organização das pastas (só uma vez)

Deixe o sistema (a pasta `JM-main`) num lugar fixo e as **pastas de semana ao
lado dela** — uma por edição. Cada pasta de semana tem o briefing `.docx`, as
fotos, os cartazes e os QR codes daquela edição:

```
JM SYSTEM/
├── JM-main/        ← o sistema (não muda toda semana)
├── JM 21-07/       ← uma edição: briefing + fotos + QRs
├── JM 30-06/       ← outra edição
└── ...
```

Como backup e pra usar em outra máquina, mantenha a `JM-main` num **repositório
GitHub grátis** (pode ser público, não tem nada sensível): sempre que um
arquivo do sistema mudar, suba a versão nova.

## Passo 2 — Rodar toda semana (Terminal)

Coloque **todas as fotos e QR codes na pasta da semana ANTES de rodar** (o
motor lê a pasta no instante em que roda). Depois é **um comando só**, passando
só a pasta:

```
cd ~/Documents/Artes/AGA/2026/JM\ SYSTEM/JM-main && python3 auto_jm.py "../JM 21-07"
```

O `auto_jm.py` acha o briefing `.docx`, encontra os QR codes pelo nome do
arquivo e casa cada um com a matéria certa, mostra o que detectou e já gera o
`AGA_<pasta>-<ddmmaa>.pptx` dentro da própria pasta da semana.

- Pra só **conferir a detecção sem gerar nada**, acrescente `--dry-run` no fim.
- Se um QR estiver com nome muito diferente do título e não casar, o
  `auto_jm.py` **avisa** — aí use o comando manual abaixo pra informá-lo.

### Comando manual (quando precisar informar o QR na mão)

```
python3 run_jm.py "../PASTA/BRIEFING.docx" "../PASTA" --qr 'Título exato da matéria=../PASTA/arquivo_qr.png'
```

Use **aspas simples** nos `--qr`: se o título tiver `!`, aspas duplas dão erro
no terminal. O lado esquerdo do `=` é o **título da matéria** (como está no
briefing), não o nome do arquivo do QR.

### Plano B — num computador sem o sistema instalado (claude.ai)

Dá pra rodar pelo claude.ai: abra uma conversa nova, anexe o briefing + as
fotos + os QRs, e peça pra baixar os arquivos do seu repositório GitHub e rodar
o `montar_jornal_mural`. É o caminho de emergência; o normal é rodar pelo
Terminal com o `auto_jm.py`.

## O que o sistema já resolve sozinho

- Lê o Word e identifica capa, matéria única, matéria dupla e cartazes —
  inclusive quando o Word não tem um cabeçalho "Cartazes:" explícito.
- Casa cada nome de foto do Word com o arquivo real, tolerando pequenas
  diferenças de acento/espaço no nome do arquivo.
- Reconhece "UNIR ESSA: ... e ESSA: ..." e junta as fotos automaticamente,
  com o corte diagonal aplicado no conjunto todo.
- Ajusta a quantidade de slides de cartaz pro número real daquela semana.
- Aplica toda a diagramação fina: tag alinhada com a foto, título nunca
  quebra virando viúva, fonte do corpo se ajusta pra nunca sobrepor a
  próxima matéria, título encolhe quando o QR reduz o espaço, corpo nunca
  passa da base da própria foto, logo de fechamento só na matéria que
  realmente for a última.
- **Enquadra as pessoas na foto** automaticamente: detecta os rostos e desloca
  o corte pra não cortar ninguém (ver seção mais abaixo).
- Monta **boxes de informação** (horários, contatos) quando o briefing pede
  (ver seção mais abaixo).
- Regenera a página final com as miniaturas certas dessa edição.

## Foto que não pode ser cortada (banner, telão, arte com texto)

Às vezes a foto tem conteúdo importante nas bordas (um banner, um telão com
texto) e o corte automático comeria as laterais. Nesses casos, escreva na
linha da foto do briefing, depois do nome do arquivo:

```
Foto: NOME DO ARQUIVO.png (fundo desfocado)
```

Aí a foto entra **inteira**, centralizada, e o espaço que sobra é preenchido
com uma versão ampliada e desfocada dela mesma — igual ao acabamento feito à
mão. Também funciona escrevendo "(foto inteira)", "(sem cortar)" ou
"(banner)". Use só quando precisar; a maioria das fotos fica melhor com o
corte normal.

## Enquadramento inteligente pelo ASSUNTO (automático)

Quando uma foto é mais larga que o espaço dela na arte, o motor precisa cortar
as laterais. Em vez de cortar sempre pelo **centro** (que deixava gente pela
metade, ou jogava o assunto pra fora), o motor agora enquadra pelo **assunto da
foto**, combinando dois sinais:

- **Rostos:** detecta as pessoas e desloca o corte pra mantê-las enquadradas.
- **Saliência visual:** identifica onde está a "energia" da imagem (o palestrante
  isolado num palco, o globo brilhante, o objeto em destaque) e **ignora padrões
  repetitivos** (a textura de cadeiras/nucas de uma plateia). É isso que impede
  a plateia de "sequestrar" o corte e faz o corte achar o assunto mesmo numa
  foto **sem rosto**.

Não precisa marcar nada no briefing — é tudo automático.

- Em foto **sem assunto claro** (prédio de fachada cheia, logo/gráfico já
  centrado), o motor **mantém o centro** — não descentraliza à toa.
- Depende da biblioteca **OpenCV** (rostos) — instalada automaticamente pelos
  instaladores. A saliência roda em **numpy puro**, sem dependência extra. Se o
  OpenCV faltar, o JM é gerado normal (só sem o ajuste por rosto).
- O modelo de detecção (`models/face_detection_yunet.onnx`, ~230 KB) vai junto
  no repositório, como as tags e as fontes.

### Quando quiser forçar o enquadramento à mão (`FOCO:`)

O automático acerta na quase totalidade dos casos, mas se um corte específico
sair errado, dá pra mandar o motor à mão. Escreva na linha da foto, ou numa
linha `Foco:` própria logo abaixo:

```
Foto: NOME DO ARQUIVO.jpg (foco: direita)
```

ou

```
Foto: NOME DO ARQUIVO.jpg
Foco: direita
```

Valores aceitos: **esquerda / centro / direita** (corte horizontal) e
**cima / baixo / centro** (corte vertical):

- **esquerda / direita / cima / baixo** → empurra o corte pra aquele lado
  (encosta naquela borda). Use quando o assunto está bem num canto.
- **centro** → centraliza no **assunto** (as pessoas / a parte importante),
  não no centro geométrico cego da foto. Na prática é o mesmo que deixar o
  automático agir — serve pra "resetar" um enquadramento que você queira
  garantir centralizado no que importa.

Um valor que não faz sentido pro corte daquela foto (ex.: `cima` numa foto que
só corta as laterais) é ignorado sem quebrar.

## Box de informações (horários, contatos, telefones)

Quando uma matéria precisa destacar um bloco de informações — horários de
funcionamento, telefones, um "serviço" — dá pra virar **box**: aqueles blocos
com o rótulo em **laranja** e uma **linha divisória**, separados do texto
corrido (igual ao modelo do CEA / Centro de Memória).

No briefing, escreva o texto normal da matéria e, quando chegar nos boxes,
comece a linha com **`ABRIR BOX COM ESSAS INFORMAÇÕES:`** e coloque **cada box
em um parágrafo**, no formato **`Rótulo: conteúdo`**:

```
Visite o CEA e o Centro de Memória nas férias
O Centro de Educação Ambiental oferece contato com a natureza... Confira os horários e programe sua visita!
ABRIR BOX COM ESSAS INFORMAÇÕES: CEA: quarta a sexta-feira, das 8h30 às 11h30... Informações: (31) 97199-8901.
Centro de Memória: quarta a sexta-feira, das 8h às 12h... WhatsApp (31) 97200-7978.
```

Isso vira, automaticamente:

- O texto **antes** do `ABRIR BOX...` fica como corpo normal da matéria.
- A instrução `ABRIR BOX COM ESSAS INFORMAÇÕES:` **não aparece** na arte — é só
  o gatilho.
- Cada parágrafo seguinte vira **um box**: o que vem antes dos dois-pontos
  (`CEA`, `Centro de Memória`) fica **laranja em negrito**, e o resto em preto,
  com as linhas divisórias no lugar.

Detalhes que ajudam:

- **Um box por parágrafo** (uma quebra de linha entre eles). É a quebra que
  separa um box do outro.
- O **rótulo é a parte antes do primeiro `:`** — mantenha curto (ex.: "CEA:",
  "Centro de Memória:", "Atendimento:"). Os outros dois-pontos do conteúdo
  (ex.: "Informações:") continuam normais, em preto.
- O layout pronto do cliente comporta **2 boxes**. Se você mandar mais que
  isso, o motor junta o excedente no último box e avisa — então o ideal é
  manter até 2.

## O que ainda depende de você

- **Baixar as fotos da semana** do SharePoint pra uma pasta — não tenho
  acesso à sua conta.
- **Avisar qual QR code é de qual matéria** — o nome do arquivo sozinho não
  deixa isso claro o suficiente pra eu adivinhar com segurança.
- **Revisar o resultado antes de publicar** — o motor erra bem menos hoje,
  mas a palavra final de aprovação é sempre sua.

## Se aparecer um caso novo que o motor não reconhece

Me mostra o print e o texto do briefing, do jeito que fizemos aqui — eu
ajusto o motor e te devolvo os arquivos atualizados pra você subir de novo
no repositório, substituindo os antigos.
