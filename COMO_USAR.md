# Jornal Mural (AGA) — sistema completo, pronto pra usar

## O que você tem em mãos

Um conjunto de arquivos que juntos automatizam a montagem do JM:

- **`Montar JM.command`** (Mac) / **`Montar JM (Windows).bat`** (Windows) — o
  atalho de **dois cliques**: abre uma janela pra você escolher a pasta da
  semana e já monta o JM (sem digitar caminho). É o jeito mais fácil de rodar no
  dia a dia. Use o que combina com o seu computador.
- **`auto_jm.py`** — o motor do atalho. Recebe **só a pasta da semana** e monta
  tudo (acha o briefing e casa os QR codes sozinho). É o que o `Montar JM`
  chama por baixo — dá pra rodar direto pelo Terminal também.
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

## Passo 2 — Rodar toda semana

Coloque **todas as fotos e QR codes na pasta da semana ANTES de rodar** (o
motor lê a pasta no instante em que roda). Depois, escolha um dos dois jeitos:

### Jeito fácil — dois cliques (sem digitar caminho)

Dê **dois cliques** no atalho da sua máquina (dentro da `JM-main`):

- **No Mac:** `Montar JM.command`
- **No Windows:** `Montar JM (Windows).bat`

Abre uma janela pra **escolher a pasta da semana** — clique nela e confirme.
Pronto: o motor monta o `.pptx` dentro da própria pasta e mostra o resultado ali
mesmo. Nenhum caminho pra digitar, nada pra decorar.

- **Mac, 1ª vez:** o macOS pode dizer "desenvolvedor não identificado" —
  clique com o **botão direito** no `Montar JM.command` → **Abrir** → **Abrir**.
  Depois disso, dois cliques normais funcionam sempre.
- **Atalho com ícone da marca — nasce sozinho na instalação.** Quando você
  roda o `setup_mac.command` (Mac) ou `setup_windows.bat` (Windows), no fim ele
  já cria o atalho com o ícone da Nossa Voz, sem nenhum passo extra:
  - **Mac:** um **`Montar JM.app`** na pasta "JM SYSTEM" (ao lado das pastas de
    semana). Arraste pro Dock se quiser. O ícone fica dentro do app, então viaja
    junto se você copiar o app.
  - **Windows:** um atalho **"Montar JM"** na Área de Trabalho.
  - Se algum dia quiser recriar o app do Mac (ex.: trocou o ícone em
    `icone_montar_jm.png`), rode o **`Criar atalho (Mac).command`**.
- **Windows, 1ª vez:** pode aparecer um aviso azul "O Windows protegeu o
  computador" — clique em **Mais informações** → **Executar assim mesmo**.
- Dica: pra ter o atalho à mão, faça um **alias/atalho** dele na Área de
  Trabalho (Mac: botão direito → "Criar alias"; Windows: botão direito →
  "Enviar para → Área de trabalho (criar atalho)").

### Jeito Terminal (alternativa)

Se preferir o Terminal, é **um comando só**, passando só a pasta:

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
  com o corte diagonal aplicado no conjunto todo (e dá pra dar um foco
  **por foto** da união — ver a seção do `FOCO:`).
- Ajusta a quantidade de slides de cartaz pro número real daquela semana.
- **QR code na matéria dupla E na página única**: o motor casa o QR pela
  semelhança do nome do arquivo com o título da matéria (você não informa
  nada — só deixa o QR na pasta). Na dupla, o QR vai no canto superior
  direito; na **página única**, ele entra no **canto inferior esquerdo da
  foto**, na posição que está desenhada no molde do template.
- **Descobre sozinho o papel de cada slide-modelo do template** (capa, dupla,
  única, única-com-QR, cartaz, página final) pela cara de cada um — então dá
  pra **incluir ou reordenar páginas** no `template_compactado.pptx` sem
  quebrar nada: nenhum número de slide fica preso no código.
- Aplica toda a diagramação fina: tag alinhada com a foto (e com um respiro
  da foto na capa, pra editoria comprida não encostar), **título nunca termina
  com uma palavra sozinha** (viúva) e **as duas últimas linhas do corpo ficam
  equilibradas** (desce quantas palavras forem necessárias pra a última linha
  não ficar vazia demais, sem virar a maior), fonte do corpo se ajusta pra
  nunca sobrepor a próxima matéria, título encolhe quando o QR reduz o espaço,
  corpo nunca passa da base da própria foto, logo de fechamento só na matéria
  que realmente for a última.
- **Aproveita o espaço do título da capa**: se o título quebra em N linhas mas
  uma redução pequena da fonte (até ~12%) o encaixa em uma linha a menos, com
  linhas mais cheias, o motor reduz só o necessário (nunca cresce). Título que
  já quebra bem, ou que só caberia com fonte bem menor, fica no tamanho do
  template. O corpo se reposiciona sozinho embaixo.
- **Reequilibra a quebra do título** (capa, dupla e única): se a quebra
  automática deixaria uma linha curta no meio do título (ex.: "CMG" sozinho, ou
  "voto é pela" num buraco), o motor redistribui as palavras pra deixar as
  linhas mais parelhas, **sem mudar o número de linhas e sem mexer na largura da
  caixa** (continua alinhada com o QR/foto). Título que já quebra bem fica como
  está.
- **Aperta um tico a caixa do corpo pra fechar a última linha**: quando o
  parágrafo terminaria com uma linha final curta (ex.: "equipes premiadas!"
  sozinho, com um "às" pendurado na linha de cima), o motor estreita levemente a
  caixa do corpo (no máximo ~8%) pra o texto refluir e a frase de fecho cair
  inteira numa linha — a mesma decisão que se toma no olho, fechando um
  pouquinho a caixa. Só age quando existe mesmo uma última linha curta; corpo
  que já fecha bem fica intacto. Não muda o nº de linhas (nada de transbordo).
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

- **esquerda / direita / cima / baixo** → move o **assunto** pra aquele lado do
  quadro, na direção da seta: `esquerda ⟵`, `direita ⟶`, `cima ⬆`, `baixo ⬇`.
  Sem número, vai até o limite (o assunto encosta naquele lado).
- **centro** → centraliza no **assunto** (as pessoas / a parte importante),
  não no centro geométrico cego da foto. Na prática é o mesmo que deixar o
  automático agir — serve pra "resetar" um enquadramento.

**Cada foto só corta num eixo.** Foto mais larga que o espaço é aparada nas
**laterais** (aí valem `esquerda`/`direita`); foto mais alta é aparada em
**cima/baixo** (aí valem `cima`/`baixo`). Se você pedir o eixo errado (ex.:
`cima` numa foto aparada nas laterais), o motor **avisa** ao gerar que aquele
FOCO ficou sem efeito — é só trocar pelo par certo.

**Ajuste fino com porcentagem (opcional):** em vez de ir direto pra borda, dá
pra empurrar só um pouquinho, escrevendo um número de 0 a 100 depois da direção
(o `%` é opcional):

```
Foto: NOME DO ARQUIVO.jpg (foco: direita 20%)
```

A porcentagem é **relativa ao enquadramento automático**: `direita 20%` pega o
corte que o motor já achou de bom e dá um empurrão de 20% do caminho rumo à
borda direita. `direita` sem número = 100% (vai até a borda, como antes). Vale
pros quatro lados: `esquerda 15%`, `baixo 30%`, etc. Aceita decimal
(`direita 22,5%`) e nunca deixa o corte sair da foto (no máximo encosta na
borda). Um valor sem sentido pro corte daquela foto (ex.: `cima` numa foto que
só corta as laterais) é ignorado sem quebrar.

**Ajustar os DOIS eixos ao mesmo tempo (foco 2D):** como cada foto corta só num
eixo, normalmente só um lado (ou só cima/baixo) tem efeito. Se você quiser
mexer **na horizontal E na vertical juntas**, escreva os dois focos na mesma
linha:

```
Foto: NOME DO ARQUIVO.jpg (foco: direita 25% baixo 30%)
```

Quando o motor vê um foco horizontal **e** um vertical, ele **aproxima a foto
um pouco** (um zoom automático discreto) só pra criar folga nos dois lados, e aí
posiciona o assunto nos dois eixos de uma vez. Isso é **opt-in**: só liga com
dois focos juntos — com **um foco só** (ou nenhum) nada muda, o corte é
exatamente o de sempre. O preço do 2D é fechar um pouquinho no assunto (mostra
um tico menos de fundo) — pra capa costuma até ajudar. As porcentagens de cada
eixo funcionam igual (relativas ao automático); a ordem tanto faz
(`baixo 30% direita 25%` = `direita 25% baixo 30%`). Você **não** controla o
zoom no briefing — ele é decidido pelo motor.

**Foco POR FOTO numa união (`UNIR ESSA: ...`):** quando você junta duas (ou
mais) fotos, cada uma vira uma fatia e é enquadrada **sozinha**. Às vezes o
automático acerta uma e erra a outra (ex.: uma foto de prédio com jardim
colorido — a saliência puxa pro jardim e corta o prédio). Pra corrigir, você
gruda o foco **em cada foto**, dentro do próprio `UNIR ESSA`:

```
Foto: UNIR ESSA: FOTO_A.jpg E ESSA: FOTO_B.jpg (foco: direita)
```

ou, mexendo nas duas:

```
Foto: UNIR ESSA: FOTO_A.jpg (foco: esquerda) E ESSA: FOTO_B.jpg (foco: direita 50%)
```

Cada foto respeita o **seu** foco; a foto sem foco fica no automático. É a mesma
sintaxe de sempre (esquerda/direita/cima/baixo, com `%`), só que agora **por
foto**. União **sem** foco = exatamente o corte de antes. Um nome de arquivo com
parênteses (ex.: `... (2).jpg`) **não** é confundido com foco — o marcador exige
a palavra `foco:` dentro dos parênteses.

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
