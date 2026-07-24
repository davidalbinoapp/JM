# Jornal Mural (AGA) — sistema completo, pronto pra usar

## O que você tem em mãos

6 arquivos que juntos automatizam a montagem do JM:

- **`run_jm.py`** — o script único. É só ele que você (ou o Claude) chama toda semana.
- **`jm_engine.py`** — o motor (parser de fotos, diagramação automática, etc.)
- **`briefing_parser.py`** — lê o Word e entende a estrutura.
- **`template_compactado.pptx`** — o modelo visual (capa, matérias, cartazes, página final).
- **`tags_editorias/`** — as 11 tags de editoria + a logo de fechamento, prontas.

Esses arquivos não mudam mais toda semana — só o briefing e as fotos mudam.

## Instalação (uma vez só, com dois cliques)

Não precisa mexer no terminal. Escolha o do seu sistema:

- **Mac** → dois cliques em **`setup_mac.command`**
  - Se o Mac avisar "desenvolvedor não identificado": clique com o botão
    direito no arquivo → **Abrir** → **Abrir**.
- **Windows** → dois cliques em **`setup_windows.bat`**

O instalador confere e instala tudo sozinho: as bibliotecas do Python
(`python-pptx`, `python-docx`, `Pillow`, `pymupdf`) e o **LibreOffice**
(usado só pra montar a página final de miniaturas). No fim ele mostra um
✅ **TUDO PRONTO** — ou aponta exatamente o que faltou.

Sobre a **fonte Arial**: você quase nunca vai precisar instalar à mão. Mac
e Windows já vêm com Arial, e o sistema ainda traz a Liberation Sans (mesmas
medidas da Arial) na pasta `fonts/` como rede de segurança. Se um dia a Arial
faltar num computador, é só instalar a fonte por lá — mas na prática isso
raramente acontece.

Se o LibreOffice não for instalado, o JM ainda é gerado normalmente — só a
última página (as miniaturas) fica com a versão genérica do template, e
aparece um aviso dizendo como resolver.

## Passo 1 — Guardar os arquivos num lugar fixo (só uma vez)

Já conversamos sobre isso: como Projeto no Claude exige plano pago, a
alternativa gratuita é subir esses 6 itens num **repositório GitHub grátis**
(pode ser público, não tem nada sensível):

1. Crie uma conta grátis em github.com, se ainda não tiver.
2. Crie um repositório novo.
3. Arraste os 6 arquivos/pastas pra dentro (Add file → Upload files).
4. Guarde o link de cada arquivo (botão "Raw" em cada um).

## Passo 2 — Rodar toda semana

Abra uma conversa nova no claude.ai (plano Free já serve), anexe:
- o Word do briefing da semana
- as fotos da semana (baixadas do SharePoint), numa pasta ou .zip
- os QR codes da semana, se tiver (com um nome que deixe claro de qual matéria é)

E cole uma mensagem parecida com essa:

```
Baixa esses arquivos do meu repositório:
- Motor: <link raw do jm_engine.py>
- Parser: <link raw do briefing_parser.py>
- Script: <link raw do run_jm.py>
- Template: <link raw do template_compactado.pptx>
- Tags: <link raw da pasta tags_editorias, ou de cada arquivo dentro dela>

Depois roda isso (ajustando os nomes das matérias com QR code, se houver):

from run_jm import montar_jornal_mural
avisos = montar_jornal_mural(
    briefing_docx="<caminho do Word>",
    fotos_dir="<pasta com as fotos>",
    output_path="JM_pronto.pptx",
    qr_codes={"Título exato da matéria": "<caminho do QR>"},
)

Me devolve o .pptx e me mostra os avisos, se tiver algum.
```

## O que o `run_jm.py` já resolve sozinho

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
- Regenera a página final com as miniaturas certas dessa edição.

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
