# Backend (Google Sheets + Apps Script) — e a tela hospedada

Guarda os lançamentos importados do Protheus e os lançamentos manuais, e
calcula o status "Duplicado – revisar" quando um lançamento manual e um do
ERP têm o mesmo Código Fornecedor + Valor + Vencimento.

A mesma implantação do Apps Script **serve a tela do app** (`Index.html`)
— visitar a URL gerada no passo 8 abaixo, num navegador qualquer, já abre
o app funcionando. Tudo roda na nuvem: não há arquivo local para baixar ou
abrir.

## Passo a passo (uma vez só)

1. Crie uma Google Sheet nova (em branco), com o nome que preferir — ex.:
   `Entrada até Baixas`.
2. Nela, abra **Extensões → Apps Script**.
3. Apague o conteúdo padrão de `Code.gs` e cole o conteúdo do arquivo
   `apps-script/Code.gs` deste repositório.
4. No menu lateral esquerdo do editor, clique no **+** ao lado de
   "Arquivos", escolha **HTML**, nomeie exatamente **Index** (sem
   extensão — o Apps Script adiciona `.html` sozinho) e cole o conteúdo do
   arquivo `apps-script/Index.html` deste repositório.
5. Salve o projeto (ícone de disquete).
6. Clique em **Implantar → Nova implantação**.
   - Tipo: **App da Web**.
   - Executar como: **Eu (seu e-mail)**.
   - Quem pode acessar: **Qualquer pessoa** (necessário para a tela web
     conseguir chamar o backend sem login adicional).
7. Na primeira execução, o Google vai pedir para autorizar o acesso às
   Planilhas — autorize (é a sua própria conta agindo em seu nome).
8. Copie a **URL do app da Web** gerada (termina em `/exec`) — **essa é a
   URL para usar o app**. Salve como favorito no navegador/celular.

As abas `ERP` e `Manual` são criadas automaticamente na primeira chamada.

## Como usar

Abra a URL do passo 8 direto no navegador — computador, celular, tablet,
de qualquer lugar. Não precisa configurar nada nem copiar arquivo.

## Acesso leitura/edição (uma senha por pessoa)

Por padrão, qualquer pessoa que abrir a URL do app vê tudo (dashboard,
lançamentos, exportações) mas **não consegue editar** — o envio do arquivo
de títulos (baixas) na aba **Atualizar Dados**, o formulário de lançamento
manual e o botão de excluir duplicados ficam escondidos (o envio do
extrato bancário, na mesma aba, continua disponível mesmo sem edição). Para
habilitar a edição, a pessoa clica em **🔓 Habilitar edição** (topo
direito) e digita a sua senha; se for válida, o app já
mostra "bem-vindo(a), <nome>" e o navegador guarda a sessão (localStorage)
até clicar em **🔒 Voltar para leitura**.

A validação é conferida de verdade no servidor (não só uma restrição de
tela) — mesmo que alguém tente chamar uma ação de escrita direto por HTTP,
sem uma senha válida a ação é recusada.

**Para cadastrar quem pode editar** (quantas pessoas você quiser, cada uma
com a própria senha):
1. Abra a Google Sheet do app — vai ter surgido uma aba nova chamada
   **Usuários** (criada automaticamente assim que alguém abrir a tela pela
   primeira vez após esta atualização, ou você pode criá-la à mão se
   preferir).
2. Nela, cada linha é uma pessoa: coluna **Nome** e coluna **Senha**. Para
   adicionar alguém, é só preencher uma linha nova; para revogar o acesso
   de alguém, apague a linha dela; para trocar a senha de alguém, edite o
   valor da célula. Tudo direto na planilha, sem precisar mexer em código
   nem reimplantar o Apps Script.
3. Não existe limite de quantas pessoas cadastrar.

Enquanto a aba **Usuários** estiver vazia (nenhuma linha de dados), o app
libera a edição com **qualquer** senha digitada — assim a introdução dessa
trava não quebra o uso de quem ainda não cadastrou ninguém. Cadastre pelo
menos uma pessoa assim que possível para que a trava valha de verdade.

O parâmetro `?modo=leitura` na URL continua funcionando como reforço: força
a tela em modo leitura mesmo que o navegador já tenha uma sessão salva —
útil para um link específico que você quer garantir que nunca vai editar
nada, independentemente de quem o abra.

## Links de acesso restrito por pessoa (`?usuario=`)

Além do `?modo=leitura` (que só força leitura, sem mais restrição nenhuma),
dá pra criar um link próprio pra uma pessoa específica que só deve ver
**uma aba** e/ou **não deve ver alguns fornecedores**. Hoje existe um
perfil assim cadastrado: `?usuario=sandro-costa`, que abre a tela travada
só na aba **Lançamentos**, sem opção de editar, e sem os títulos de 14
fornecedores específicos (nem em Lançamentos, nem em nenhum outro lugar do
app — se um dia esse perfil ganhar acesso a outra aba, os fornecedores
continuam de fora de tudo).

Diferente do `?modo=leitura` (que só esconde botão na tela), esse filtro é
aplicado dentro do `api_carregar()` no servidor — os títulos dos
fornecedores excluídos nunca chegam a sair da planilha, então não tem como
"destravar" isso só editando algo no navegador.

Link pronto pra compartilhar com o Sandro: pegue a URL normal do Web App
(a mesma usada por todo mundo) e acrescente `?usuario=sandro-costa` no
final. Exemplo:
```
https://script.google.com/macros/s/SEU_ID_DE_IMPLANTACAO/exec?usuario=sandro-costa
```

**Para criar outro perfil restrito** (outra aba liberada e/ou outra lista
de fornecedores excluídos), edite o objeto `PERFIS_RESTRITOS` no início do
`Code.gs` e reimplante uma nova versão — é código, não dá pra cadastrar
pela planilha (diferente da aba **Usuários**, que é só pra senha de
edição). Cada chave do objeto vira o valor de `?usuario=` no link.

## Atualizar Dados (upload manual)

A aba **Atualizar Dados** reúne os dois uploads manuais do app, lado a
lado:

- **Títulos (baixas)** — envie o export do Protheus em `.xlsx`, com
  qualquer nome de arquivo (o nome não precisa mais seguir o padrão
  `AAAA.MM.DD.xlsx`; se o nome bater com esse padrão, a data base ainda é
  detectada automaticamente, senão ela fica marcada como "não
  identificada"). Não há mais busca automática numa pasta do Drive.
- **Extrato bancário** — envie o extrato do banco em `.xlsx`, também com
  qualquer nome. O resultado do cruzamento aparece na aba **Conciliação**
  (ver abaixo), não nesta aba.

O upload de títulos exige edição habilitada; o do extrato bancário
funciona mesmo em modo leitura (só salvar o resultado no histórico da
planilha exige edição — ver seção Conciliação Bancária).

## Anexos (documentos e comprovantes) por lançamento

Na aba **Lançamentos**, a coluna "Anexos" de cada linha tem dois botões (📄
documento, 🧾 comprovante). Diferente de versões anteriores, o link **não é
mais buscado automaticamente no Drive** (a busca automática não conseguia
achar os arquivos de forma confiável) — agora é colado manualmente:

- Clicar num botão sem link ainda cadastrado abre um campo pra colar o link
  (do Google Drive ou de onde o arquivo estiver salvo) e salva.
- Clicar num botão que já tem link cadastrado abre o mesmo campo,
  pré-preenchido — dá pra só confirmar pra abrir o link, ou editar e
  confirmar pra trocá-lo. O botão fica azul quando já tem um link salvo.

Os links ficam guardados nas colunas `Link Documento`/`Link Comprovante` da
planilha (abas `ERP` e `Manual`) e são **preservados entre reimportações**
do ERP — a reimportação diária substitui todos os títulos, mas reconhece o
mesmo título pela combinação Nº Documento + Código Fornecedor + Parcela e
recoloca o link que já tinha sido salvo para ele.

## Detalhamento por clique (dashboard, tabelas e Análise)

Em qualquer lugar que mostre um número agregado — cards do topo, barras/linhas
dos gráficos do Dashboard, linhas de "Resumo por Aporte" e "Curva ABC", e os
links dentro da aba **Análise** — clicar abre uma janela com a lista dos
títulos por trás daquele número (fornecedor, vencimento, valor, status),
já com botão próprio de "Exportar Excel". Serve pra ir direto do gráfico/KPI
para os documentos específicos que precisam de baixa ou revisão, sem precisar
filtrar manualmente em "Todos os lançamentos".

## Exportação em Excel

Todos os botões "⬇ Exportar Excel" geram um `.xlsx` com a mesma aparência da
tela (cabeçalho azul-claro em negrito, bordas, linha de cabeçalho fixa ao
rolar) — em "Detalhe por Aporte" a separação visual entre um vencimento e
outro também é reproduzida no arquivo. Essa formatação usa a biblioteca
ExcelJS (carregada via CDN, como o Chart.js); o SheetJS (`XLSX`) continua
sendo usado só para **ler** o `.xlsx` importado do Protheus.

## Histórico de alterações (auditoria)

Toda ação de escrita (importar, incluir/remover lançamento manual, excluir
duplicados, definir anexo) fica registrada com data/hora, usuário e detalhes
numa aba **Log**, criada automaticamente na planilha. A aba **Histórico** da
tela mostra as últimas 500 entradas (mais recente primeiro), com busca e
exportação em Excel — para o histórico completo, abra a aba "Log" direto na
planilha. A aba Log é podada automaticamente (mantém as ~2.000 mais
recentes) para não crescer sem limite.

## Alerta automático por e-mail

Um gatilho diário pode avisar **consorciovltce@gmail.com** quando:

- aparece um título vencido novo (que ainda não tinha entrado em nenhum
  alerta anterior) — não repete a lista de um dia pro outro, só o que é novo;
- a base de dados está desatualizada há 2 dias ou mais (mesmo critério do
  aviso "Base de dados" no topo da tela) — esse aviso se repete todo dia
  enquanto continuar desatualizada, como lembrete.

Sem nada de novo pra avisar, nenhum e-mail é enviado (não é um resumo diário
de rotina).

**Para ativar** (uma vez só), duas formas — escolha a que preferir:

- **Rodando uma função**: abra **Extensões → Apps Script** na planilha, no
  menu de funções no topo do editor selecione `configurarGatilhoDiario` e
  clique em **Executar**.
- **Direto pela interface, sem rodar código**: no editor do Apps Script,
  clique no ícone de relógio (⏰ **Gatilhos**) no menu lateral esquerdo → **+
  Adicionar gatilho** → em "Escolher qual função executar" selecione
  `enviarAlertaDiario` → em "Selecionar origem do evento" escolha **Baseado
  em tempo** → **Timer diário** → escolha um horário → **Salvar**.

Em ambos os casos, na primeira vez o Google vai pedir autorização para
enviar e-mail e gerenciar gatilhos (é a sua própria conta agindo em seu
nome). Isso instala um gatilho que roda todo dia (por padrão, 7h no fuso da
planilha, se ativado pela primeira forma). Para trocar o horário depois,
edite `.atHour(7)` na função `configurarGatilhoDiario` em `Code.gs` e rode a
função de novo — ela sempre remove o gatilho anterior antes de criar um
novo, então é seguro executar quantas vezes precisar (ou simplesmente edite
o horário direto na tela de Gatilhos, se tiver usado a segunda forma).

## Conciliação Bancária

Envie o extrato do banco em `.xlsx` na aba **Atualizar Dados** (painel
"Extrato bancário") — o resultado aparece na aba **Conciliação**, em duas
tabelas.

O app reconhece dois formatos de extrato pelo nome das colunas no
cabeçalho: uma única coluna **Valor** (débito já negativo) ou colunas
separadas **Crédito**/**Débito** (formato do Bradesco Net Empresa, por
exemplo — cada uma preenchida só quando é o caso da linha). A coluna de
histórico/descrição aceita variações como "Lançamento", "Data Movimento",
"Histórico" etc.

**Títulos pagos x Débitos do extrato.** Só entram na comparação os
títulos com **Data Baixa OU Vencimento dentro do período coberto pelo
extrato** (com uma folga de 5 dias) — títulos de outras épocas nunca
poderiam aparecer no extrato mesmo, então nem entram, evitando ruído em
"Sem correspondência". Também são excluídos títulos com Vencimento no
ano 2000 (marcador do Protheus pra imposto retido/recolhido
automaticamente, ex.: ISS retido na fonte — não é um pagamento avulso que
sai do banco como débito à parte). Três critérios de cruzamento, nessa
ordem — cada um só tenta pros títulos que o anterior não resolveu:

1. **Nº Documento + valor** — procura, no texto do lançamento, o número
   da nota citado (ex.: "COBRANCA NF 5858 ..." reconhece "5858"; aceita os
   mesmos prefixos do Tipo: NF, NFS, NFAG, ND, NF3E, DACTE, BOL, FAT, FOL,
   FGTS, DARF, DAM). Se bater com o Nº Documento do título **e** o valor
   for idêntico (tolerância de 1 centavo), concilia — sem limite de dias
   de diferença, já que é um sinal bem mais forte que só valor+data (o
   "Dcto." do próprio extrato é só uma referência interna do banco, não
   bate com o Nº Documento do título, por isso o número certo vem do
   texto do lançamento).
2. **Valor + Data Baixa mais próxima** — valor idêntico e Data Baixa
   igual ou até 5 dias de diferença da data do lançamento.
3. **Valor + Vencimento mais próximo** — mesma lógica do passo 2, mas
   usando o Vencimento em vez da Data Baixa (às vezes a Data Baixa
   lançada no sistema não bate com quando o banco processou de fato, mas
   o Vencimento fica mais próximo da data real do débito).

4 situações possíveis:

- **Conciliado** — bate valor e data.
- **Divergência de data** — bate o valor, mas a data no extrato é diferente
  da Data Baixa registrada.
- **Sem correspondência no extrato** — o título está marcado como baixado,
  mas nenhum débito do extrato bate com ele (pode ser baixa registrada
  errada, ou pago por outra conta).
- **Sem título correspondente** — tem um débito no extrato que não bate
  com nenhum título baixado (pode ser um pagamento não lançado no sistema).

**Créditos recebidos (aportes e rendimentos).** Toda linha de **crédito**
do extrato é agrupada pela origem, extraída do próprio texto do
Lançamento (ex.: "PIX RECEBIDO REM: CONSTRUTORA A GASPAR 01/07" vira
"CONSTRUTORA A GASPAR") — não depende de nenhuma lista cadastrada, então
reconhece qualquer consorciada nova automaticamente. Dois créditos recebem
categoria própria, fora da conta de aporte de consorciada:

- **Estado do Ceará** (🏛️) — pagamento da fatura que origina os aportes
  diários das consorciadas.
- **Rentabilidade / Aplicações financeiras** (📈) — rendimento de
  investimento (reconhece "RENTAB.", "Rendimento" etc. no texto).

Para as demais origens (aportes de consorciada), a tabela também mostra
"Débito Atribuído" e "Sobra Estimada": o total pago a fornecedores no
período é distribuído proporcionalmente entre as consorciadas, conforme o
peso do aporte de cada uma — é uma **estimativa** (a conta é única e
compartilhada; o app não sabe de fato qual aporte pagou qual fornecedor),
não um valor contábil exato.

Nada do extrato é salvo na planilha — o cruzamento acontece só na tela, a
cada vez que um arquivo é enviado.

## Atualizações do código

Sempre que `apps-script/Code.gs` for alterado neste repositório, cole o
conteúdo novo no editor e implante de novo (passos 3, 4 e 6 acima) para
publicar a versão mais recente — a URL
de implantação muda a cada "Nova implantação"; se preferir manter a mesma
URL, use **Gerenciar implantações → editar (ícone de lápis) → Nova
versão** em vez de criar uma implantação nova.
