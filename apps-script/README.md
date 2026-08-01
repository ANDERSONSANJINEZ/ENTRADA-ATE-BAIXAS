# Backend (Google Sheets + Apps Script) — e a tela hospedada

Guarda os lançamentos importados do Protheus e os lançamentos manuais, e
calcula o status "Duplicado – revisar" quando um lançamento manual e um do
ERP têm o mesmo Código Fornecedor + Valor + Vencimento.

A mesma implantação do Apps Script **serve a tela do app** (`Index.html`)
— visitar a URL gerada no passo 9 abaixo, num navegador qualquer, já abre
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
5. Habilite o serviço avançado do Drive (necessário para o botão "Buscar
   arquivo mais recente do Drive" converter o .xlsx em planilha e lê-lo):
   no menu lateral esquerdo, clique no **+** ao lado de "Serviços",
   selecione **Drive API** e clique em Adicionar.
6. Confirme que sua conta Google tem acesso de visualização à pasta do
   Drive `17 CONTAS A PAGAR` (ela aparece em "Compartilhados comigo") —
   é a mesma conta que vai fazer a implantação como "Eu" no passo 8.
7. Salve o projeto (ícone de disquete).
8. Clique em **Implantar → Nova implantação**.
   - Tipo: **App da Web**.
   - Executar como: **Eu (seu e-mail)**.
   - Quem pode acessar: **Qualquer pessoa** (necessário para a tela web
     conseguir chamar o backend sem login adicional).
9. Na primeira execução, o Google vai pedir para autorizar o acesso ao
   Drive/Planilhas — autorize (é a sua própria conta agindo em seu nome).
10. Copie a **URL do app da Web** gerada (termina em `/exec`) — **essa é a
    URL para usar o app**. Salve como favorito no navegador/celular.

As abas `ERP` e `Manual` são criadas automaticamente na primeira chamada.

## Como usar

Abra a URL do passo 10 direto no navegador — computador, celular, tablet,
de qualquer lugar. Não precisa configurar nada nem copiar arquivo.

## Acesso leitura/edição (uma senha por pessoa)

Por padrão, qualquer pessoa que abrir a URL do app vê tudo (dashboard,
lançamentos, exportações) mas **não consegue editar** — a aba **Importar**,
o formulário de lançamento manual e o botão de excluir duplicados ficam
escondidos. Para habilitar a edição, a pessoa clica em **🔓 Habilitar
edição** (topo direito) e digita a sua senha; se for válida, o app já
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

## Importação automática da pasta do Drive

Na aba **Importar** da tela, o botão "🔄 Buscar arquivo mais recente do
Drive" chama a ação `importarDoDrive`: o backend procura, na pasta
`17 CONTAS A PAGAR` (ID `1sVlF29VGWDzHelgBGIeFvVK3OCpMjGmD`), o arquivo
`AAAA.MM.DD.xlsx` com a data mais recente no nome, converte para Google
Sheets temporariamente para ler os dados, e apaga a cópia temporária em
seguida. Não é preciso baixar/selecionar o arquivo manualmente — a opção
de importar por upload continua disponível como alternativa.

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

## Atualizações do código

Sempre que `apps-script/Code.gs` for alterado neste repositório, repita os
passos 3–6 (nova implantação) para publicar a versão mais recente — a URL
de implantação muda a cada "Nova implantação"; se preferir manter a mesma
URL, use **Gerenciar implantações → editar (ícone de lápis) → Nova
versão** em vez de criar uma implantação nova.
