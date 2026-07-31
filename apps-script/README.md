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

## Acesso leitura/edição (senha de edição)

Por padrão, qualquer pessoa que abrir a URL do app vê tudo (dashboard,
lançamentos, exportações) mas **não consegue editar** — a aba **Importar**,
o formulário de lançamento manual e o botão de excluir duplicados ficam
escondidos. Para habilitar a edição, a pessoa clica em **🔓 Habilitar
edição** (topo direito) e digita a senha de edição da equipe; o navegador
guarda essa senha (localStorage) para não pedir de novo nas próximas
visitas, até clicar em **🔒 Voltar para leitura**.

Diferente do antigo `?modo=leitura` (que só escondia botões na tela), essa
senha é conferida de verdade no servidor a cada ação de escrita — então
mesmo que alguém tente chamar a ação direto por HTTP, sem a senha certa a
ação é recusada.

**Para configurar a senha** (recomendado fazer isso assim que implantar o
app):
1. No editor do Apps Script, clique no ícone de engrenagem **⚙️
   Configurações do projeto** no menu lateral esquerdo.
2. Em **Propriedades do script**, clique em **Adicionar propriedade do
   script**.
3. Propriedade: `senhaEdicao` — Valor: a senha que a equipe vai usar.
   Salve.

Enquanto essa propriedade não for configurada, o app libera a edição com
**qualquer** senha digitada (ou seja, sem trava real) — assim a introdução
dessa funcionalidade não quebra o uso de quem ainda não configurou nada.
Configure a propriedade assim que possível para que a trava valha de
verdade.

O parâmetro `?modo=leitura` na URL continua funcionando como antes, como
reforço: força a tela em modo leitura mesmo que o navegador já tenha uma
senha de edição salva — útil para um link específico que você quer
garantir que nunca vai editar nada, independentemente de quem o abra.

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
documento, 🧾 comprovante) que buscam no Drive, pelo Nº Documento (e
refinando pelo nome do fornecedor quando há mais de um resultado), o
arquivo correspondente e abrem em nova aba. As pastas de busca estão fixas
em `Code.gs` (`PASTAS_DOCUMENTOS`, `PASTAS_COMPROVANTES`) e sua conta Google
precisa ter acesso de visualização a elas, do mesmo jeito que à pasta do
export diário. Se algum dia essas pastas mudarem, atualize os IDs ali.

Se o botão disser "Nenhum documento/comprovante encontrado" mas você sabe
que o arquivo existe na pasta, o clique agora mostra o erro real do Drive
(em vez de simplesmente não achar nada) — normalmente é uma destas causas:
- O serviço avançado **Drive API** foi adicionado como v2 em vez de v3 (o
  código já tenta os dois formatos automaticamente, mas confirme no passo
  5 acima que o serviço está habilitado).
- As pastas estão dentro de um **Drive compartilhado** (Shared Drive) da
  organização e a conta que fez a implantação não tem acesso a elas.
- O nome do arquivo no Drive não contém o Nº Documento exatamente como
  está na planilha (ex.: zeros à esquerda a mais/a menos).

## Atualizações do código

Sempre que `apps-script/Code.gs` for alterado neste repositório, repita os
passos 3–6 (nova implantação) para publicar a versão mais recente — a URL
de implantação muda a cada "Nova implantação"; se preferir manter a mesma
URL, use **Gerenciar implantações → editar (ícone de lápis) → Nova
versão** em vez de criar uma implantação nova.
