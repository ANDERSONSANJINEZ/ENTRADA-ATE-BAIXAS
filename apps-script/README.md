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

## Compartilhar com terceiros (modo somente leitura)

Repassar a URL do passo 10 dá acesso **completo** — quem tiver o link
importa, lança manual e remove, igual ao dono da planilha (o script sempre
roda "como você", não como quem está acessando). Para compartilhar uma
versão que esconde essas ações, acrescente `?modo=leitura` no fim da URL:

```
https://script.google.com/macros/s/SEU_ID/exec?modo=leitura
```

Isso esconde a aba **Importar** e o formulário de lançamento manual na
tela. **É só uma restrição de interface**, não uma trava de segurança real:
tecnicamente ainda é possível chamar as ações de escrita direto (via
requisição HTTP), já que o backend não diferencia quem está pedindo. Para
uma restrição de verdade (por login/conta Google), seria necessário mudar
"Quem pode acessar" na implantação — o que exige Google Workspace para
restringir a contas específicas.

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

## Atualizações do código

Sempre que `apps-script/Code.gs` for alterado neste repositório, repita os
passos 3–6 (nova implantação) para publicar a versão mais recente — a URL
de implantação muda a cada "Nova implantação"; se preferir manter a mesma
URL, use **Gerenciar implantações → editar (ícone de lápis) → Nova
versão** em vez de criar uma implantação nova.
