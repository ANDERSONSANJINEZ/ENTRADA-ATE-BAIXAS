# Backend (Google Sheets + Apps Script)

Guarda os lançamentos importados do Protheus e os lançamentos manuais, e
calcula o status "Duplicado – revisar" quando um lançamento manual e um do
ERP têm o mesmo Código Fornecedor + Valor + Vencimento.

## Passo a passo (uma vez só)

1. Crie uma Google Sheet nova (em branco), com o nome que preferir — ex.:
   `Entrada até Baixas`.
2. Nela, abra **Extensões → Apps Script**.
3. Apague o conteúdo padrão de `Code.gs` e cole o conteúdo do arquivo
   `apps-script/Code.gs` deste repositório.
4. Habilite o serviço avançado do Drive (necessário para o botão "Buscar
   arquivo mais recente do Drive" converter o .xlsx em planilha e lê-lo):
   no menu lateral esquerdo, clique no **+** ao lado de "Serviços",
   selecione **Drive API** e clique em Adicionar.
5. Confirme que sua conta Google tem acesso de visualização à pasta do
   Drive `17 CONTAS A PAGAR` (ela aparece em "Compartilhados comigo") —
   é a mesma conta que vai fazer a implantação como "Eu" no passo 6.
6. Salve o projeto (ícone de disquete).
7. Clique em **Implantar → Nova implantação**.
   - Tipo: **App da Web**.
   - Executar como: **Eu (seu e-mail)**.
   - Quem pode acessar: **Qualquer pessoa** (necessário para a tela web
     conseguir chamar o backend sem login adicional).
8. Na primeira execução, o Google vai pedir para autorizar o acesso ao
   Drive/Planilhas — autorize (é a sua própria conta agindo em seu nome).
9. Copie a **URL do app da Web** gerada (termina em `/exec`).
10. Abra `web/index.html` no navegador, cole essa URL no campo de
    configuração exibido na primeira execução (fica salva no navegador).

As abas `ERP` e `Manual` são criadas automaticamente na primeira chamada.

## Importação automática da pasta do Drive

Na aba **Importar** da tela, o botão "🔄 Buscar arquivo mais recente do
Drive" chama a ação `importarDoDrive`: o backend procura, na pasta
`17 CONTAS A PAGAR` (ID `1sVlF29VGWDzHelgBGIeFvVK3OCpMjGmD`), o arquivo
`AAAA.MM.DD.xlsx` com a data mais recente no nome, converte para Google
Sheets temporariamente para ler os dados, e apaga a cópia temporária em
seguida. Não é preciso baixar/selecionar o arquivo manualmente — a opção
de importar por upload continua disponível como alternativa.

## Atualizações do código

Sempre que `apps-script/Code.gs` for alterado neste repositório, repita os
passos 3–6 (nova implantação) para publicar a versão mais recente — a URL
de implantação muda a cada "Nova implantação"; se preferir manter a mesma
URL, use **Gerenciar implantações → editar (ícone de lápis) → Nova
versão** em vez de criar uma implantação nova.
