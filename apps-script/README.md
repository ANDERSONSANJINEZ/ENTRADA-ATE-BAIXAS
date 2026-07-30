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
4. Salve o projeto (ícone de disquete).
5. Clique em **Implantar → Nova implantação**.
   - Tipo: **App da Web**.
   - Executar como: **Eu (seu e-mail)**.
   - Quem pode acessar: **Qualquer pessoa** (necessário para a tela web
     conseguir chamar o backend sem login adicional).
6. Copie a **URL do app da Web** gerada (termina em `/exec`).
7. Abra `web/index.html` no navegador, cole essa URL no campo de
   configuração exibido na primeira execução (fica salva no navegador).

As abas `ERP` e `Manual` são criadas automaticamente na primeira chamada.

## Atualizações do código

Sempre que `apps-script/Code.gs` for alterado neste repositório, repita os
passos 3–6 (nova implantação) para publicar a versão mais recente — a URL
de implantação muda a cada "Nova implantação"; se preferir manter a mesma
URL, use **Gerenciar implantações → editar (ícone de lápis) → Nova
versão** em vez de criar uma implantação nova.
