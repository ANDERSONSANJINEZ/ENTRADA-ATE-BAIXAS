# Entrada até Baixas

Ferramenta para acompanhar o ciclo de contas a pagar do TOTVS Protheus — da
entrada do título (emissão) até a baixa (pagamento) — com uma tela
interativa que importa o export do Protheus, permite lançamentos manuais
para itens que ainda não chegaram ao ERP e sinaliza quando um lançamento
manual e um título do ERP são o mesmo (duplicidade).

## Estrutura

- `web/index.html` — a tela (abrir direto no navegador, não precisa de servidor/instalação).
  - Importa o export "Listagem do Browse" do Protheus (.xlsx) e o transforma
    no layout definido para o app (colunas renomeadas/reordenadas, zeros à
    esquerda removidos de Nº Documento, Código Fornecedor e Remessa).
  - Dashboard com indicadores (total, em aberto, vencido, aguardando
    aprovação), aging de títulos em aberto, fluxo de caixa projetado
    (mensal) e curva ABC de fornecedores.
  - Aba de Lançamentos: lista tudo (ERP + Manual) com coluna **Origem**
    (ERP/Manual) e **Status**; quando um lançamento manual bate com um
    título do ERP (mesmo Código Fornecedor + Valor + Vencimento), ambos são
    marcados como "Duplicado – revisar" para que o manual seja removido.
  - `web/vendor/` — SheetJS e Chart.js incluídos localmente (sem depender de CDN).
- `apps-script/` — backend (Google Apps Script + Google Sheets) que guarda
  os títulos importados e os lançamentos manuais entre sessões. Veja
  `apps-script/README.md` para o passo a passo de implantação (uma vez só).

## Uso do dia a dia

1. Abra `web/index.html` no navegador (primeira vez: clique em "⚙
   Configurar planilha" e cole a URL do Web App do Apps Script).
2. Na aba **Importar**, selecione o export mais recente da pasta do Drive
   (`4 - CE 007_ADMINISTRATIVO / 4.11 - FINANCEIRO / 17 CONTAS A PAGAR`).
   O arquivo deve manter o nome padrão `AAAA.MM.DD.xlsx` (ex.:
   `2026.07.30.xlsx`) — o app lê essa data do nome do arquivo e mostra no
   topo da tela como "Base de dados"; o indicador fica em destaque
   (amarelo) quando o arquivo importado tem 2 dias ou mais.
3. Na aba **Lançamentos**, cadastre títulos que já existem na prática mas
   ainda não saíram no Protheus; quando o Protheus os incluir, o app
   sinaliza a duplicidade para você remover o lançamento manual.
4. Acompanhe o **Dashboard** para indicadores, aging, fluxo de caixa e
   concentração de fornecedores.
