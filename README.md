# Entrada até Baixas

Ferramenta para acompanhar o ciclo de contas a pagar do TOTVS Protheus — da
entrada do título (emissão) até a baixa (pagamento) — com uma tela
interativa que importa o export do Protheus, permite lançamentos manuais
para itens que ainda não chegaram ao ERP e sinaliza quando um lançamento
manual e um título do ERP são o mesmo (duplicidade).

## Estrutura

- `apps-script/` — backend (Google Apps Script + Google Sheets) que guarda
  os títulos importados e os lançamentos manuais entre sessões, **e também
  hospeda a tela** (`Index.html`) num link único acessível de qualquer
  dispositivo. Veja `apps-script/README.md` para o passo a passo de
  implantação (uma vez só) — é o caminho recomendado.
- `web/index.html` — a mesma tela, como arquivo local (abrir direto no
  navegador, sem servidor), apontando para a mesma URL do Apps Script.
  Alternativa/backup ao link único.
  - `web/vendor/` — SheetJS, Chart.js e o plugin de rótulos incluídos
    localmente (sem depender de CDN), usados só por essa versão local.

Funcionalidades (iguais nas duas formas de uso):
- Importa o export "Listagem do Browse" do Protheus (.xlsx) e o transforma
  no layout definido para o app (colunas renomeadas/reordenadas, zeros à
  esquerda removidos de Nº Documento, Código Fornecedor e Remessa) — manual
  ou automaticamente, buscando o arquivo mais recente direto da pasta do Drive.
- Dashboard com indicadores, aging de títulos em aberto, fluxo de caixa
  projetado, situação das baixas ao longo do tempo, valor por remessa,
  resumo por aporte (ciclo semanal) e curva ABC de fornecedores.
- Aba de Lançamentos: lista tudo (ERP + Manual) com coluna **Origem**
  (ERP/Manual) e **Status**; quando um lançamento manual bate com um
  título do ERP (mesmo Código Fornecedor + Valor + Vencimento), ambos são
  marcados como "Duplicado – revisar" para que o manual seja removido.

## Uso do dia a dia

1. Abra a URL do app (gerada na implantação do Apps Script — veja
   `apps-script/README.md`) no navegador, de qualquer dispositivo.
2. Na aba **Importar**, clique em "🔄 Buscar arquivo mais recente do
   Drive" (busca sozinho na pasta `17 CONTAS A PAGAR`) ou selecione um
   `AAAA.MM.DD.xlsx` manualmente. O app lê a data do nome do arquivo e
   mostra no topo como "Base de dados"; fica em destaque (amarelo) quando
   o arquivo importado tem 2 dias ou mais.
3. Na aba **Lançamentos**, cadastre títulos que já existem na prática mas
   ainda não saíram no Protheus; quando o Protheus os incluir, o app
   sinaliza a duplicidade para você remover o lançamento manual.
4. Acompanhe o **Dashboard** para indicadores, aging, fluxo de caixa,
   situação das baixas, remessas e concentração de fornecedores.
