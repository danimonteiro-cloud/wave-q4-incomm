# Wave Q4 2026 · Acompanhamento Incomm

Site de acompanhamento das visitas e materiais da Wave Q4 2026 (Incomm), execução Bfly Consultoria.

Endereço: https://danimonteiro-cloud.github.io/wave-q4-incomm/

## Atualização

Os números ficam em `dados/pdvs.json` e `dados/historico.json`, atualizados todos os dias às 16h (Brasília).
A página (`index.html`) lê esses arquivos a cada acesso.

Para remontar o site a partir dos dados do painel:

```
python3 ferramentas/montar_site.py <pdvs.json> <historico.json> ferramentas/logo.png <pasta_saida>
```

Nenhum nome ou CPF de promotor é publicado.
