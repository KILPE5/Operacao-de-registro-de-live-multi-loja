# SoulBom - Sistema de Apresentadoras Multi Loja V9

## Correção principal: Admin leve

Esta versão foi feita para resolver o travamento causado por informações demais visíveis no Admin.

Mudanças:

- O overlay "Carregando..." foi removido da prática. Agora aparece só um indicador pequeno "Sincronizando...".
- A interface nunca fica bloqueada por carregamento.
- O Admin carrega apenas uma amostra recente:
  - 25 blocos recentes;
  - 80 vendas recentes.
- O histórico completo continua disponível por CSV.
- A visualização de "Blocos registrados" é vertical, em cards, um embaixo do outro.
- O Admin inteiro foi forçado em coluna única.
- A biblioteca de print html2canvas agora carrega só quando precisa encerrar/trocar live, deixando a abertura mais leve.
- O backend passou a buscar vendas de baixo para cima, evitando ler milhares de linhas no carregamento comum.

## Arquivos para o Apps Script

Use somente:

- Code.gs
- Index.html

Não crie appsscript.json, Styles.gs, Script.gs, Admin.gs ou arquivos extras.

## Instalação

1. Substitua Code.gs.
2. Substitua Index.html.
3. Salve.
4. Rode manualmente:

instalarSistemaApresentadorasMultiLoja

5. Autorize.
6. Vá em Implantar > Gerenciar implantações > Editar.
7. Escolha Nova versão.
8. Salve.

PIN padrão:

1234
