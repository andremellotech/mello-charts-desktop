# Como criar um gráfico com boa apresentação

1. Na aba **Mello Charts** do PowerPoint, escolha **Build-up**, **Cascata** ou **Comparativo**.
2. Preencha a tabela ou use **Importar tabela / Colar do Excel** com os cabeçalhos. Trabalhe com 4 a 8 categorias quando possível. Até 12 são aceitas.
3. No painel lateral, escolha **Executivo**, **Essencial** ou **Escuro**. As paletas Atlântico, Floresta e Grafite mantêm as séries consistentes. O estilo Escuro adapta os tons ao fundo.
4. Escreva um título que diga o que a pessoa está comparando. Informe a unidade real dos dados, como `R$`, `R$ mil` ou `%`, e a fonte. A unidade é uma legenda: o programa não converte os valores.
5. Confira a prévia. Se os números forem longos, use **Abreviar números**. Se houver informação demais, oculte as linhas de referência ou os valores. A linha de zero continua visível.
6. Clique em **Inserir no slide**. Para rever os dados, selecione o grupo completo e use **Editar dados**. Use **Atualizar gráfico** e salve a apresentação.

## Qual gráfico escolher

| Pergunta | Tipo |
| --- | --- |
| Quanto já foi realizado e quanto falta projetar em cada frente? | Build-up |
| Quais variações levam do baseline ao EAC? | Cascata |
| Como realizado e projetado se comparam em cada período? | Comparativo |

No build-up, **EAC = realizado + projetado restante**. O traço de baseline indica uma referência e não entra na soma. O consolidado opcional soma as categorias exibidas. Remova da tabela qualquer linha de total já calculada antes de usar **Adicionar total**, para não somar duas vezes.

Na cascata, os sinais indicam aumento e redução numérica. Eles não classificam automaticamente o resultado como bom ou ruim. A linha inicial é `base`, as variações são `delta` e os subtotais/totais são `total`, com valor calculado automaticamente.

## Como a abreviação funciona

As abreviações usam a unidade da tabela: `mil = 1.000`, `mi = 1.000.000` e `bi = 1.000.000.000`. Se você informou `R$ mil`, um rótulo `1,5 mil` representa 1.500 nessa mesma unidade, ou R$ 1.500.000. Para evitar essa combinação, prefira valores em `R$` com abreviação ou dados já escalados em `R$ mil` sem abreviação.

As casas decimais afetam a apresentação. Números abreviados usam pelo menos uma casa para reduzir ambiguidades. A tabela, o projeto JSON, os dados recuperados do grupo e o CSV preservam a precisão original. As marcas do eixo podem usar mais casas para representar corretamente seus intervalos.

## Exportar e reutilizar

- **Inserir no slide:** cria formas e textos nativos editáveis. Os dados ficam associados ao grupo. Mantenha o grupo para recuperar o gráfico pelo suplemento.
- **Exportar PNG:** salva uma imagem em 1920 × 1080 com a aparência da prévia. A imagem não contém formas editáveis nem os dados do gráfico.
- **Salvar projeto:** guarda dados e opções de aparência em JSON. Projetos anteriores à versão 0.4 recebem o estilo Executivo e continuam aceitos.
- **Exportar CSV:** guarda a tabela. Para preservar também a aparência, salve o JSON.

Os três modelos da pasta `examples/gallery` podem ser abertos pelo botão **Abrir projeto**. Use-os como ponto de partida e substitua os dados fictícios.

O tamanho inicial se ajusta ao slide, mantendo proporção 16:9 e margens. Use **Ajustar ao slide** para centralizar novamente e remover a rotação. Amplie o grupo proporcionalmente se precisar de rótulos maiores em uma apresentação. Ajustes manuais dentro do grupo são substituídos quando o suplemento redesenha o gráfico.
