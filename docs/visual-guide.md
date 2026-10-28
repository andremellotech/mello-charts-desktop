> Atualização0.8.0: o catálogo inclui Gantt, dispersão e bolhas. Veja os [14 modelos com exemplos](../examples/gallery/README.md) e o [guia atualizado](native-integration.md).

# Como criar um gráfico com boa apresentação

1. Na aba **Mello Charts** do PowerPoint, abra **Galeria de gráficos**, selecione uma miniatura e clique em **Usar modelo**.
2. Preencha a tabela ou use **Importar tabela / Colar do Excel** com os cabeçalhos. Trabalhe com 4 a 8 categorias quando possível. Até 12 são aceitas.
3. No painel lateral, escolha **Executivo**, **Essencial** ou **Escuro**. As paletas Atlântico, Floresta e Grafite mantêm as séries consistentes. O estilo Escuro adapta os tons ao fundo.
4. Escreva um título que diga o que a pessoa está comparando. Informe a unidade real dos dados, como `R$`, `R$ mil` ou `%`, e a fonte. A unidade é uma legenda: o programa não converte os valores.
5. Confira a prévia. Se os números forem longos, use **Abreviar números**. Se houver informação demais, oculte as linhas de referência ou os valores. A linha de zero continua visível.
6. Clique em **Inserir no slide**. Para rever os dados, selecione o grupo completo e use **Editar dados**. Use **Atualizar gráfico** e salve a apresentação.

## Qual gráfico escolher

| Tipo | Como funciona |
| --- | --- |
| Build-up | Mostra quanto foi realizado e quanto falta projetar em cada frente. |
| Cascata | Explica as variações do baseline ao EAC. |
| Comparativo | Compara realizado e projetado em cada período. |
| Barras horizontais | Duas barras por categoria e marcador opcional de baseline. Aceita negativos. |
| Linhas | Trajetórias com marcadores, baseline opcional e rótulos diretos nos últimos pontos. Aceita negativos. |
| Área comparativa | Áreas sobrepostas de realizado e projetado; não soma as séries. Valores não negativos. |
| Colunas + linha | Realizado em colunas e projetado em linha na mesma unidade e escala. Aceita negativos. |
| Empilhado 100% | Realizado + projetado restante normalizados por categoria. Parcelas não negativas. |
| Pizza / Rosca | Composição de uma série não negativa com total positivo. A rosca mostra o total no centro. |
| Pareto | Barras ordenadas por valor, curva acumulada e referência em 80%. Preserva a ordem da tabela no projeto. |

No build-up, **EAC = realizado + projetado restante**. O traço de baseline indica uma referência e não entra na soma. O consolidado opcional soma as categorias exibidas. Remova da tabela qualquer linha de total já calculada antes de usar **Adicionar total**, para não somar duas vezes.

Na cascata, os sinais indicam aumento e redução numérica. Eles não classificam automaticamente o resultado como bom ou ruim. A linha inicial é `base`, as variações são `delta` e os subtotais/totais são `total`, com valor calculado automaticamente.

Os gráficos são **grupos editáveis de formas, textos, elipses e polígonos**. Para alterar a tabela, use o editor Mello Charts. Eles não usam o objeto de gráfico do Office com planilha **Editar Dados**.

Pizza, rosca e Pareto exigem valores não negativos e total positivo. Área e empilhado 100% usam parcelas não negativas. Na área, as séries ficam sobrepostas sem serem somadas. No combinado, as duas séries usam a mesma unidade e escala. No empilhado, **realizado + projetado restante** é a base de 100% de cada categoria; uma linha toda zero aparece sem base.

Na pizza e na rosca, parcelas positivas abaixo de 0,1% aparecem como **<0,1%**. Fatias imperceptíveis são omitidas do desenho para evitar polígonos degenerados; seus valores permanecem na legenda, no total e no projeto.

O Pareto ordena somente o desenho e mostra o acumulado. A linha de 80% é uma referência fixa, não uma meta calculada.

Nas linhas, os nomes e valores finais são posicionados com linhas de ligação para reduzir sobreposições. Nas barras horizontais, o marcador de baseline evita os quadros dos valores.

## Como a abreviação funciona

As abreviações usam a unidade da tabela: `mil = 1.000`, `mi = 1.000.000` e `bi = 1.000.000.000`. Se você informou `R$ mil`, um rótulo `1,5 mil` representa 1.500 nessa mesma unidade, ou R$ 1.500.000. Para evitar essa combinação, prefira valores em `R$` com abreviação ou dados já escalados em `R$ mil` sem abreviação.

As casas decimais afetam a apresentação. Números abreviados usam pelo menos uma casa para reduzir ambiguidades. A tabela, o projeto JSON, os dados recuperados do grupo e o CSV preservam a precisão original. As marcas do eixo podem usar mais casas para representar corretamente seus intervalos.

## Exportar e reutilizar

- **Inserir no slide:** cria formas e textos nativos editáveis. Os dados ficam associados ao grupo. Mantenha o grupo para recuperar o gráfico pelo suplemento.
- **Exportar PNG:** salva uma imagem em 1920 × 1080 com a aparência da prévia. A imagem não contém formas editáveis nem os dados do gráfico.
- **Salvar projeto:** guarda dados e opções de aparência em JSON. Projetos anteriores à versão 0.4 recebem o estilo Executivo e continuam aceitos.
- **Exportar CSV:** guarda a tabela. Para preservar também a aparência, salve o JSON.

Os 11 modelos da pasta `examples/gallery` podem ser abertos pelo botão **Abrir projeto**. Use-os como ponto de partida e substitua os dados fictícios.

O tamanho inicial se ajusta ao slide, mantendo proporção 16:9 e margens. Use **Ajustar ao slide** para centralizar novamente e remover a rotação. Amplie o grupo proporcionalmente se precisar de rótulos maiores em uma apresentação. Ajustes manuais dentro do grupo são substituídos quando o suplemento redesenha o gráfico.

## Usar qualquer cor RGB

No painel do editor, clique em **Cores personalizadas…** e marque **Usar cores personalizadas**.

1. Digite um código **HEX**, como `#6D28D9` ou `#F0A`, ou **RGB**, como `rgb(109, 40, 217)`. O botão **Escolher…** abre o seletor do Windows com os canais vermelho, verde e azul.
2. Na aba **Séries e fundo**, ajuste realizado/base/total, projetado/linha, baseline/acumulado, aumentos, reduções e fundo.
3. Na aba **Pizza e rosca**, ajuste até 12 cores seguindo a ordem das categorias da tabela.
4. Clique em **Aplicar cores**, confira a prévia e insira ou atualize o gráfico.

São aceitas as **16.777.216 cores RGB opacas**, de `#000000` a `#FFFFFF`. Texto e números se adaptam para manter contraste; os preenchimentos e linhas preservam a cor escolhida. Uma cor muito parecida com o fundo pode tornar a própria série difícil de distinguir.

**Restaurar padrão** remove as substituições. Escolher outra **Paleta pronta** no editor também restaura as cores da paleta. Alterações de estilo mantêm as cores que você personalizou e adaptam as demais. Uma cor de fundo definida manualmente prevalece sobre o estilo.

O projeto JSON e os metadados do grupo preservam as cores; PNG e formas no PowerPoint usam o mesmo modelo de desenho. O CSV guarda somente a tabela. Use a versão 0.7 ou mais recente para preservar a personalização ao reabrir e salvar.

[Abrir exemplo JSON](../examples/gallery/cores-personalizadas.json) · [Ver exemplo](../examples/gallery/cores-personalizadas.png) · [Ver a janela de cores](../examples/gallery/editor-de-cores.png)
