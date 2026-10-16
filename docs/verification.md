# Verificação da versão 0.7.0

Conferido em 16 de outubro de 2026 no Windows, com compilador C# do .NET Framework. Comando executado: `./build.ps1 -RunTests -RenderGallery`. Resultado: **203 verificações passaram, nenhuma falhou**.

## Cobertura local

| Suíte | Verificações | Cobertura principal |
| --- | ---: | --- |
| Núcleo | 52 | Validação, números, JSON/CSV/TSV/XLSX, precisão, cálculos originais e texto |
| Ponte PowerPoint | 38 | COM, ribbon, metadados, elipses, polígonos e rollback com adaptadores |
| Novos gráficos | 45 | Oito tipos, cálculos, séries, limites, paletas, geometria e regressões |
| Editor e galeria | 27 | Tabelas, importação, precisão, seleção e renderização |
| Cores personalizadas | 20 | HEX/RGB, persistência, contraste, edição, importação, PNG e metadados |
| Integração | 21 | Slide de origem, seleção, cópia, ajuste, contexto, diagnóstico e rascunhos |

- DLL AnyCPU e instalador EXE compilados com DLL e ícones próprios embutidos; versões 0.7.0 consistentes.
- PNGs dos 11 modelos gerados pelo mesmo renderizador da prévia/exportação. Catálogo e exemplos de barras, linhas, rosca e Pareto inspecionados visualmente.
- Testes de 12 categorias, todos os estilos/paletas, nomes longos, valores negativos nos comparativos e números zero/subnormais.
- Legendas conferidas contra as cores desenhadas; marcadores de baseline evitam quadros dos valores nas barras.
- Eixos subnormais não repetem rótulos ou referência zero. Fatias imperceptíveis não geram polígonos degenerados; seus valores continuam no modelo/legenda.
- Área e rosca usam polígonos fechados; marcadores usam elipses nativas. Falhas de conversão/nomeação limpam os novos objetos e preservam o original.
- Precisão e projetos antigos preservados ao editar/importar. Pizza, rosca e Pareto usam tabela de duas colunas.
- Revisão independente do catálogo, editor e ponte. Galeria conferida em tamanho amplo e reduzido.
- Integração mantém o slide de origem, duplica o grupo nativo e tenta recuperar o último estado válido ao encerrar a janela original.

- Cores conferidas nos 11 tipos e na exportação PNG; janela de cores renderizada e inspecionada.
- Regressões para faixas sobre fundo médio e RGB com dígitos Unicode reproduzidas antes da correção e aprovadas após a correção.

## Validação no PowerPoint real ainda pendente

No ensaio anterior, `OfficeSmoke` tentou criar uma apresentação separada por `PowerPoint.Application`, mas recebeu `E_ACCESSDENIED` (80070005) antes do desenho. Não houve nova tentativa de contornar esse bloqueio.

**A ativação do suplemento instalado, a fidelidade dos polígonos no PowerPoint, a navegação com o editor aberto, o histórico de desfazer e o ciclo salvar/reabrir precisam do ensaio no computador de destino.** Os adaptadores testam o contrato da API; não comprovam uma instalação real do Office.

O instalador não foi executado para instalar/remover ou registrar o suplemento neste computador. A revisão de instalação/rollback foi estática. Os binários não têm assinatura digital comercial.

## Ensaio de aceitação no computador de destino

1. Salvar documentos, fechar o PowerPoint e instalar a versão 0.7.0.
2. Confirmar a aba Mello Charts e a galeria com 11 miniaturas; testar seleção e rolagem.
3. Inserir cada modelo e comparar o slide com o PNG. Conferir áreas, elipses, pizza/rosca, linhas e baseline.
4. Importar CSV de duas colunas na rosca e quatro colunas nas linhas; conferir valores e precisão.
5. Testar HEX/RGB e seletor em **Cores personalizadas…**, alterar fundo/fatias e reabrir o projeto; testar restaurar o padrão.
6. Editar, duplicar, mover, redimensionar, girar e atualizar o grupo; verificar desfazer.
7. Salvar/reabrir e recuperar a tabela pelo editor. Abrir um JSON da versão anterior.
8. Com duas apresentações, verificar o destino de origem e a recuperação ao fechar a janela original.
9. Conferir Diagnóstico/Verificar. Remover pelo instalador e confirmar que as apresentações continuam intactas.

Use apresentações de teste. Consulte o [guia de instalação](install-guide.md) se o Office não carregar o suplemento.
