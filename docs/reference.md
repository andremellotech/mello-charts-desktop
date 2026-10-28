# Referência

[Visão geral e downloads](../README.md) · [Instalar](install-guide.md) · [Usar o editor e o Office](native-integration.md) · [Cores](custom-colors.md) · [Exemplos dos 14 modelos](../examples/gallery/README.md) · [Verificação0.8.0](verification.md).

O projeto gera formas e textos editáveis no PowerPoint, com dados guardados nos metadados. Não cria objetosChart nativos nem depende de software comercial de gráficos. A versão0.8.0 inclui entrada pelo Excel, seleção XLSX por aba/intervalo, atualização explícita, Gantt, dispersão, bolhas e editor independente.

OsJSON v1 anteriores continuam compatíveis. Campos opcionais novos: `excelLink`, `fontFamily`, `seriesActual`, `seriesForecast`, `seriesBaseline`; linhas profissionais usam `startDate`, `endDate`, `progress`, `x`, `y`, `size`. Os nomes técnicos dos campos permanecem em inglês para compatibilidade, com interface em português.
