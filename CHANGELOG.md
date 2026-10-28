# Histórico

## 0.8.0 experimental — 28/10/2026

- Gantt, dispersão e bolhas:14 modelos, datas inclusivas, progresso, coordenadas negativas e área proporcional.
- AbaExcel para criar gráficos da seleção; captura precisa, calendário 1904 e percentuais.
- XLSX por aba/intervalo A1, vínculo persistente e atualização explícita preservando aparência.
- Desfazer/refazer, retenção da última prévia válida, nomes de séries e tipografia.
- Editor independente portátil e instalador por usuário para Excel e PowerPoint.
- Normalização do grupo nativo preserva geometria; encerramentoCOM corrigido.
- Testes locais ampliados e ensaios reais Excel/PowerPoint com documentos separados.
- Os projetosJSON v1 anteriores permanecem compatíveis.


## 0.7.0 experimental

- Qualquer cor RGB opaca por seletor do Windows, HEX de 3/6 dígitos ou entrada RGB.
- Cores personalizadas de séries, baseline/acumulado, aumentos/reduções, fundo e 12 fatias de pizza/rosca.
- Texto com contraste automático nos preenchimentos e fundos personalizados, incluindo tons médios e faixas de destaque.
- Persistência em JSON, metadados, edição e importação. Restauração de cores padrão e compatibilidade com projetos anteriores.
- Validação de entrada inválida sem interromper o editor; exemplo visual e guia de cores incluídos.
- 203 verificações locais passaram, incluindo 20 de cores. Instalação e renderização real no Office continuam pendentes.

## 0.6.0 experimental

- Catálogo ampliado de 3 para 11 modelos: barras horizontais, linhas, área comparativa, colunas + linha, empilhado 100%, pizza, rosca e Pareto.
- Galeria nativa com miniaturas, navegação pelo teclado, duplo clique e rolagem em telas menores.
- Rótulos diretos nas linhas, legendas coordenadas, baseline nas barras, total central na rosca e Pareto com acumulado e referência em 80%.
- Elipses e polígonos fechados na prévia e na ponte PowerPoint, preservando edição como formas e metadados no grupo.
- Tabela categoria/valor para pizza, rosca e Pareto; importação e exportação preservam os novos tipos.
- Projetos anteriores compatíveis, mantendo o formato JSON de versão 1.
- Correções de legenda, sobreposição de baseline e valores, eixos subnormais e fatias imperceptíveis.
- Galeria com PNG, JSON e CSV dos 11 modelos e dados fictícios.
- 183 verificações locais: 52 núcleo, 38 ponte, 45 novos gráficos, 27 editor/galeria e 21 integração. Ativação instalada e desenho real no Office continuam pendentes.

## 0.5.0 experimental

- Editor vinculado ao PowerPoint, permitindo continuar navegando nos slides. Destino protegido por identidade da apresentação e ID do slide.
- Ribbon com oito comandos, ícones próprios e habilitação conforme a seleção.
- Edição por seleção de elementos de um único gráfico, duplicação nativa e ajuste ao tamanho real do slide.
- Entrada de desfazer nas operações de desenho, cópia e ajuste.
- Confirmação de alterações pendentes e recuperação do último estado válido ao encerrar a janela nativa de origem.
- Diagnóstico somente de leitura na ribbon e botão Verificar no instalador.
- 98 verificações locais: 52 do núcleo, 25 da ponte e 21 de integração. Office instalado ainda precisa de validação real.

## 0.4.0 experimental

- Estilos Executivo, Essencial e Escuro, com três paletas e contraste adaptado ao fundo.
- Painel lateral, prévia ampliada, tabela com leitura mais clara e exportação PNG em 1920 × 1080.
- Totais destacados, rótulos centrados nas barras e correção de sobreposição no comparativo e na cascata.
- Eixo com intervalos convencionais e precisão decimal para passos fracionários.
- Opções de abreviar números e ocultar linhas auxiliares, preservando os valores originais.
- Projetos e metadados anteriores continuam aceitos. Modelos JSON e galeria de imagens incluídos.
- 52 testes do núcleo e 22 verificações da ponte. Renderização real no Office ainda pendente.

## 0.3.0 experimental

- Instalador EXE por usuário, sem hospedagem nem dependência de Office.js.
- Integração COM com aba Mello Charts no PowerPoint desktop para Windows.
- Editor de build-up, cascata e comparativo, com importação CSV/TSV/XLSX e colagem de tabela do Excel.
- Formas e textos nativos, metadados no grupo e atualização preservando área e rotação.
- Projeto JSON, CSV, modelos fictícios, instalação e remoção com rollback.
- Testes de cálculo, importação, precisão, metadados e interfaces COM. Ativação real do suplemento instalado pendente de verificação.
