# Guia consolidado Mello Charts

Os guias de instalação, operação, cores, validação e versões foram reunidos sem remover conteúdo. O README apresenta a entrada atual. Caminhos históricos continuam consultáveis no commit anterior; exemplos JSON/CSV/PNG ficam no pacote ZIP verificável.

- [Downloads e início rápido](../README.md)
- [Pacote completo de exemplos](../examples/MelloCharts-Examples.zip)
- Extraia com `./examples.ps1` na raiz; confira sem extrair com `./examples.ps1 -Verify`.


---

<a id="custom-colors"></a>

## docs/custom-colors.md

# Cores personalizadas

No painel do editor, clique em **Cores personalizadas…** e marque **Usar cores personalizadas**.

1. Digite um código **HEX**, como `#6D28D9` ou `#F0A`, ou **RGB**, como `rgb(109, 40, 217)`. O botão **Escolher…** abre o seletor do Windows com os canais vermelho, verde e azul.
2. Na aba **Séries e fundo**, ajuste realizado/base/total, projetado/linha, baseline/acumulado, aumentos, reduções e fundo.
3. Na aba **Pizza e rosca**, ajuste até 12 cores seguindo a ordem das categorias da tabela.
4. Clique em **Aplicar cores**, confira a prévia e insira ou atualize o gráfico.

São aceitas as **16.777.216 cores RGB opacas**, de `#000000` a `#FFFFFF`. Texto e números se adaptam para manter contraste; os preenchimentos e linhas preservam a cor escolhida. Uma cor muito parecida com o fundo pode tornar a própria série difícil de distinguir.

**Restaurar padrão** remove as substituições. Escolher outra **Paleta pronta** no editor também restaura as cores da paleta. Alterações de estilo mantêm as cores que você personalizou e adaptam as demais. Uma cor de fundo definida manualmente prevalece sobre o estilo.

O projeto JSON e os metadados do grupo preservam as cores; PNG e formas no PowerPoint usam o mesmo modelo de desenho. O CSV guarda somente a tabela. Use a versão 0.7 ou mais recente para preservar a personalização ao reabrir e salvar.

[Abrir exemplo JSON](https://github.com/andremellotech/mello-charts-desktop/blob/5975c40c24b338e2be658e698eafcf82982c4217/examples/gallery/cores-personalizadas.json) · [Ver exemplo](https://github.com/andremellotech/mello-charts-desktop/blob/5975c40c24b338e2be658e698eafcf82982c4217/examples/gallery/cores-personalizadas.png) · [Ver a janela de cores](https://github.com/andremellotech/mello-charts-desktop/blob/5975c40c24b338e2be658e698eafcf82982c4217/examples/gallery/editor-de-cores.png)


---

<a id="install-guide"></a>

## docs/install-guide.md

# Instalar, atualizar ou remover

1. Baixe o [instalador0.8.0](https://github.com/andremellotech/mello-charts-desktop/blob/5975c40c24b338e2be658e698eafcf82982c4217/release/MelloCharts-Setup.exe).
2. Salve seus documentos e feche Excel e PowerPoint.
3. Execute o instalador e use **Instalar**. A instalação é por usuário, em `%LOCALAPPDATA%\MelloCharts`.
4. Abra os aplicativos desktop e confira a aba **Mello Charts**.

Requisitos: Windows 10/11 Intel/AMD, .NET Framework 4.8, Office desktop 32/64 bits para a integração. O instalador registra classesCOM diferentes para Excel e PowerPoint nas views32/64, sem alterar políticas de segurança.

Para usar sem registrar abas, extraia todo o [pacote portátil](https://github.com/andremellotech/mello-charts-desktop/blob/5975c40c24b338e2be658e698eafcf82982c4217/release/MelloCharts-Portable.zip) e abra `MelloCharts.exe` ao lado da DLL. O editor importa/exporta gráficos; o envio ao slide exige PowerPoint instalado.

O instalador não possui assinatura digital comercial. Em ambiente gerenciado, siga o processo aprovado pela TI. **Verificar/Diagnóstico** apenas lê arquivo, versão, classesCOM e LoadBehavior; não comprova que o Office ativou a aba.

Para atualizar, feche ambos os aplicativos e execute o novo instalador. Para remover, use **Remover** no instalador ou Mello Charts Desktop em Configurações→Aplicativos. A remoção é limitada aos arquivos/registros próprios; não remove PPTX/XLSX/JSON/CSV do usuário.

Confira o SHA-256 do pacote com `Get-FileHash` e compare com [SHA256SUMS](https://github.com/andremellotech/mello-charts-desktop/blob/5975c40c24b338e2be658e698eafcf82982c4217/release/SHA256SUMS.txt) da mesma versão.

Instalação/ativação real e remoção ainda precisam de aceitação no computador de destino. A ponte COM foi validada diretamente com documentos novos no Excel/PowerPoint Microsoft 365. [Registro completo](https://github.com/andremellotech/mello-charts-desktop/blob/5975c40c24b338e2be658e698eafcf82982c4217/docs/verification.md).


---

<a id="native-integration"></a>

## docs/native-integration.md

# Usar o Mello Charts no Office

## PowerPoint

Selecione um slide em exibição Normal. Abra **Mello Charts → Galeria de gráficos** e escolha um dos 14 modelos. Edite a tabela ou use **Vincular Excel** para escolher um XLSX, aba e intervalo A1 com cabeçalho. Confira a prévia e clique em **Inserir no slide**.

O editor aberto pelo PowerPoint mantém o slide de origem como destino, mesmo durante a navegação. Para editar depois, selecione o grupo ou um elemento dele e use **Editar dados**. Atualizar preserva posição, largura, altura e rotação; recria os elementos internos. Duplicar mantém os ajustes nativos da cópia. Ajustar ao slide centraliza e redefine 16:9.

## Excel

Selecione uma tabela contínua com cabeçalho. Use **Mello Charts → Criar da seleção**, escolha o modelo e revise o editor. As fórmulas são lidas pelo resultado atual. Uma seleção de XLSX salvo guarda o vínculo; arquivos não salvos ou outros formatos usam uma cópia dos valores.

O primeiro **Enviar ao PowerPoint** cria uma apresentação nova, sem inserir em um documento pré-existente. Os seguintes envios do mesmo editor atualizam esse gráfico. O arquivo fica aberto para você salvar.

## Editor independente

Extraia todo o pacote portátil e abra `MelloCharts.exe`. Pode criar e exportar gráficos sem registrar suplemento. O envio ao slide exige PowerPoint instalado. Também aceita um caminhoJSON como argumento de abertura.

## Revisar e atualizar

**Desfazer/Refazer**, Ctrl+Z/Ctrl+Y, restaura dados e aparência. Ctrl+S salva projeto JSON. Durante entrada inválida, a prévia mantém o último estado válido e a aplicação ao slide fica desabilitada.

Nomes das séries e fonte tipográfica ficam no painel lateral para os modelos comparativos. Cores personalizadas aceitam HEX/RGB e seletor. As opções **Executivo**, **Essencial** e **Escuro** coordenam fundo, texto e linhas.

Para buscar mudanças do Excel: salve o XLSX, clique em **Atualizar dados**, confira e aplique ao slide. Se o arquivo, aba ou intervalo falhar, os dados anteriores continuam disponíveis. Não há atualização periódica automática; o intervalo é fixo e precisa incluir as linhas novas.

Use **Salvar projeto**, **Exportar PNG** e **Exportar CSV** para guardar saídas. PNG não retém edição vetorial; para formas editáveis use o botão de inserir/enviar ao PowerPoint. Não desagrupe ou apague os metadados se pretende recuperar a tabela.

Guias: [Instalação](https://github.com/andremellotech/mello-charts-desktop/blob/5975c40c24b338e2be658e698eafcf82982c4217/docs/install-guide.md) · [Cores](https://github.com/andremellotech/mello-charts-desktop/blob/5975c40c24b338e2be658e698eafcf82982c4217/docs/custom-colors.md) · [Modelos e exemplos](https://github.com/andremellotech/mello-charts-desktop/blob/5975c40c24b338e2be658e698eafcf82982c4217/examples/gallery/README.md) · [Validação](https://github.com/andremellotech/mello-charts-desktop/blob/5975c40c24b338e2be658e698eafcf82982c4217/docs/verification.md).


---

<a id="publish-guide"></a>

## docs/publish-guide.md

# Publicar uma atualização

1. Execute `./build.ps1 -RunTests -RenderGallery`.
2. Confira os pacotes `release/MelloCharts-Setup.exe`, `release/MelloCharts-Portable.zip` e `release/SHA256SUMS.txt` da mesma compilação.
3. Publique as fontes, exemplos e pacotes no repositório. A pasta `build/` deve permanecer fora do Git.
4. Confira o workflow Windows em Actions; ele recompila, testa e disponibiliza os pacotes como artefatos. Não executa Office nos runners hospedados.
5. Para uma Release, use a tag `v0.8.0`, as notas em [v0.8.0.md](https://github.com/andremellotech/mello-charts-desktop/blob/5975c40c24b338e2be658e698eafcf82982c4217/docs/releases/v0.8.0.md), os três arquivos da mesma compilação e marque a versão experimental como pré-lançamento. A CI não cria Releases automaticamente.

Não é necessário GitHub Pages. Os downloads atuais são ligados diretamente à pasta `release/` no README. [Validação](https://github.com/andremellotech/mello-charts-desktop/blob/5975c40c24b338e2be658e698eafcf82982c4217/docs/verification.md) registra os testes efetivamente executados e os limites.


---

<a id="reference"></a>

## docs/reference.md

# Referência

[Visão geral e downloads](https://github.com/andremellotech/mello-charts-desktop/blob/5975c40c24b338e2be658e698eafcf82982c4217/README.md) · [Instalar](https://github.com/andremellotech/mello-charts-desktop/blob/5975c40c24b338e2be658e698eafcf82982c4217/docs/install-guide.md) · [Usar o editor e o Office](https://github.com/andremellotech/mello-charts-desktop/blob/5975c40c24b338e2be658e698eafcf82982c4217/docs/native-integration.md) · [Cores](https://github.com/andremellotech/mello-charts-desktop/blob/5975c40c24b338e2be658e698eafcf82982c4217/docs/custom-colors.md) · [Exemplos dos 14 modelos](https://github.com/andremellotech/mello-charts-desktop/blob/5975c40c24b338e2be658e698eafcf82982c4217/examples/gallery/README.md) · [Verificação0.8.0](https://github.com/andremellotech/mello-charts-desktop/blob/5975c40c24b338e2be658e698eafcf82982c4217/docs/verification.md).

O projeto gera formas e textos editáveis no PowerPoint, com dados guardados nos metadados. Não cria objetosChart nativos nem depende de software comercial de gráficos. A versão0.8.0 inclui entrada pelo Excel, seleção XLSX por aba/intervalo, atualização explícita, Gantt, dispersão, bolhas e editor independente.

OsJSON v1 anteriores continuam compatíveis. Campos opcionais novos: `excelLink`, `fontFamily`, `seriesActual`, `seriesForecast`, `seriesBaseline`; linhas profissionais usam `startDate`, `endDate`, `progress`, `x`, `y`, `size`. Os nomes técnicos dos campos permanecem em inglês para compatibilidade, com interface em português.


---

<a id="v0.5.0"></a>

## docs/releases/v0.5.0.md

# Mello Charts Desktop 0.5.0

Suplemento COM gratuito para PowerPoint desktop no Windows, com instalador por usuário e funcionamento offline.

## Incluído

- Gráficos de build-up, cascata e comparativo, com formas e textos editáveis.
- Estilos Executivo, Essencial e Escuro, três paletas e exportação PNG em 1920 × 1080.
- Editor vinculado ao PowerPoint, mantendo a apresentação e o slide de origem como destino.
- Edição por elementos do gráfico, duplicação nativa e ajuste ao slide.
- Importação CSV/TSV/XLSX, colagem do Excel e backup JSON.
- Recuperação do último rascunho válido e diagnóstico somente de leitura.

## Baixar e instalar

Baixe **MelloCharts-Setup.exe**, feche o PowerPoint e abra o instalador. Clique em **Instalar**. Para atualizar uma versão anterior, repita esse procedimento.

Requer Windows 10/11 Intel/AMD, .NET Framework 4.8 e PowerPoint desktop de 32 ou 64 bits. Não atende PowerPoint na Web ou Mac.

## Estado da validação

**Versão experimental, sem assinatura digital comercial.** Passou por 98 verificações locais: 52 do núcleo, 25 da ponte e 21 de integração. A ativação instalada e o fluxo completo no PowerPoint real ainda precisam de validação. Consulte `docs/verification.md` no repositório.

O arquivo **SHA256SUMS.txt** contém o hash do instalador anexado. Alterações de assinatura ou compilação podem produzir outro hash, mesmo mantendo a versão.

Código próprio sob licença MIT. Não inclui software ou chaves de licença de terceiros.


---

<a id="v0.6.0"></a>

## docs/releases/v0.6.0.md

# Mello Charts Desktop 0.6.0

O catálogo passa de 3 para 11 modelos de gráficos, com galeria visual no PowerPoint e formas editáveis no slide.

- Novos: barras horizontais, linhas, área comparativa, colunas + linha, empilhado 100%, pizza, rosca e Pareto.
- Rótulos diretos, cores coordenadas, total central na rosca, acumulado do Pareto e baseline opcional nos comparativos.
- Miniaturas, navegação pelo teclado e modelos fictícios prontos em PNG/JSON/CSV.
- Tabela categoria/valor para composição e Pareto; projetos anteriores continuam aceitos.

## Instalar ou atualizar

Baixe `MelloCharts-Setup.exe` e `SHA256SUMS.txt` da mesma compilação. Feche o PowerPoint, execute o instalador e clique em Instalar. Abra **Mello Charts → Galeria de gráficos → Usar modelo**.

Requer Windows 10/11 Intel/AMD, .NET Framework 4.8 e PowerPoint desktop 32/64 bits. Não exige hospedagem nem conta Mello Charts.

## Validação

**Versão experimental, sem assinatura digital comercial.** 183 verificações locais passaram: 52 núcleo, 38 ponte, 45 novos gráficos, 27 editor/galeria e 21 integração. Ativação instalada e renderização no Office real ainda precisam de ensaio, conforme `docs/verification.md`.

São grupos de formas e textos. Use o editor Mello Charts para a tabela; não são o gráfico nativo com planilha Editar Dados. Não inclui Gantt, Mekko ou vínculo automático com Excel.

Código próprio sob licença MIT; não distribui componentes ou licenças de terceiros.


---

<a id="v0.7.0"></a>

## docs/releases/v0.7.0.md

# Mello Charts Desktop 0.7.0

Os 11 modelos agora aceitam qualquer cor RGB opaca, escolhida por seletor, HEX ou valores RGB.

- Cores de realizado/base/total, projetado/linha, baseline/acumulado, aumentos, reduções e fundo.
- Até 12 cores individuais para pizza e rosca.
- Contraste automático dos textos e números; valores originais preservados.
- JSON, metadados no grupo e importação mantêm a personalização.
- Restaurar padrão e paletas prontas permitem voltar às cores originais.

Feche o PowerPoint, execute MelloCharts-Setup.exe e clique em Instalar. No editor, abra **Cores personalizadas…**.

**Experimental:** 203 verificações locais passaram (52 núcleo, 38 ponte, 45 novos gráficos, 27 editor/galeria, 20 cores e 21 integração). Instalação e renderização no PowerPoint real continuam pendentes. O instalador não tem assinatura digital comercial.

Requer Windows 10/11 Intel/AMD, .NET Framework 4.8 e PowerPoint desktop 32/64 bits. Os gráficos são grupos de formas editáveis; a tabela é editada pelo Mello Charts. Código próprio sob licença MIT.


---

<a id="v0.8.0"></a>

## docs/releases/v0.8.0.md

# Mello Charts 0.8.0 experimental

Criar gráficos a partir do Excel, revisar com prévia e inserir como formas editáveis no PowerPoint. Três novos modelos: Gantt, dispersão e bolhas. ImportaçãoXLSX por aba/intervalo, atualização local explícita, desfazer/refazer, nomes de séries e fontes. Editor portátil e instalador das abas para ambos os aplicativos.

[Instalador](https://github.com/andremellotech/mello-charts-desktop/blob/5975c40c24b338e2be658e698eafcf82982c4217/release/MelloCharts-Setup.exe) · [Portátil](https://github.com/andremellotech/mello-charts-desktop/blob/5975c40c24b338e2be658e698eafcf82982c4217/release/MelloCharts-Portable.zip) · [SHA-256](https://github.com/andremellotech/mello-charts-desktop/blob/5975c40c24b338e2be658e698eafcf82982c4217/release/SHA256SUMS.txt) · [Validação](https://github.com/andremellotech/mello-charts-desktop/blob/5975c40c24b338e2be658e698eafcf82982c4217/docs/verification.md)

Windows e .NET Framework 4.8. IntegraçãoOffice desktop; execuçãoCOM validada com Microsoft 365 16.0.20527.20088. Ativação instalada e políticas corporativas dependem do computador de destino. Os pacotes não possuem assinatura digital comercial.


---

<a id="2026-10-28-chart-workbench"></a>

## docs/superpowers/plans/2026-10-28-chart-workbench.md

# Mello Charts 0.8 — criação de gráficos para Office

Goal: acelerar a transformação de tabelas do Excel em gráficos editáveis no PowerPoint, com um editor mais claro e novos modelos úteis.

Architecture: preservar o núcleo de cenas vetoriais e os grupos nativos de formas do PowerPoint. Adicionar importação XLSX por planilha e intervalo, vínculo local com atualização explícita, histórico no editor e gráficos Gantt, dispersão e bolhas. Manter projetos JSON v1 compatíveis por campos opcionais.

Tech stack: C#/.NET Framework 4.8, WinForms, COM Office, Open XML via ZIP/XML sem dependências de terceiros.

Spec: pedido do usuário de aprimorar o aplicativo para criar gráficos no Excel/PowerPoint e publicá-lo no GitHub. A implementação é uma evolução verificável, sem alegar equivalência a todas as funções de produtos comerciais.

Global constraints: não alterar arquivos do usuário; não executar instalador recebido; testar somente documentos novos; não mudar políticas de segurança do Office; preservar limites e validação do importador; interface em português. Arquivos vinculados só são consultados por ação do usuário. Dados locais não são enviados para serviços externos.

Review focus: compatibilidade de metadados; Gantt por datas e bolhas por área; erro de atualização preserva gráfico/dados anteriores; segurança ZIP/XML; cleanup COM idempotente; instalação por usuário e documentação honesta.

## Trabalho

1. Preservar baseline e reproduzir testes (concluído: 203 testes; smoke criou três slides e revelou falha no cleanup COM).
2. Testes antes do código para novos modelos: agente de renderização trabalha em ChartCatalog.cs, ProfessionalCharts.cs e ProfessionalChartTests.cs. O coordenador integra campos opcionais e rotas no núcleo.
3. Testes antes do código para importação por seleção: agente de dados trabalha em WorkbookData.cs, WorkbookImportForm.cs e WorkbookDataTests.cs. Núcleo expõe ImportTable e propriedade excelLink.
4. Coordenador integra editor: histórico desfazer/refazer, tabela por modelo, intervalo do Excel e atualização explícita; atualiza suporte e instalação quando necessário.
5. Corrigir cleanup COM com evidência; compilar e executar suíte completa, renderizar galeria e verificar no PowerPoint real em apresentação nova. Validar importação com arquivo Excel novo.
6. Revisar fontes, documentação, workflow e instalador; publicar código e artefatos no repositório do usuário; verificar CI e relatar limitações.

## Contratos

- ChartRow: campos opcionais startDate/endDate (ISO yyyy-MM-dd), progress (0–100), x/y/size (double).
- ChartCore: IsProfessionalType, ProfessionalExample, ValidateProfessionalRow, ReadProfessionalRow, ImportProfessionalRows, ProfessionalCsv, BuildProfessionalScene implementados em partial separado. Modelos novos limitados a 24 tarefas ou 80 pontos; antigos mantêm 12 categorias nesta versão.
- WorkbookSource: path, sheet, range; ChartDocument.excelLink opcional.
- WorkbookData.ListSheets(path), Import(path,sheet,range,type), Refresh(document). ImportTable(List<List<string>>, type, invariantNumbers) é a entrada validada do núcleo. A1 inclui cabeçalho. Refresh preserva aparência e substitui apenas linhas/vínculo.
- WorkbookImportForm(type) e Result. Seleção/resultado só substituem documento após validação completa.

## Verificação

Testes exercitam dados, roundtrip, limites, erros, geometria e preservação. Galeria visual inclui os 14 tipos. Smoke Office cria, atualiza, lê metadados e salva somente apresentações novas. A instalação/ribbon será descrita segundo o que for de fato verificado.


---

<a id="verification"></a>

## docs/verification.md

# Verificação 0.8.0 — 28/10/2026

## Build e testes

Comando: `./build.ps1 -RunTests -RenderGallery`, Windows, compilador do .NET Framework4.x. A versão0.7.0 recebida tinha203 verificações aprovadas. A evolução adiciona testes de modelos, importação, editor e Excel.

| Suíte | Verificações |
| --- | ---: |
| CoreTests |52|
| BridgeTests |39|
| ExtendedChartTests |45|
| EditorCatalogTests |33|
| ColorTests |20|
| ProfessionalChartTests |17|
| WorkbookDataTests |25|
| WorkbenchTests |6|
| ExcelIntegrationTests |15|
| IntegrationTests |21|
| **Total** |**273**|

Os testes verificam roundtrip JSON/CSV/metadados, compatibilidade legada, precisão, cores e contraste, geometria, datas, áreas das bolhas, limites, rótulos densos, seleção Excel, atualização atômica, segurançaZIP/XML, desfazer/refazer e retenção da última prévia válida. Galeria, editor Gantt e diálogo Excel renderizados e inspecionados.

## PowerPoint real

`./test-office.ps1`, Microsoft 365desktop 16.0.20527.20088: os 14 tipos foram criados, atualizados, lidos dos metadados e salvos em uma apresentação nova. Conferidas posição, largura, altura e rotação. O agrupamento nativo aumentava a altura em cerca de 3,87 pt; o teste reproduziu a falha e aprovou após normalização do grupo. O cleanup COM anterior falhava após fechar a apresentação; os argumentos agora usam objetos estáticos e o ensaio encerrou com código 0.

O primeiro teste no sandbox recebeu E_ACCESSDENIED na ativaçãoCOM. Os ensaios finais foram executados fora desse sandbox, com a revisão de permissão do ambiente; nenhuma política do Office foi modificada.

## Excel real

`./test-excel.ps1`, mesma versão Office: arquivo novo e segunda aba, seleçãoH500:K502, fórmula calculada 3.125 e atualização salva 12.375. O vínculo recarregou os dados e preservou título, nomes das séries e casas decimais. Gantt com calendário 1904 e progresso 50% foi comparado entre captura da seleção e importação XLSX; célula 0.5 sem formatação permaneceu 0.5 ponto. Código de saída 0. O ensaio fecha apenas o workbook criado e não encerra uma instância pré-existente do Excel.

## Limites do que foi comprovado

- O instalador e o editor independente foram compilados das fontes. Não foi executada uma instalação/remoção real neste computador, nem comprovada a ativação das abas a partir do registro.
- Testes COM diretos não equivalem a teste de todas as versões do Office, arquiteturas ou políticas corporativas.
- Não há assinatura digital comercial. A integração é exclusiva do Office desktop no Windows.
- Vínculos XLSX são locais e atualizados por ação explícita, com fórmulas em cache. Não há sincronização em tempo real.
- A atualização recria o grupo: ajustes manuais dentro dele são substituídos. O original é removido somente após confirmar os dados no novo grupo.

## Aceitação no destino

Salvar arquivos, fechar Excel/PowerPoint, instalar, conferir ambas as abas. Testar uma seleção com cabeçalhos, enviar à apresentação nova, salvar/reabrir, editar e atualizar. Trocar de apresentação com editor vinculado deve bloquear inserção no arquivo errado. Verificar remoção e integridade dos documentos. Para ambientes gerenciados, usar o procedimento da TI.


---

<a id="visual-guide"></a>

## docs/visual-guide.md

> Atualização0.8.0: o catálogo inclui Gantt, dispersão e bolhas. Veja os [14 modelos com exemplos](https://github.com/andremellotech/mello-charts-desktop/blob/5975c40c24b338e2be658e698eafcf82982c4217/examples/gallery/README.md) e o [guia atualizado](https://github.com/andremellotech/mello-charts-desktop/blob/5975c40c24b338e2be658e698eafcf82982c4217/docs/native-integration.md).

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

[Abrir exemplo JSON](https://github.com/andremellotech/mello-charts-desktop/blob/5975c40c24b338e2be658e698eafcf82982c4217/examples/gallery/cores-personalizadas.json) · [Ver exemplo](https://github.com/andremellotech/mello-charts-desktop/blob/5975c40c24b338e2be658e698eafcf82982c4217/examples/gallery/cores-personalizadas.png) · [Ver a janela de cores](https://github.com/andremellotech/mello-charts-desktop/blob/5975c40c24b338e2be658e698eafcf82982c4217/examples/gallery/editor-de-cores.png)
