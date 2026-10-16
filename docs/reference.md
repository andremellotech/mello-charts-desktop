# Mello Charts Desktop

Suplemento gratuito e de código aberto para **PowerPoint no Windows**, com instalador EXE e funcionamento offline. Crie gráficos financeiros com dados em tabela, prévia e formas editáveis no slide.

**Versão 0.6.0 experimental.** O instalador foi compilado; o núcleo e a ponte COM passaram pelos testes descritos em [Verificação](verification.md). A ativação do suplemento instalado em um PowerPoint real ainda precisa ser conferida. O ambiente de desenvolvimento negou a inicialização automática do PowerPoint com `E_ACCESSDENIED`.

## Baixar e instalar

1. Baixe **[MelloCharts-Setup.exe](../release/MelloCharts-Setup.exe)**. Não é necessário clonar o código para instalar.
2. Salve seus documentos e **feche o PowerPoint**.
3. Abra o instalador e clique em **Instalar**. Ele instala somente para seu usuário, sem administrador.
4. Abra o PowerPoint desktop. Procure a aba **Mello Charts**.
5. Selecione um slide e abra **Galeria de gráficos** e escolha um dos 11 modelos.

Requer Windows 10/11 Intel/AMD, .NET Framework 4.8 e PowerPoint desktop instalado (32 ou 64 bits). Não funciona no PowerPoint na Web, Mac, celular ou iPad. Não depende de hospedagem, conta Mello Charts, Office.js ou chave de licença. O uso do próprio PowerPoint continua sujeito à sua conta/licença Microsoft.

O instalador ainda não tem assinatura digital comercial. Políticas corporativas podem impedir seu uso; nesse caso, solicite a instalação aprovada à TI. Ele não altera políticas de macros, segurança, certificados ou proteção do Office.

Veja o [guia completo de instalação](install-guide.md).

## Criar e atualizar

1. Selecione um slide em exibição **Normal**.
2. Na aba **Mello Charts**, abra **Galeria de gráficos** e confirme **Usar modelo**.
3. Edite a tabela, use **Colar do Excel** com cabeçalho ou **Importar tabela** para CSV/TSV/XLSX.
4. Escolha o estilo **Executivo**, **Essencial** ou **Escuro** no painel lateral. Ajuste cores, título, fonte, unidade, baseline e valores.
5. Clique em **Inserir no slide**. O gráfico se ajusta ao tamanho real do slide, centralizado em 16:9, com margem proporcional. O destino permanece no slide em que o editor foi aberto.
6. Para editar depois, selecione o grupo completo e clique em **Editar dados**. Altere os dados e clique em **Atualizar gráfico**.
7. Salve a apresentação. Use **Salvar projeto** para guardar também um backup JSON.

A atualização preserva posição, tamanho e rotação do grupo, criando o novo antes de remover o anterior. Edições manuais dentro do grupo são substituídas pelo desenho do editor. O novo grupo vai para a frente dos objetos existentes; ajuste sua ordem no PowerPoint se necessário.

## Integração com o PowerPoint

A versão 0.6 mantém o editor aberto como uma janela vinculada ao PowerPoint, permitindo navegar nos slides. Cada editor fica associado à apresentação e ao slide de origem. Trocar de seleção não altera o destino. Se outro arquivo estiver ativo, a inserção/atualização é bloqueada até você voltar ao arquivo original. Slide excluído ou apresentação fechada exige abrir um novo editor.

| Comando da aba Mello Charts | Comportamento |
| --- | --- |
| Galeria de gráficos | Abre a seleção visual dos 11 modelos |
| Build-up / Cascata / Comparativo | Abre um novo gráfico para o slide atual |
| Editar dados | Aceita o grupo ou elementos de um único gráfico |
| Duplicar gráfico | Usa a duplicação nativa do PowerPoint e atribui identidade própria à cópia |
| Ajustar ao slide | Centraliza o grupo em 16:9, remove a rotação e preserva seus dados e elementos |
| Diagnóstico | Verifica DLL e registros próprios, sem alterar configurações ou ativar o Office |

Os botões de edição ficam indisponíveis sem um gráfico compatível selecionado. Os comandos possuem ícones distintos. Operações de inserir, atualizar, duplicar e ajustar iniciam uma entrada no histórico de desfazer do PowerPoint. O comportamento final de desfazer deve fazer parte do ensaio no Office real.

O editor confirma o fechamento de alterações não aplicadas nem salvas em JSON. No encerramento do host ou da janela nativa de origem, tenta salvar o último estado válido em `%LOCALAPPDATA%\MelloCharts\Drafts`. Use **Recuperar rascunho** no editor para abrir esses JSON. A recuperação depende de permissão de gravação local e não cobre encerramento abrupto do processo nem conteúdo inválido ainda não aceito. Não substitui salvar a apresentação e seu projeto.

Veja o [guia da integração](native-integration.md).

## Aparência e exemplos

O painel lateral reúne os ajustes em um painel lateral e amplia a prévia. O desenho usa títulos com hierarquia, paletas coordenadas, totais destacados, rótulos alinhados às barras e intervalos legíveis no eixo.

![Build-up no estilo Executivo](../examples/gallery/buildup-executivo.png)

| Estilo | Uso sugerido | Modelo pronto |
| --- | --- | --- |
| Executivo | Reuniões de acompanhamento e comparações financeiras | [Build-up](../examples/gallery/buildup-executivo.json) |
| Essencial | Relatórios com menos elementos decorativos | [Cascata](../examples/gallery/cascata-essencial.json) |
| Escuro | Apresentações com fundo escuro e cores adaptadas | [Comparativo](../examples/gallery/comparativo-escuro.json) |

Abra qualquer JSON pelo botão **Abrir projeto**. Os modelos usam dados fictícios. Veja também a [cascata](../examples/gallery/cascata-essencial.png), o [comparativo](../examples/gallery/comparativo-escuro.png) e o [guia de aparência](visual-guide.md).

**Abreviar números** mostra, por exemplo, `1.500` como `1,5 mil`, sempre na unidade informada. A tabela, o CSV, o JSON e os cálculos preservam os valores originais. É uma opção de apresentação, sem conversão de moeda ou unidade.

Use **Exportar PNG** para salvar uma imagem em **1920 × 1080**, desenhada pelo mesmo renderizador da prévia. Para manter formas e textos editáveis no PowerPoint, use **Inserir no slide**. O JSON guarda os dados e as opções de aparência. Projetos e grupos da versão 0.3 continuam sendo aceitos.

## Tipos de gráfico

| Tipo | Como funciona |
|---|---|
| Build-up | Realizado e projetado restante empilhados na mesma coluna; EAC é a soma das duas parcelas. Baseline é uma referência e há total opcional. |
| Cascata | Base inicial, variações positivas e negativas, subtotais e total automático. Na coluna Tipo, use `base`, `delta` ou `total`. |
| Comparativo | Realizado e projetado em colunas independentes, com baseline opcional em linha. Aceita valores negativos. |
| Barras horizontais | Duas barras por categoria e marcador opcional de baseline. Aceita negativos. |
| Linhas | Trajetórias com marcadores, baseline opcional e rótulos diretos nos últimos pontos. Aceita negativos. |
| Área comparativa | Áreas sobrepostas de realizado e projetado; não soma as séries. Valores não negativos. |
| Colunas + linha | Realizado em colunas e projetado em linha na mesma unidade e escala. Aceita negativos. |
| Empilhado 100% | Realizado + projetado restante normalizados por categoria. Parcelas não negativas. |
| Pizza / Rosca | Composição de uma série não negativa com total positivo. A rosca mostra o total no centro. |
| Pareto | Barras ordenadas por valor, curva acumulada e referência em 80%. Preserva a ordem da tabela no projeto. |

Os exemplos são fictícios. A unidade é um rótulo, sem conversão automática: um valor `320` com unidade `R$ mil` representa R$ 320 mil. As casas decimais afetam a exibição, sem arredondar os cálculos do modelo.

Os gráficos são **grupos editáveis de formas, textos, elipses e polígonos**. Para alterar a tabela, use o editor Mello Charts. Eles não usam o objeto de gráfico do Office com planilha **Editar Dados**.

Pizza, rosca e Pareto exigem valores não negativos e total positivo. Área e empilhado 100% usam parcelas não negativas. Na área, as séries ficam sobrepostas sem serem somadas. No combinado, as duas séries usam a mesma unidade e escala. No empilhado, **realizado + projetado restante** é a base de 100% de cada categoria; uma linha toda zero aparece sem base.

Na pizza e na rosca, parcelas positivas abaixo de 0,1% aparecem como **<0,1%**. Fatias imperceptíveis são omitidas do desenho para evitar polígonos degenerados; seus valores permanecem na legenda, no total e no projeto.

O Pareto ordena somente o desenho e mostra o acumulado. A linha de 80% é uma referência fixa, não uma meta calculada.

## Dados de entrada

- **Build-up / Comparativo / Barras / Linhas / Área / Colunas + linha / Empilhado 100%:** `categoria;realizado;projetado;baseline`
- **Cascata:** `categoria;valor;tipo`
- **Pizza / Rosca / Pareto:** `categoria;valor`
- Até 12 categorias por gráfico e números individuais com magnitude de até 1 bilhão.
- CSV aceita ponto e vírgula, vírgula ou tabulação, aspas e números brasileiros. Para colar do Excel, inclua o cabeçalho.
- No XLSX, usa a primeira aba, com o cabeçalho na primeira linha. Fórmulas usam resultados gravados no arquivo; células com erro ou fórmula sem resultado salvo são rejeitadas. Arquivos `.xls`, macros e planilhas protegidas não são suportados.
- O JSON continua na versão de formato 1. A versão 0.6 abre projetos anteriores; os novos tipos não são reconhecidos por versões antigas.
- No CSV de cascata, linhas `total` exportam valor `0` como marcador. Seu valor é sempre recalculado a partir da base e das variações.

Modelos em [examples](../examples). Os dados do gráfico são guardados em tags no próprio grupo; salvar e copiar esse grupo preserva seus metadados. Arquivos gerados por outras ferramentas não se tornam editáveis pelo Mello Charts automaticamente. Desagrupar ou remover os metadados impede a recuperação pelo suplemento.

## Remover

Feche o PowerPoint e use **Configurações do Windows → Aplicativos → Mello Charts Desktop → Desinstalar**, ou abra o instalador e clique em **Remover**. Se estiver usando a cópia instalada para remover, o próprio executável pode permanecer na pasta até você fechar a janela. A mensagem final informa esse caminho. As apresentações e projetos não são removidos.

## Compilar e testar

Não há dependências NuGet para compilar. Use Windows com .NET Framework 4.x e o compilador `csc.exe` incluído nesse framework:

```powershell
./build.ps1 -RunTests -RenderGallery
```

O comando compila a DLL, embute-a em `release/MelloCharts-Setup.exe` e executa testes locais. **Não instala nem registra o suplemento.** Se a execução de scripts for controlada pela organização, siga o procedimento autorizado da TI para compilação.

```text
src/ChartCore.cs          Validação, números, JSON/CSV/XLSX e desenhos originais
src/ChartCatalog.cs       Definições dos 11 tipos e esquemas
src/AdvancedCharts.cs     Oito modelos adicionais e geometria
src/ChartGalleryForm.cs   Galeria nativa de modelos
src/PowerPointAddIn.cs    Ribbon, ponte COM e metadados por grupo
src/NativeIntegration.cs Contexto de apresentação, seleção, cópia, ajuste e rascunhos
src/InstallationHealth.cs Diagnóstico local e relatório de instalação
src/EditorForm.cs         Editor WinForms e prévia local
installer/Setup.cs       Instalação/remoção por usuário e rollback
tests/                   Testes de núcleo, COM e ensaio opcional no Office
assets/                  Ícones próprios
examples/                Dados fictícios para começar
release/                 Instalador pronto para baixar
```

O teste opcional `tests/OfficeSmoke.cs` cria apenas uma apresentação de exemplo e exercita desenho/atualização via COM, sem instalar o suplemento. Ele exige permissão para automatizar o PowerPoint; não foi concluído neste ambiente.

## Escopo desta versão

Esta versão cobre os 11 gráficos acima. Ainda não inclui Gantt, Mekko, vínculo automático com Excel, arrastar rótulos no editor ou todos os recursos do think-cell. A interface e o código são próprios, com referência aos fluxos públicos de [edição de gráficos](https://www.think-cell.com/en/resources/manual/introductioncharting) e [cascata](https://www.think-cell.com/en/resources/manual/waterfall). O instalador/chave/binários do think-cell não fazem parte deste projeto.

Código próprio sob [licença MIT](../LICENSE). APIs do Windows, .NET e PowerPoint são utilizadas pela instalação existente no computador, sem redistribuir binários da Microsoft.
