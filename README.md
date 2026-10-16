# Mello Charts Desktop

**Gráficos executivos no PowerPoint, com dados em tabela e elementos editáveis no slide.**

O Mello Charts Desktop é um suplemento gratuito e de código aberto para PowerPoint no Windows. Escolha um modelo, preencha ou importe seus dados, confira a prévia e insira o gráfico na apresentação.

![Versão experimental](https://img.shields.io/badge/versão-0.7.0%20experimental-orange)
![Windows](https://img.shields.io/badge/plataforma-Windows%2010%2F11-blue)
![Licença MIT](https://img.shields.io/badge/licença-MIT-green)

**[Baixar instalador](https://github.com/andremellotech/mello-charts-desktop/raw/refs/heads/main/release/MelloCharts-Setup.exe)** · [Instalação](docs/install-guide.md) · [Guia de uso](docs/native-integration.md) · [Galeria de exemplos](examples/gallery/README.md)

![Galeria dos 11 modelos de gráficos do Mello Charts](examples/gallery/catalogo.png)

*Prévias dos modelos disponíveis. A captura de tela ou demonstração do fluxo dentro do PowerPoint ainda está pendente.*

> [!WARNING]
> **Versão 0.7.0 experimental.** A ativação do suplemento instalado e o fluxo completo no PowerPoint real ainda precisam de validação. Os testes locais não substituem esse ensaio. Consulte o [registro de verificação](docs/verification.md).
>
> O instalador não possui assinatura digital comercial e pode gerar um aviso do SmartScreen. Em ambientes corporativos, siga o procedimento aprovado pela TI.

## Por que usar

- **11 modelos de gráficos:** de comparativos e cascatas a Pareto, pizza e rosca.
- **Dados em tabela:** digite, cole do Excel ou importe CSV, TSV e XLSX.
- **Prévia antes de inserir:** confira o resultado enquanto ajusta os dados e a aparência.
- **Elementos editáveis:** os gráficos são inseridos como grupos de formas e textos do PowerPoint.
- **Cores personalizadas:** defina séries, fundo e fatias por seletor, HEX ou RGB.
- **Funcionamento offline:** sem conta ou licença própria do suplemento.
- **Exportação e backup:** salve imagens PNG, projetos JSON e tabelas CSV.

## Navegação

[Requisitos](#requisitos) · [Instalação](#instalação) · [Primeiro gráfico](#seu-primeiro-gráfico) · [Modelos](#modelos-disponíveis) · [Edição](#editar-e-personalizar) · [Dados](#dados-e-importação) · [Perguntas frequentes](#perguntas-frequentes) · [Desenvolvimento](#desenvolvimento)

## Requisitos

| Componente | Requisito previsto |
| --- | --- |
| Sistema operacional | Windows 10 ou 11, Intel/AMD |
| Aplicativo | PowerPoint desktop para Windows, 32 ou 64 bits |
| Runtime | .NET Framework 4.8 |
| PowerPoint na Web | Não suportado |
| PowerPoint para Mac | Não suportado |

**Compatibilidade prevista não significa compatibilidade já testada.** Ainda não há versões do Office confirmadas por ensaio completo com o suplemento instalado.

O Mello Charts não exige licença própria. O uso do PowerPoint continua sujeito à licença Microsoft.

## Instalação

1. Baixe o **[MelloCharts-Setup.exe](https://github.com/andremellotech/mello-charts-desktop/raw/refs/heads/main/release/MelloCharts-Setup.exe)**.
2. Salve seus documentos e **feche todas as janelas do PowerPoint**.
3. Execute o instalador e clique em **Instalar**.
4. Abra o PowerPoint e selecione um slide na exibição **Normal**.
5. Procure a aba **Mello Charts**.

A instalação é feita por usuário. O instalador não altera as políticas de segurança do Office.

### Conferir a integridade do instalador

Opcionalmente, calcule o SHA-256 do arquivo baixado:

```powershell
Get-FileHash .\MelloCharts-Setup.exe -Algorithm SHA256
```

Compare o campo `Hash` com o valor do instalador no arquivo [`release/SHA256SUMS.txt`](release/SHA256SUMS.txt) **da mesma versão**.

### Atualizar

Salve seus documentos, feche o PowerPoint e execute o instalador da nova versão. Clique em **Instalar**.

### Desinstalar

1. Salve seus documentos e feche todas as janelas do PowerPoint.
2. Execute o instalador e clique em **Remover**.

Também é possível iniciar a desinstalação pela entrada **Mello Charts Desktop** em **Configurações → Aplicativos**. Na janela do instalador, clique em **Remover**.

A remoção foi projetada para excluir somente os arquivos e registros próprios do suplemento, preservando os PPTX, JSON e CSV salvos. Se o executável do desinstalador permanecer na pasta instalada, feche sua janela antes de removê-lo manualmente.

*O procedimento foi conferido no código e na documentação. O ciclo de instalação e remoção ainda precisa ser validado no ambiente real.*

## Seu primeiro gráfico

1. No PowerPoint, selecione o slide que receberá o gráfico.
2. Abra **Mello Charts → Galeria de gráficos**.
3. Escolha um modelo e confirme **Usar modelo**.
4. Preencha a tabela, cole dados do Excel ou importe um arquivo.
5. Ajuste o estilo, a paleta, a unidade e os rótulos.
6. Confira a prévia e clique em **Inserir no slide**.
7. Salve a apresentação e, como backup dos dados, exporte o projeto em JSON.

> [!IMPORTANT]
> Ao atualizar um gráfico pelo Mello Charts, seus elementos são redesenhados. Ajustes manuais feitos nos textos, barras ou outros elementos internos serão substituídos pelos dados e opções do editor.

## Modelos disponíveis

| Modelo | Uso principal |
| --- | --- |
| **Build-up** | Mostrar o realizado e o projetado restante por frente de trabalho. |
| **Cascata** | Explicar as variações entre um valor inicial e o resultado final. |
| **Comparativo** | Comparar realizado e projetado por período ou categoria. |
| **Barras horizontais** | Comparar duas séries por categoria, com baseline opcional. |
| **Linhas** | Mostrar trajetórias com marcadores e rótulos nos últimos pontos. |
| **Área comparativa** | Comparar áreas sobrepostas de realizado e projetado. |
| **Colunas + linha** | Combinar duas séries na mesma unidade e escala. |
| **Empilhado 100%** | Mostrar a participação do realizado e do projetado restante em cada categoria. |
| **Pizza** | Mostrar a composição de uma série. |
| **Rosca** | Mostrar a composição de uma série, com o total no centro. |
| **Pareto** | Ordenar valores e acompanhar sua participação acumulada. |

**Termos usados nos modelos:**

- **Realizado:** parcela já executada.
- **Projetado restante:** parcela prevista que ainda falta executar.
- **EAC — Estimate at Completion:** realizado + projetado restante.
- **Baseline:** referência de comparação; não entra na soma do EAC.

Todos os exemplos do repositório usam **dados fictícios**.

### Exemplos para explorar

| Exemplo | Prévia | Projeto |
| --- | --- | --- |
| Build-up Executivo | [Ver imagem](examples/gallery/buildup-executivo.png) | [JSON](examples/gallery/buildup-executivo.json) |
| Cascata Essencial | [Ver imagem](examples/gallery/cascata-essencial.png) | [JSON](examples/gallery/cascata-essencial.json) |
| Comparativo Escuro | [Ver imagem](examples/gallery/comparativo-escuro.png) | [JSON](examples/gallery/comparativo-escuro.json) |

Abra os arquivos JSON pelo botão **Abrir projeto**. A [galeria completa](examples/gallery/README.md) reúne PNG, CSV e JSON dos 11 modelos.

## Editar e personalizar

Os gráficos são inseridos como **grupos de formas, textos, elipses e polígonos**. Eles não usam o objeto de gráfico do Office nem sua planilha **Editar Dados**.

| Ação | Como funciona |
| --- | --- |
| **Editar dados** | Selecione o grupo ou um elemento de um único gráfico e abra o editor Mello Charts. |
| **Atualizar gráfico** | Aplica os dados e a aparência do editor, redesenhando os elementos internos. |
| **Duplicar gráfico** | Cria uma cópia nativa com identidade própria e preserva os ajustes feitos no PowerPoint. |
| **Ajustar ao slide** | Centraliza o grupo, mantém a proporção 16:9 e remove a rotação. |
| **Alterar aparência** | Escolha entre Executivo, Essencial e Escuro, com três paletas disponíveis. |
| **Personalizar cores** | Defina séries, fundo e fatias pelo seletor, HEX ou RGB. |
| **Exportar** | Salve PNG em 1920 × 1080, projeto JSON ou tabela CSV. |

O editor permanece vinculado à apresentação e ao slide de origem enquanto você navega pelos slides.

Para manter a recuperação dos dados pelo suplemento, **preserve o grupo e suas tags**. Mantenha também um JSON de backup.

[Guia de integração](docs/native-integration.md) · [Cores personalizadas](docs/custom-colors.md) · [Guia visual](docs/visual-guide.md)

## Dados e importação

Digite os dados no editor, use **Colar do Excel** com o cabeçalho ou importe arquivos **CSV, TSV e XLSX**.

Cada gráfico aceita até **12 categorias**.

### Estrutura das tabelas

| Modelos | Cabeçalho |
| --- | --- |
| Build-up, Comparativo, Barras, Linhas, Área, Colunas + linha e Empilhado 100% | `categoria;realizado;projetado;baseline` |
| Cascata | `categoria;valor;tipo` |
| Pizza, Rosca e Pareto | `categoria;valor` |

Na cascata, a coluna `tipo` aceita:

| Valor | Significado |
| --- | --- |
| `base` | Ponto de partida |
| `delta` | Variação positiva ou negativa |
| `total` | Subtotal ou resultado calculado automaticamente |

### Regras dos modelos

- **Barras horizontais, linhas e colunas + linha:** aceitam valores negativos.
- **Pizza, rosca e Pareto:** exigem valores não negativos e total positivo.
- **Área comparativa:** exige parcelas não negativas; as áreas ficam sobrepostas, sem somar as séries.
- **Empilhado 100%:** exige parcelas não negativas. Realizado + projetado restante formam a base de 100% de cada categoria; uma linha toda zero aparece sem base.
- **Colunas + linha:** as duas séries compartilham a mesma unidade e escala.
- **Pareto:** ordena o desenho e preserva a ordem original no projeto. A referência de 80% é fixa.

Na pizza e na rosca, parcelas positivas abaixo de 0,1% recebem o rótulo **<0,1%**. Fatias imperceptíveis podem ser omitidas do desenho para evitar polígonos degenerados; seus valores permanecem na legenda, no total e no projeto.

### Unidades e precisão

A unidade é um rótulo, sem conversão automática. Por exemplo, `320` com a unidade `R$ mil` representa **R$ 320 mil**.

As opções de casas decimais e abreviação alteram a apresentação dos números, preservando os valores do modelo.

## Perguntas frequentes

### Quais versões do Office foram testadas?

Ainda não há versões do Office confirmadas por ensaio completo com o suplemento instalado.

O alvo é o PowerPoint desktop de 32 ou 64 bits no Windows 10/11. Microsoft 365, Office 2019 e Office 2016 **não devem ser apresentados como versões já testadas**.

No ensaio documentado, a automação do PowerPoint retornou `E_ACCESSDENIED` (`80070005`) antes do desenho. Os detalhes estão em [Verificação](docs/verification.md).

### O que acontece ao abrir a apresentação sem o suplemento?

Como os gráficos são gravados como formas e textos do PowerPoint, o comportamento esperado é que continuem visíveis na apresentação e permitam edição manual dos elementos.

Para editar a tabela pelo Mello Charts ou atualizar o gráfico pelo editor, é necessário ter o suplemento instalado e carregado.

**Essa expectativa é baseada na implementação. O ciclo de abrir, apresentar, editar e salvar sem o suplemento ainda precisa de validação no PowerPoint real.**

### Posso editar os gráficos pela opção “Editar Dados” do PowerPoint?

Não. Essa opção pertence aos objetos de gráfico do Office.

Para alterar a tabela de um gráfico Mello Charts, selecione-o e use **Editar dados** na aba do suplemento.

### Meus ajustes manuais são preservados?

A duplicação nativa foi projetada para preservar os ajustes existentes.

Já a atualização pelo editor redesenha os elementos internos e substitui ajustes manuais de barras, textos e outros componentes. Salve a apresentação antes de atualizar e mantenha um JSON de backup.

### A aba Mello Charts não apareceu. O que verificar?

No PowerPoint, abra **Arquivo → Opções → Suplementos**. Em **Gerenciar**, selecione **Suplementos COM → Ir** e confira se o Mello Charts aparece ou se há uma mensagem de erro.

Use também **Verificar** no instalador. A presença de arquivos e registros, por si só, não comprova que o Office carregou o suplemento.

Consulte o [guia de instalação](docs/install-guide.md). Em computadores corporativos, siga as orientações da TI.

### Preciso de internet ou de uma conta?

O suplemento funciona offline e não exige conta. A licença do PowerPoint é independente.

### Como reportar um problema?

Abra uma issue com:

- versão do Mello Charts;
- versão do Windows;
- versão e build do PowerPoint, incluindo 32 ou 64 bits;
- passos para reproduzir;
- resultado esperado e resultado observado;
- mensagem de erro e captura de tela, quando disponíveis.

Se precisar anexar uma apresentação, use uma cópia com dados fictícios ou anonimizados.

## Privacidade e segurança

- O suplemento funciona offline e não exige conta.
- O diagnóstico consulta arquivos e registros próprios, sem ler o conteúdo das apresentações.
- O instalador não altera políticas de segurança do Office.
- O código está disponível sob licença MIT para consulta e auditoria.
- Os binários desta versão não possuem assinatura digital comercial.

## Estado da versão

A versão **0.7.0 é experimental**. O projeto possui verificações locais do núcleo, da ponte COM, dos modelos, do editor, das cores e dos contratos de integração.

Ainda precisam de ensaio no PowerPoint real:

- instalação, carregamento e remoção do suplemento;
- inserção e fidelidade visual dos gráficos;
- navegação com o editor aberto;
- comportamento do histórico de desfazer;
- ciclo de salvar e reabrir;
- abertura e edição da apresentação sem o suplemento;
- captura de tela ou GIF do fluxo completo.

Consulte [o que foi verificado e o que falta](docs/verification.md) antes de adotar o suplemento em uma apresentação de trabalho.

## Limitações e próximos passos

Esta versão oferece os 11 modelos descritos acima. Ainda não inclui:

- vínculo automático com Excel;
- gráficos Gantt e Mekko;
- recurso específico para arrastar rótulos;
- equivalência funcional completa com ferramentas comerciais.

## Desenvolvimento

No Windows, com o .NET Framework instalado, execute na raiz do repositório:

```powershell
./build.ps1 -RunTests -RenderGallery
```

Não há dependências NuGet.

O comando compila a DLL e o instalador, executa as verificações locais, renderiza a galeria e gera o SHA-256 do instalador. Ele **não instala nem registra o suplemento e não abre o PowerPoint**. Alguns testes usam janelas nativas invisíveis e rascunhos temporários.

A versão documenta **203 verificações**:

| Área | Verificações |
| --- | ---: |
| Núcleo | 52 |
| Ponte | 38 |
| Novos gráficos | 45 |
| Editor e galeria | 27 |
| Cores | 20 |
| Integração | 21 |
| **Total** | **203** |

O workflow em **Actions** compila e testa no Windows. Os instaladores produzidos no CI são experimentais e não possuem assinatura digital comercial. A validação no Office é um ensaio separado.

### Estrutura do projeto

```text
src/          Cálculos, editor, integração COM e diagnóstico
installer/    Instalação e remoção por usuário
tests/        Testes locais e ensaio opcional no Office
examples/     Tabelas e projetos com dados fictícios
assets/       Ícones próprios
docs/         Guias, referência técnica e notas de versão
release/      Instalador e SHA256SUMS.txt
.github/      Workflows e modelos de colaboração
```

Para contribuir, leia [CONTRIBUTING.md](CONTRIBUTING.md).

## Documentação

- [Instalação e diagnóstico](docs/install-guide.md)
- [Integração com o PowerPoint](docs/native-integration.md)
- [Aparência e estilos](docs/visual-guide.md)
- [Cores personalizadas](docs/custom-colors.md)
- [Referência completa](docs/reference.md)
- [Galeria de exemplos](examples/gallery/README.md)
- [Verificação e pendências](docs/verification.md)
- [Histórico de alterações](CHANGELOG.md)

## Licença

O código próprio é distribuído sob a [licença MIT](LICENSE).

Nenhum instalador, binário, chave de licença ou código do think-cell é distribuído. Os componentes da Microsoft utilizados pertencem à instalação existente no computador e não são redistribuídos pelo projeto.
