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
