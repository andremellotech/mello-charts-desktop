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
