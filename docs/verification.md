# Verificação da versão 0.5.0

Conferido em 15 de outubro de 2026, Windows, com compilador C# do .NET Framework.

## Concluído

- Compilação da DLL AnyCPU e do instalador EXE com a DLL e os ícones próprios embutidos.
- **52 testes do núcleo**, abrangendo números brasileiros/internacionais, precisão, validação, JSON estrito, CSV/TSV, fórmulas com cache no XLSX, limites de leitura/descompressão, cálculos dos três tipos, totais, subtotais e ajuste de texto nos quadros de desenho.
- **25 verificações da ponte**, incluindo metadados em tags com acentos preservados, XML e callbacks da ribbon, interfaces COM dual e `QueryInterface` real do objeto .NET, precisão ao abrir o editor e contratos de criação/atualização com um adaptador de formas.
- Os contratos de falha verificam preservação do original quando uma nova forma não pode ser nomeada ou os metadados não podem ser salvos. Falha na exclusão final mantém as duas versões e é informada.
- Revisão independente da ponte COM, callbacks, registro HKCU nas views 32/64, limites de remoção, rollback do instalador e editor.
- Resolução dos nomes dos callbacks pelo `IDispatch` real do objeto .NET, sem registro ou instalação do suplemento.
- Hash e integridade do pacote e do payload embutido conferidos antes da entrega.

- Galeria de três gráficos renderizada em PNG pelo mesmo código da prévia e exportação. Imagens inspecionadas para alinhamento, escala, contraste, legendas e destaque do consolidado.
- Novos testes: compatibilidade com projetos antigos, persistência de estilos, números abreviados sem perda dos dados, eixo com passos convencionais e fracionários, quadros dos rótulos de colunas e limites com 12 categorias.
- Contraste dos valores externos da cascata conferido nas nove combinações de estilo e paleta. Referência zero mantida quando a grade é ocultada.
- Leitura do editor e importação preservam estilo, abreviação, grade e precisão. A revisão independente identificou o desalinhamento de rótulos de colunas, corrigido e coberto por regressão para 1, 4 e 12 categorias.

- **21 verificações de integração** com adaptadores e janelas nativas invisíveis: dimensões do slide, vínculo ao slide original, proteção entre apresentações, seleção de elementos, duplicação nativa, geometria/rollback, estados da ribbon, ícones, diagnóstico e rascunhos.
- A regressão de fechamento do proprietário nativo reproduz a destruição do HWND de uma janela invisível independente do Office. Confere um único JSON válido salvo e o descarte posterior do editor, sem reentrância na destruição nativa.
- A duplicação usa `Shape.Duplicate` e verifica os metadados antes de concluir. O teste local cobre preservação de personalizações pelo contrato de duplicação nativa e remoção da cópia em falha, sem excluir o original.
- Revisão independente encerrou os dois achados importantes: destruição do dono nativo sem FormClosing e duplicação que redesenhava em vez de copiar o grupo.

## Não confirmado neste ambiente

O teste `OfficeSmoke` tentou criar uma apresentação separada por `PowerPoint.Application`. A ativação retornou `E_ACCESSDENIED` (80070005) antes de qualquer desenho. Não foi tentado contornar esse bloqueio.

Assim, **a ativação do suplemento instalado, os desenhos no PowerPoint real, a navegação com o editor aberto, a resposta do histórico de desfazer, a fidelidade visual da interface aberta e o ciclo salvar/reabrir da apresentação ainda precisam do ensaio no computador de destino**. Testes simulados da ponte não equivalem à validação do Office.

O instalador não foi executado para instalar/remover e nenhuma chave do suplemento foi criada neste computador. A revisão de instalação/rollback foi estática. Não houve assinatura digital comercial dos binários.

## Ensaio de aceitação

1. Instalar explicitamente pelo EXE com o PowerPoint fechado.
2. Confirmar a aba Mello Charts.
3. Inserir os três tipos, testar os estilos e conferir rótulos, totais e linha de baseline. Exportar PNG e comparar com o slide.
4. Mover, redimensionar e girar o grupo, então editar e atualizar.
5. Salvar/reabrir e recuperar os dados pelo botão Editar dados.
6. Copiar o grupo para outro slide/apresentação e conferir os metadados.
7. Com duas apresentações abertas, testar o bloqueio de atualização no arquivo errado e a recuperação ao fechar a janela de origem.
8. Conferir Diagnóstico e Verificar. Remover pelo instalador e conferir que as apresentações continuam intactas.

Use inicialmente apresentações de teste. O [guia de instalação](install-guide.md) contém os passos para diagnóstico caso o Office não carregue o suplemento.
