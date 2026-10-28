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

Guias: [Instalação](install-guide.md) · [Cores](custom-colors.md) · [Modelos e exemplos](../examples/gallery/README.md) · [Validação](verification.md).
