# Usar como suplemento no PowerPoint

## Instalar ou atualizar

1. Salve os documentos e feche todas as janelas do PowerPoint.
2. Abra `release/MelloCharts-Setup.exe` e confira a versão **0.6.0**.
3. Clique em **Verificar** se quiser conferir a instalação atual. Essa opção só lê o estado local.
4. Clique em **Instalar**. A atualização usa a mesma pasta e o mesmo usuário do Windows.
5. Abra o PowerPoint desktop e procure a aba **Mello Charts**. Faça o primeiro ensaio em uma apresentação de teste.

O suplemento é COM para Windows, com DLL AnyCPU e editor WinForms. O instalador configura os registros próprios nas views 32/64. O PowerPoint e o .NET Framework 4.8 devem estar instalados. O executável ainda não tem assinatura digital comercial. Políticas da organização podem exigir aprovação da TI.

## Trabalhar com o editor aberto

Na aba Mello Charts, abra **Galeria de gráficos**, selecione um dos 11 modelos e clique em **Usar modelo**. Os atalhos Build-up, Cascata e Comparativo continuam disponíveis. O editor é uma janela vinculada ao PowerPoint e permite continuar navegando na apresentação. Ele mostra o número do slide em que foi aberto. A identidade do slide é mantida mesmo se você mudar a seleção.

**Inserir no slide** e **Atualizar gráfico** sempre usam esse destino de origem. Se outra apresentação estiver ativa, volte à original para aplicar. Se você excluir o slide ou fechar o arquivo original, salve seu JSON e abra um novo editor. O destino não é transferido automaticamente para outro arquivo.

Há um editor por instância do suplemento. Abrir outro gráfico fecha o editor anterior, com confirmação se houver alterações não aplicadas nem salvas em JSON. O editor não altera a apresentação enquanto você apenas modifica a tabela ou a aparência.

## Escolher o modelo

Compare as miniaturas e confirme pelo botão ou dê duplo clique no cartão. A tabela e as opções mudam conforme o tipo: pizza, rosca e Pareto pedem categoria e valor; os comparativos usam realizado, projetado e baseline. Consulte o [guia visual](visual-guide.md).

Os gráficos são grupos de formas e textos, incluindo elipses e polígonos. Use **Editar dados** do Mello Charts para a tabela. O comando de edição dos gráficos nativos do Office não é usado.

## Editar, copiar e ajustar

- **Editar dados:** selecione o grupo ou elementos de um único gráfico. Seleções que misturem gráficos ou objetos externos são rejeitadas.
- **Duplicar gráfico:** duplica o grupo nativo do PowerPoint, mantém suas personalizações e atribui uma identidade própria à cópia. A posição recebe um pequeno deslocamento, limitado ao slide.
- **Ajustar ao slide:** centraliza, redefine a proporção para 16:9 e remove a rotação. Preserva a identidade, os dados e os elementos nativos do grupo.
- **Desfazer:** as operações iniciam uma entrada no histórico do PowerPoint. Confira o comportamento no ensaio real antes de usar em um arquivo de trabalho.

Atualizar dados redesenha o gráfico a partir do modelo salvo. Ajustes manuais internos de texto, cor e posição são substituídos nessa atualização. A duplicação e o ajuste de tamanho preservam o grupo nativo.

## Rascunhos e recuperação

Ao fechar normalmente o editor com alterações pendentes, você pode salvar o JSON, descartar ou cancelar o fechamento.

Se o PowerPoint encerrar a conexão ou destruir a janela de origem, o editor tenta guardar o **último estado válido** em `%LOCALAPPDATA%\MelloCharts\Drafts`. Essa gravação é local. A recuperação usa arquivos com nomes únicos para não substituir outro rascunho.

Para recuperar, abra um novo gráfico no slide desejado e clique em **Recuperar rascunho**. Escolha o JSON pela data do nome, confira os dados e insira no slide. Você também pode usar **Abrir projeto**.

O rascunho não garante recuperação após queda de energia, encerramento abrupto do processo ou bloqueio da pasta de gravação. Conteúdo inválido ainda não aceito não faz parte do último estado válido. Continue salvando a apresentação e o JSON. Os rascunhos contêm seus dados e podem ser excluídos quando você não precisar mais deles.

## Diagnóstico

**Mello Charts → Diagnóstico** abre um relatório local que pode ser salvo em TXT. O botão **Verificar** do instalador usa a mesma verificação.

São conferidos a versão da DLL, o runtime/processo, os registros próprios da classe COM e do suplemento, o CodeBase e o LoadBehavior. O relatório não lê conteúdo ou nomes das apresentações, não ativa o Office e não modifica segurança ou registro.

Se a aba não aparecer, veja **Arquivo → Opções → Suplementos → Gerenciar: Suplementos COM → Ir**. Registre a mensagem exibida pelo Office e consulte o [guia de instalação](install-guide.md). Encontrar DLL e registros não comprova que o Office ativou o suplemento.

## Validação disponível

O pacote passou por 183 verificações locais. A destruição da janela nativa foi reproduzida em uma janela invisível de teste, independente do Office. A ativação do suplemento instalado e o fluxo completo no PowerPoint real ainda não foram confirmados, conforme [Verificação](verification.md).

Referências de API: [desfazer no PowerPoint](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.application.startnewundoentry), [grupo pai de uma forma](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.shape.parentgroup), [localização por ID do slide](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.slides.findbyslideid) e [atualização da ribbon](https://learn.microsoft.com/en-us/office/vba/api/office.iribbonui.invalidate).
