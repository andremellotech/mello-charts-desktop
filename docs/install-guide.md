# Instalar o Mello Charts Desktop

## Instalação

1. Baixe `MelloCharts-Setup.exe` do pacote entregue ou da área Releases do repositório, quando publicada.
2. Salve seus arquivos e feche todas as janelas do PowerPoint.
3. Abra o instalador. Abrir a janela ainda não instala nem remove nada.
4. Confira o produto **Mello Charts Desktop 0.6.0** e a pasta indicada.
5. Se desejar, clique em **Verificar** para conferir DLL e registros sem alterar nada. Clique em **Instalar** e aguarde a confirmação.
6. Abra o PowerPoint desktop para Windows e uma apresentação.
7. Procure a aba **Mello Charts** e crie um gráfico em um slide de teste.

É uma instalação por usuário. Os arquivos ficam em `%LOCALAPPDATA%\MelloCharts`; o registro usa somente as chaves próprias do suplemento em HKCU. O instalador configura as views COM 32/64, permitindo que a DLL AnyCPU seja usada pelo PowerPoint da mesma arquitetura.

Não é um suplemento web nem um arquivo PPAM com macros. Não precisa habilitar macros ou mudar configurações de confiança. O EXE ainda não tem assinatura digital comercial. Se houver bloqueio por política corporativa, utilize o processo aprovado pela TI.

## Primeiro uso

1. Selecione um slide na exibição **Normal**.
2. Clique em **Mello Charts → Build-up**.
3. Altere um valor, confira a prévia e clique em **Inserir no slide**.
4. Continue navegando no PowerPoint com o editor aberto. O destino continua no slide de origem. Para abrir outro editor, feche o atual, salvando alterações se necessário.
5. Selecione o grupo, uma barra ou um texto dele e clique em **Editar dados**.
6. Altere outro valor e clique em **Atualizar gráfico**.
7. Salve, feche e abra a apresentação novamente. Confira se o grupo e os dados continuam disponíveis.
8. Use **Duplicar gráfico** e **Ajustar ao slide**. Salve/reabra, recupere os dados e teste também o Desfazer do PowerPoint.

Faça esse ensaio antes de usar em uma apresentação de trabalho. A ativação instalada ainda não foi validada neste ambiente.

## Diagnóstico

Use **Verificar** no instalador ou **Mello Charts → Diagnóstico** no PowerPoint. O relatório verifica a DLL local, sua versão, classe COM, CodeBase e LoadBehavior nas views 32/64. Não lê conteúdo ou nomes de apresentações, não ativa o Office e não modifica configurações. A presença de arquivos e registros não comprova que o Office carregou o suplemento.

## Se a aba não aparecer

Abra **Arquivo → Opções → Suplementos**. Em **Gerenciar**, escolha **Suplementos COM** e **Ir**. Confira se **Mello Charts** aparece e se existe uma mensagem de erro de carregamento. Se o Office tiver desabilitado o suplemento por política ou erro, informe a mensagem para diagnóstico. Não altere as políticas de segurança para forçar o carregamento.

Se não aparecer na lista, feche o PowerPoint e execute novamente o instalador com o mesmo usuário do Windows que utiliza o Office. Confira Windows desktop Intel/AMD, .NET Framework 4.8 e PowerPoint desktop 32/64 bits.

## Atualizar ou remover

Para atualizar, feche o PowerPoint, execute o novo instalador e clique em **Instalar**. Para remover, use **Remover** no instalador ou a entrada Mello Charts Desktop em **Configurações → Aplicativos**. A remoção só afeta arquivos e registros do suplemento. Seus PPTX, JSON e CSV permanecem no lugar em que foram salvos.

Quando o próprio desinstalador estiver em execução dentro da pasta instalada, ele pode permanecer como arquivo residual. Feche sua janela e remova esse arquivo manualmente se desejar. Não é deixado serviço ou tarefa agendada de limpeza.

## O que é preservado nos gráficos

O editor salva o documento do gráfico no próprio grupo do PowerPoint, em tags codificadas para manter acentos e maiúsculas/minúsculas. A atualização cria e verifica o novo grupo antes de excluir o anterior. Cópias do grupo levam esses dados junto.

Não desagrupe nem apague as tags se quiser recuperar os dados. Se fizer edições manuais de textos ou barras, elas serão substituídas na próxima atualização pelo modelo do editor. Use **Salvar projeto** para manter um JSON de backup.

Referência técnica de integração: [Microsoft — suplemento COM com ribbon](https://learn.microsoft.com/en-us/office/vba/library-reference/concepts/customize-the-office-fluent-ribbon-by-using-a-managed-com-add-in).
