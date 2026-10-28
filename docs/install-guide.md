# Instalar, atualizar ou remover

1. Baixe o [instalador0.8.0](../release/MelloCharts-Setup.exe).
2. Salve seus documentos e feche Excel e PowerPoint.
3. Execute o instalador e use **Instalar**. A instalação é por usuário, em `%LOCALAPPDATA%\MelloCharts`.
4. Abra os aplicativos desktop e confira a aba **Mello Charts**.

Requisitos: Windows 10/11 Intel/AMD, .NET Framework 4.8, Office desktop 32/64 bits para a integração. O instalador registra classesCOM diferentes para Excel e PowerPoint nas views32/64, sem alterar políticas de segurança.

Para usar sem registrar abas, extraia todo o [pacote portátil](../release/MelloCharts-Portable.zip) e abra `MelloCharts.exe` ao lado da DLL. O editor importa/exporta gráficos; o envio ao slide exige PowerPoint instalado.

O instalador não possui assinatura digital comercial. Em ambiente gerenciado, siga o processo aprovado pela TI. **Verificar/Diagnóstico** apenas lê arquivo, versão, classesCOM e LoadBehavior; não comprova que o Office ativou a aba.

Para atualizar, feche ambos os aplicativos e execute o novo instalador. Para remover, use **Remover** no instalador ou Mello Charts Desktop em Configurações→Aplicativos. A remoção é limitada aos arquivos/registros próprios; não remove PPTX/XLSX/JSON/CSV do usuário.

Confira o SHA-256 do pacote com `Get-FileHash` e compare com [SHA256SUMS](../release/SHA256SUMS.txt) da mesma versão.

Instalação/ativação real e remoção ainda precisam de aceitação no computador de destino. A ponte COM foi validada diretamente com documentos novos no Excel/PowerPoint Microsoft 365. [Registro completo](verification.md).
