# Contribuir com o Mello Charts

O projeto é um suplemento COM experimental para PowerPoint desktop no Windows. Mudanças de código devem preservar os dados dos gráficos e a instalação por usuário.

## Preparar o ambiente

1. Use Windows com .NET Framework 4.8 e o compilador `csc.exe` incluído no framework.
2. Clone ou baixe o repositório, preservando suas pastas.
3. Execute `./build.ps1 -RunTests -RenderGallery` na raiz.

O código utiliza sintaxe compatível com C# 5 e bibliotecas do .NET Framework, sem pacotes NuGet. O script não instala nem registra componentes. O PowerPoint só é necessário para o ensaio real da integração.

## Propor uma alteração

- Descreva o problema e a mudança de comportamento.
- Inclua um exemplo com dados fictícios. Não envie arquivos de clientes ou empresas.
- Rode os testes pertinentes e o build completo antes de abrir o pull request.
- Para mudanças visuais, confira as imagens geradas em `examples/gallery`.
- Para mudanças de COM, instalador, metadados ou editor, confira também falhas e preservação dos originais.
- Informe explicitamente se o teste no PowerPoint real foi executado ou se continua pendente.

Não inclua `build/`, rascunhos pessoais, certificados, credenciais, binários de terceiros ou chaves de licença. Os instaladores publicados devem corresponder ao código da versão e ter seu hash atualizado.

## Testar no Office

O ensaio `test-office.ps1` é opcional e ativa o PowerPoint por COM para criar uma apresentação separada. Leia [Verificação](docs/verification.md) antes de executar. Use apresentações de teste e respeite as permissões da organização.

## Publicar uma versão

Atualize as versões da DLL e do instalador juntas, rode o build e confira o payload embutido. Publique o EXE e `SHA256SUMS.txt` como ativos da Release. Veja [Publicação](docs/publish-guide.md).

Contribuições são distribuídas sob a [licença MIT](LICENSE) do projeto.
