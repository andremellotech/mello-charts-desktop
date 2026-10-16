# Mello Charts Desktop

Suplemento gratuito para criar gráficos executivos no **PowerPoint para Windows**, com dados em tabela, prévia visual e formas editáveis no slide.

![Versão](https://img.shields.io/badge/versão-0.7.0%20experimental-orange)
![Plataforma](https://img.shields.io/badge/plataforma-Windows%2010%2F11-blue)
![Licença](https://img.shields.io/badge/licença-MIT-green)

**[Baixar instalador](https://github.com/andremellotech/mello-charts-desktop/raw/refs/heads/main/release/MelloCharts-Setup.exe) · [Como instalar](docs/install-guide.md) · [Guia de uso](docs/native-integration.md) · [Exemplos](examples/gallery)**

![Galeria dos 11 modelos Mello Charts](examples/gallery/catalogo.png)

**Demonstração no PowerPoint:** captura de tela ou GIF do fluxo Galeria → preencher tabela → Inserir no slide ainda pendente. A imagem acima mostra a galeria de modelos; não comprova a execução do suplemento dentro do Office.

> [!WARNING]
> **Versão 0.7.0 experimental.**
> - A ativação do suplemento instalado e o fluxo completo dentro do PowerPoint real **ainda precisam ser validados**. Veja [o que foi verificado e o que falta](docs/verification.md).
> - O instalador **não tem assinatura digital comercial**; o Windows pode exibir um aviso do SmartScreen.
> - Atualizar um gráfico redesenha seus elementos e **substitui ajustes manuais internos**. Salve a apresentação e mantenha um JSON de backup.

## Sumário

[Requisitos](#requisitos) · [Instalar e começar](#instalar-e-começar) · [O que você pode criar](#o-que-você-pode-criar) · [Trabalhar no PowerPoint](#trabalhar-no-powerpoint) · [Dados e modelos](#dados-e-modelos) · [Privacidade](#privacidade-e-segurança) · [Perguntas frequentes](#perguntas-frequentes) · [Limites e roadmap](#limites-e-roadmap) · [Para desenvolver](#para-desenvolver)

## Requisitos

- Windows 10/11 Intel/AMD
- .NET Framework 4.8
- PowerPoint desktop de 32 ou 64 bits
- **Não funciona** no PowerPoint na Web nem no Mac

Esses são os requisitos previstos. As versões do Office ainda precisam de validação com o suplemento instalado.

O suplemento funciona offline e não exige conta ou licença própria. O uso do PowerPoint continua sujeito à licença Microsoft.

## Instalar e começar

1. Baixe o [MelloCharts-Setup.exe](https://github.com/andremellotech/mello-charts-desktop/raw/refs/heads/main/release/MelloCharts-Setup.exe). Na página do arquivo, use **Download raw file**. Você também pode baixar os arquivos publicados na seção **Releases**, quando disponíveis.
2. Opcionalmente, confira a integridade do arquivo e compare com o valor em `release/SHA256SUMS.txt` da mesma versão:

   ```powershell
   Get-FileHash .\MelloCharts-Setup.exe -Algorithm SHA256
