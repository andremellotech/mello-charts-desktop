# Mello Charts Desktop

Suplemento gratuito para criar gráficos financeiros no **PowerPoint para Windows**, com dados em tabela, prévia visual e formas editáveis no slide.

**[Baixar instalador](release/MelloCharts-Setup.exe) · [Como instalar](docs/install-guide.md) · [Guia de uso](docs/native-integration.md) · [Exemplos](examples/gallery)**

![Build-up de investimento por frente de trabalho, com realizado, projetado restante e total](examples/gallery/buildup-executivo.png)

**Versão 0.5.0 experimental.** Passou por 98 verificações locais. A ativação do suplemento instalado e o fluxo completo dentro do PowerPoint real ainda precisam ser validados. O instalador não tem assinatura digital comercial. Veja [o que foi verificado e o que falta](docs/verification.md).

## O que você pode criar

| Gráfico | Para responder |
| --- | --- |
| **Build-up** | Quanto já foi realizado e quanto ainda falta projetar em cada frente? |
| **Cascata** | Quais variações levam do baseline ao EAC? |
| **Comparativo** | Como realizado e projetado se comparam por período ou categoria? |

No build-up, **EAC = realizado + projetado restante**. O baseline é uma referência e não entra na soma. Todos os exemplos deste repositório usam dados fictícios.

## Instalar e começar

1. Baixe [MelloCharts-Setup.exe](release/MelloCharts-Setup.exe). Na página do arquivo, use **Download raw file**. Você também pode baixar os arquivos publicados na seção **Releases**.
2. Salve os documentos e **feche todas as janelas do PowerPoint**.
3. Abra o instalador e clique em **Instalar**. Para atualizar, use o mesmo procedimento.
4. Abra o PowerPoint desktop, selecione um slide em exibição **Normal** e procure a aba **Mello Charts**.
5. Escolha um gráfico, preencha a tabela ou importe os dados e clique em **Inserir no slide**.

Requer **Windows 10/11 Intel/AMD, .NET Framework 4.8 e PowerPoint desktop de 32 ou 64 bits**. Não funciona no PowerPoint na Web ou no Mac. O suplemento funciona offline e não exige uma conta ou licença própria. O uso do PowerPoint continua sujeito à licença Microsoft.

Se a instalação for bloqueada por uma política da organização, use o procedimento aprovado pela TI. O instalador não altera as políticas de segurança do Office.

## Trabalhar no PowerPoint

- **Editor aberto:** continue navegando nos slides. O destino permanece vinculado à apresentação e ao slide de origem.
- **Editar dados:** selecione o grupo ou uma barra/texto de um único gráfico.
- **Duplicar gráfico:** cria uma cópia nativa com identidade própria, preservando ajustes feitos no PowerPoint.
- **Ajustar ao slide:** centraliza o grupo, mantém a proporção 16:9 e remove a rotação.
- **Aparência:** escolha Executivo, Essencial ou Escuro e uma das três paletas.
- **Exportar e recuperar:** salve PNG em 1920 × 1080, projeto JSON ou tabela CSV. Use o último rascunho válido quando disponível.
- **Diagnóstico:** confira arquivos e registros próprios sem alterar a instalação ou ler o conteúdo das apresentações.

Atualizar um gráfico redesenha seus elementos a partir do modelo do editor. Ajustes manuais internos são substituídos nessa atualização. Salve sua apresentação e mantenha um JSON de backup.

## Dados e modelos

| Tipo | Cabeçalho da tabela |
| --- | --- |
| Build-up / Comparativo | `categoria;realizado;projetado;baseline` |
| Cascata | `categoria;valor;tipo` |

Use **Colar do Excel** incluindo o cabeçalho ou importe CSV, TSV ou XLSX. Na cascata, os tipos são `base`, `delta` e `total`. Os totais são calculados automaticamente. Até 12 categorias são aceitas por gráfico.

A unidade é um rótulo, sem conversão automática. `320` com unidade `R$ mil` representa R$ 320 mil. As casas decimais e abreviações alteram a apresentação, preservando os valores do modelo.

Abra os modelos JSON pelo botão **Abrir projeto**:

| Modelo | Prévia | Projeto |
| --- | --- | --- |
| Build-up Executivo | [Ver imagem](examples/gallery/buildup-executivo.png) | [Abrir modelo](examples/gallery/buildup-executivo.json) |
| Cascata Essencial | [Ver imagem](examples/gallery/cascata-essencial.png) | [Abrir modelo](examples/gallery/cascata-essencial.json) |
| Comparativo Escuro | [Ver imagem](examples/gallery/comparativo-escuro.png) | [Abrir modelo](examples/gallery/comparativo-escuro.json) |

## Para desenvolver

No Windows, com o .NET Framework instalado, execute na raiz:

```powershell
./build.ps1 -RunTests -RenderGallery
```

Não há dependências NuGet. O comando compila a DLL, o instalador e executa **52 testes do núcleo, 25 verificações da ponte e 21 testes de integração**. Também gera as prévias e o SHA-256 do instalador. Não instala nem registra o suplemento e não abre o PowerPoint. Alguns testes criam janelas nativas invisíveis e rascunhos temporários.

O workflow em **Actions** compila e testa no Windows. Os instaladores produzidos no CI são artefatos experimentais, sem assinatura digital. O ensaio real no Office é separado e está descrito em [Verificação](docs/verification.md).

```text
src/          Cálculos, editor, integração COM e diagnóstico
installer/    Instalação e remoção por usuário
tests/        Testes locais e ensaio opcional no Office
examples/     Tabelas e projetos fictícios
assets/       Ícones próprios
docs/         Guias, referência técnica e notas de versão
release/      Instalador e SHA256SUMS.txt
.github/      Workflow de compilação e modelos de colaboração
```

Leia [CONTRIBUTING.md](CONTRIBUTING.md) para propor melhorias.

## Documentação e limites

[Instalação](docs/install-guide.md) · [Integração com o PowerPoint](docs/native-integration.md) · [Aparência](docs/visual-guide.md) · [Referência completa](docs/reference.md) · [Histórico](CHANGELOG.md)

Esta versão cobre os três tipos de gráfico acima. Ainda não inclui Gantt, Mekko, vínculo automático com Excel, arrastar rótulos ou equivalência funcional completa com ferramentas comerciais.

Código próprio sob [licença MIT](LICENSE). Nenhum instalador, binário, chave de licença ou código do think-cell é distribuído. Componentes da Microsoft são utilizados pela instalação existente no computador e não são redistribuídos.
