# Publicar no GitHub

1. Crie o repositório público **mello-charts-desktop** na conta `andremellotech`. Deixe desmarcadas as opções de criar README, licença e .gitignore: o pacote já contém esses arquivos.
2. Extraia o ZIP e abra a pasta do projeto. Pelo GitHub Web, use **Add file → Upload files** e arraste as pastas e os arquivos da raiz juntos. Arraste as pastas inteiras para preservar sua estrutura; não selecione todos os arquivos internos de uma vez.
3. Antes de confirmar, confira os caminhos: `src/ChartCore.cs`, `docs/install-guide.md` e `examples/gallery/buildup-executivo.png`. O `README.md` deve ficar na raiz. Não envie `build`, `.git` ou o ZIP dentro do repositório.
4. O workflow precisa estar exatamente em `.github/workflows/ci.yml`. Se a pasta oculta `.github` não aparecer no explorador ou não for enviada, crie o arquivo pelo GitHub em **Add file → Create new file**, usando esse caminho e o conteúdo fornecido no pacote.
5. Clique em **Commit changes**. Confira em **Actions** a execução de compilação e testes. O CI gera um instalador experimental como artefato, mas não publica Releases automaticamente.
6. Em **Releases → Draft a new release**, crie a tag `v0.6.0`, use o título **Mello Charts Desktop 0.6.0 — experimental** e copie a descrição de [v0.6.0.md](releases/v0.6.0.md).
7. Anexe **os dois arquivos da mesma compilação**: `release/MelloCharts-Setup.exe` e `release/SHA256SUMS.txt`. Marque **Set as a pre-release** e publique. Os arquivos da pasta `release` também podem ser baixados diretamente pelo botão **Download raw file** do GitHub.

Não é necessário GitHub Pages para esta versão instalada. O usuário baixa o EXE e instala no seu Windows. O pacote não publica nem modifica sua conta automaticamente. A seção [Verificação](verification.md) distingue os testes locais da validação que ainda falta no PowerPoint real.

Descrição sugerida: **Suplemento gratuito para PowerPoint Windows: 11 modelos de gráficos executivos, galeria visual, formas editáveis e importação de Excel/CSV. Versão experimental.**
