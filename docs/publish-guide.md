# Publicar uma atualização

1. Execute `./build.ps1 -RunTests -RenderGallery`.
2. Confira os pacotes `release/MelloCharts-Setup.exe`, `release/MelloCharts-Portable.zip` e `release/SHA256SUMS.txt` da mesma compilação.
3. Publique as fontes, exemplos e pacotes no repositório. A pasta `build/` deve permanecer fora do Git.
4. Confira o workflow Windows em Actions; ele recompila, testa e disponibiliza os pacotes como artefatos. Não executa Office nos runners hospedados.
5. Para uma Release, use a tag `v0.8.0`, as notas em [v0.8.0.md](releases/v0.8.0.md), os três arquivos da mesma compilação e marque a versão experimental como pré-lançamento. A CI não cria Releases automaticamente.

Não é necessário GitHub Pages. Os downloads atuais são ligados diretamente à pasta `release/` no README. [Validação](verification.md) registra os testes efetivamente executados e os limites.
