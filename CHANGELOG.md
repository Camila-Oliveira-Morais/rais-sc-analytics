# Changelog

Todas as mudanças relevantes deste projeto passam a ser registradas aqui.

O formato segue, de maneira simplificada, a ideia de "Keep a Changelog".

## [Não versionado] - 2026-07-24

### Alterado

- Scripts reorganizados em subpastas de `scripts/` por finalidade: `estabelecimentos/`, `remuneracao/`, `perfil_trabalhadores/`, `setorial/`, `mapas/` e `documentacao/`.
- Scripts `gerar_matriz_perfil_setorial_rais.py` e `gerar_indicadores_setoriais_sc_2025.py` renomeados para o padrão `create_*` (agora `create_matriz_perfil_setorial_rais.py` e `create_indicadores_setoriais_sc_2025.py`).
- README.md reescrito como sumário central do projeto, documentando todas as pastas e os 15 scripts existentes.
- `docs/GUIA_USUARIO.md`, `docs/EXEMPLOS.md` e `docs/RELEASE_CHECKLIST.md` atualizados com os novos caminhos de scripts.

### Removido

- Arquivos temporários e obsoletos em `data/output/` (`_tmp_*`, saídas de nomenclatura antiga já substituídas por versões mais recentes) e uma pasta de perfil do navegador Edge gerada por engano (`_tmp_edge_profile/`).

### Observações

- Os arquivos remanescentes em `data/output/` com nomes acentuados (`remuneração média SC.xlsx`, `validação.xlsx`, `orientações_usuario_final.doc`) foram renomeados para o padrão ASCII já usado pelos scripts (`remuneracao_media_sc.xlsx`, `validacao.xlsx`, `orientacoes_usuario_final.doc`), sem alteração de conteúdo.

## [0.1.0] - 2026-03-30

### Adicionado

- README reestruturado para publicação no GitHub.
- Guia do usuário em `docs/GUIA_USUARIO.md`.
- Exemplos de execução em `docs/EXEMPLOS.md`.
- Checklist de release em `docs/RELEASE_CHECKLIST.md`.
- Guia de contribuição em `CONTRIBUTING.md`.
- Licença MIT em `LICENSE`.
- Metadata do projeto em `pyproject.toml`.
- Script para geração de validação.
- Script para geração de documento orientativo ao usuário final.
- Enriquecimento por nome de ocupação a partir de `data/dict/dicionario_cbo.xlsx`.
- Extração industrial por `CBO 2002 Ocupação` com detalhamento por subclasse CNAE.
- Pasta `examples/` com arquivos sintéticos de referência.

### Alterado

- Padronização da documentação principal do repositório.
- Padronização dos scripts para uso por linha de comando com argumentos mais previsíveis.
- Padronização de nomes de saídas novas em formato ASCII quando possível.
- Robustez da leitura da aba do dicionário CBO.

### Removido

- Arquivo redundante `src/_init_.py`.

### Observações

- Dados brutos e saídas analíticas continuam fora do versionamento.
- O projeto permanece orientado a execução local, principalmente em Windows + PowerShell.
