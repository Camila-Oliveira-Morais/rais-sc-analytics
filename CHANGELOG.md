# Changelog

Todas as mudanças relevantes deste projeto passam a ser registradas aqui.

O formato segue, de maneira simplificada, a ideia de "Keep a Changelog".

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
