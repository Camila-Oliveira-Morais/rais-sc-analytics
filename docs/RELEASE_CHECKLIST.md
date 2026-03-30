# Checklist de Release

## Antes de publicar

- Revisar se `README.md` está atualizado.
- Revisar se `docs/GUIA_USUARIO.md` continua coerente com os scripts.
- Revisar se `docs/EXEMPLOS.md` reflete os comandos reais.
- Revisar `requirements.txt` e `pyproject.toml`.
- Confirmar que a licença MIT está presente.

## Limpeza do repositório

- Garantir que `data/raw/` não será versionado.
- Garantir que `data/output/` não será versionado.
- Garantir que `venv/` não será versionado.
- Garantir que caches e arquivos temporários estão no `.gitignore`.

## Verificações funcionais mínimas

- Rodar `python .\scripts\run_rais_sc.py --help`
- Rodar `python .\scripts\run_remuneracao_media_sc.py --help`
- Rodar `python .\scripts\run_remuneracao_industrial_cbo_ocupacao_sc.py --help`
- Rodar `python .\scripts\create_validacao_remuneracao_media_ano.py --help`
- Rodar `python .\scripts\create_orientacoes_usuario_final_doc.py --help`

## Conferência de dados auxiliares

- Confirmar presença de `data/cnae_dimensao.xlsx`
- Confirmar presença de `data/dict/dicionario_cbo.xlsx`
- Confirmar presença de `data/reference/municipios_sc_mesorregioes.csv`

## Publicação

- Criar branch ou tag de release
- Adicionar notas de versão
- Publicar com instruções claras de execução local
