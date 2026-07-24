# Coleta RAIS SC

Projeto em Python para extração, tratamento e agregação de microdados da RAIS com foco em Santa Catarina.

O repositório é organizado por finalidade: cada subpasta de `scripts/` reúne um tipo de entrega (estabelecimentos, remuneração, perfil de trabalhadores, indicadores setoriais, mapas ou documentação), e a lógica reaproveitável entre pipelines fica em `src/`.

## Objetivos

Este projeto atende principalmente:

1. Consolidação da RAIS de estabelecimentos para Santa Catarina (e, em um recorte nacional, para todo o Brasil) com agregações por município, mesorregião/UF e divisão CNAE.
2. Extrações de remuneração para RAIS vínculo — incluindo recortes industriais, por divisão CNAE e a versão deflacionada (nominal e real) por CBO e subclasse.
3. Perfil de trabalhadores (sexo, escolaridade) e indicadores/matrizes setoriais por CNAE.
4. Visualizações (heatmap e mapa de predominância setorial) e documentação de apoio ao usuário final.

## Sumário: o que tem em cada pasta

```text
coleta_rais_ftp/
├── data/
│   ├── cnae_dimensao.xlsx        # dimensão CNAE (aba "CNAE-SUB"), usada por quase todos os scripts
│   ├── dict/
│   │   └── dicionario_cbo.xlsx   # dicionário de ocupações CBO (código -> nome)
│   ├── layouts/
│   │   └── RAIS_estabelecimento_layout2018e2019.xls  # layout oficial de campos da RAIS
│   ├── reference/
│   │   ├── municipios_sc_mesorregioes.csv     # referência território: município -> mesorregião (SC)
│   │   ├── ibge_sc_microrregioes.json/.svg    # malha geográfica de microrregiões de SC (API IBGE)
│   │   └── ibge_sc_municipios.json            # malha geográfica de municípios de SC (API IBGE)
│   ├── raw/       # (fora do git) microdados brutos baixados do FTP da RAIS, por ano
│   └── output/    # (fora do git) saídas geradas localmente pelos scripts
├── docs/
│   ├── GUIA_USUARIO.md        # como preparar, executar e interpretar as saídas
│   ├── EXEMPLOS.md            # comandos de exemplo prontos para copiar/colar
│   └── RELEASE_CHECKLIST.md   # checklist antes de publicar uma nova versão
├── examples/      # exemplos sintéticos (não são dados reais) para onboarding
├── scripts/       # pontos de entrada de linha de comando, agrupados por finalidade (ver abaixo)
├── src/           # lógica reutilizável por trás dos pipelines de estabelecimentos
├── CHANGELOG.md
├── CONTRIBUTING.md
├── LICENSE
├── pyproject.toml
└── requirements.txt
```

### `src/` — motor dos pipelines de estabelecimentos

| Arquivo | O que faz |
|---|---|
| `rais_sc_pipeline.py` | Núcleo do pipeline: download via FTP, extração de `.7z`, leitura em chunks, filtro territorial para SC, enriquecimento com dimensão CNAE, exportação em Excel/SQLite. Reaproveitado pelos dois módulos abaixo. |
| `rais_sc_long_panel.py` | Painel "long" por município e divisão CNAE para SC, combinando estabelecimentos e vínculos não abandonados. |
| `rais_brasil_uf_divisao_panel.py` | Painel nacional (todas as regiões/UFs) por divisão CNAE, com estabelecimentos e vínculos. |
| `download_rais.py` | Utilitário de download da RAIS via FTP. |

### `scripts/estabelecimentos/` — consolidação de estabelecimentos

| Script | O que faz |
|---|---|
| `run_rais_sc.py` | Consolida estabelecimentos da RAIS para SC, um ou mais anos. |
| `run_rais_sc_long_panel.py` | Gera o painel long (município x divisão CNAE) para SC. |
| `run_rais_brasil_uf_divisao_panel.py` | Gera o painel nacional por UF e divisão CNAE. |

### `scripts/remuneracao/` — remuneração e deflacionamento salarial

| Script | O que faz |
|---|---|
| `run_remuneracao_media_sc.py` | Remuneração média de dezembro por grande grupo CBO, SC. |
| `run_remuneracao_industrial_cbo_ocupacao_sc.py` | Remuneração média anual nominal industrial por CBO ocupação, SC. |
| `run_remuneracao_media_divisao_sc.py` | Remuneração média anual nominal por divisão CNAE, SC, 2022-2025. |
| `run_remuneracao_3112_cbo_subclasse_sc.py` | Remuneração de dezembro **nominal e real** (deflacionada via INPC/SIDRA-IBGE) por CBO e subclasse CNAE, SC, 2022-2025. **É este o script que deflaciona salário.** |
| `create_validacao_remuneracao_media_ano.py` | Gera planilha de validação a partir do arquivo industrial consolidado. |

### `scripts/perfil_trabalhadores/`

| Script | O que faz |
|---|---|
| `run_perfil_trabalhadores_divisao_sc.py` | Perfil de trabalhadores (sexo, escolaridade) por divisão CNAE e mesorregião, SC, 2022-2025. |

### `scripts/setorial/` — indicadores e matriz setorial

| Script | O que faz |
|---|---|
| `create_matriz_perfil_setorial_rais.py` | Matriz-resumo do perfil setorial (2022 vs 2025) para setores selecionados (alimentos, têxtil, madeira, metalurgia, TI etc.) em SC. |
| `create_indicadores_setoriais_sc_2025.py` | Indicadores setoriais consolidados por CNAE, SC, 2022-2025. |

### `scripts/mapas/` — visualizações

| Script | O que faz |
|---|---|
| `create_heatmap_estoque_empregos_sc_competitiva_2025.py` | Heatmap de estoque de empregos por agrupamento "SC Competitiva" e mesorregião. |
| `create_mapa_predominancia_gr_setor_microrregiao_sc_2025.py` | Mapa SVG de predominância setorial (agropecuária/indústria/serviços) por microrregião de SC, a partir da malha geográfica do IBGE. |

### `scripts/documentacao/` — artefatos de apoio

| Script | O que faz |
|---|---|
| `create_orientacoes_usuario_final_doc.py` | Gera documento `.doc` com orientações de uso para o usuário final. |
| `create_example_workbooks.py` | Gera os workbooks de exemplo em `examples/`. |

## Requisitos

- Python 3.11+ recomendado
- Windows com PowerShell foi o ambiente usado no desenvolvimento
- memória suficiente para leitura de chunks da RAIS

Dependências Python:

- `pandas`
- `openpyxl`
- `XlsxWriter`
- `py7zr`
- `tqdm`
- `xlrd`

## Instalação

### 1. Clonar o repositório

```powershell
git clone <url-do-repositorio>
cd coleta_rais_ftp
```

### 2. Criar ambiente virtual

```powershell
python -m venv venv
.\venv\Scripts\Activate.ps1
```

### 3. Instalar dependências

```powershell
pip install -r requirements.txt
```

Ou, se preferir usar o metadata do projeto:

```powershell
pip install -e .
```

## Arquivos esperados

Alguns arquivos auxiliares são necessários para a execução completa:

- `data/cnae_dimensao.xlsx`
  Planilha de dimensão CNAE com a aba `CNAE-SUB`.
- `data/dict/dicionario_cbo.xlsx`
  Dicionário de ocupações CBO.
- `data/reference/municipios_sc_mesorregioes.csv`
  Referência territorial de municípios e mesorregiões de SC.

Observação:
arquivos brutos e saídas geradas ficam fora do versionamento por padrão e devem ser mantidos localmente.

## Como executar

Todos os comandos abaixo usam o Python do ambiente virtual e assumem execução a partir da raiz do projeto.

### 1. Pipeline de estabelecimentos

```powershell
.\venv\Scripts\python.exe .\scripts\estabelecimentos\run_rais_sc.py --years 2022 2023 2024
```

Ou apenas para um ano:

```powershell
.\venv\Scripts\python.exe .\scripts\estabelecimentos\run_rais_sc.py --year 2024 --cnae-dimension-path .\data\cnae_dimensao.xlsx
```

Saídas principais:

- `data/output/rais_estabelecimentos_sc_2024.xlsx`
- `data/output/rais_estabelecimentos_sc_2022_2024.xlsx`
- `data/output/rais_estabelecimentos_sc_2022_2024.sqlite`

Painéis relacionados:

```powershell
.\venv\Scripts\python.exe .\scripts\estabelecimentos\run_rais_sc_long_panel.py
.\venv\Scripts\python.exe .\scripts\estabelecimentos\run_rais_brasil_uf_divisao_panel.py
```

### 2. Remuneração média de dezembro por grande grupo CBO

```powershell
.\venv\Scripts\python.exe .\scripts\remuneracao\run_remuneracao_media_sc.py
```

Também suporta filtro adicional de vínculo ativo em 31/12:

```powershell
.\venv\Scripts\python.exe .\scripts\remuneracao\run_remuneracao_media_sc.py --active-3112-only
```

### 3. Remuneração média anual nominal industrial por CBO ocupação

```powershell
.\venv\Scripts\python.exe .\scripts\remuneracao\run_remuneracao_industrial_cbo_ocupacao_sc.py
```

Saídas principais:

- `data/output/remuneracao_media_ano_industrial_sc_cbo_ocupacao_2024.xlsx`
- `data/output/remuneracao_media_ano_industrial_abdon_batista_cbo_ocupacao_2024.xlsx`

### 4. Remuneração deflacionada (nominal e real) por CBO e subclasse CNAE

```powershell
.\venv\Scripts\python.exe .\scripts\remuneracao\run_remuneracao_3112_cbo_subclasse_sc.py
```

Saída principal:

- `data/output/remuneracao_3112_real_sc_cbo_subclasse_2022_2025.xlsx`

### 5. Arquivo de validação

```powershell
.\venv\Scripts\python.exe .\scripts\remuneracao\create_validacao_remuneracao_media_ano.py
```

Saída:

- `data/output/validacao.xlsx`

### 6. Perfil de trabalhadores e indicadores setoriais

```powershell
.\venv\Scripts\python.exe .\scripts\perfil_trabalhadores\run_perfil_trabalhadores_divisao_sc.py
.\venv\Scripts\python.exe .\scripts\setorial\create_matriz_perfil_setorial_rais.py
.\venv\Scripts\python.exe .\scripts\setorial\create_indicadores_setoriais_sc_2025.py
```

### 7. Mapas

```powershell
.\venv\Scripts\python.exe .\scripts\mapas\create_heatmap_estoque_empregos_sc_competitiva_2025.py
.\venv\Scripts\python.exe .\scripts\mapas\create_mapa_predominancia_gr_setor_microrregiao_sc_2025.py
```

### 8. Documento orientativo ao usuário final

```powershell
.\venv\Scripts\python.exe .\scripts\documentacao\create_orientacoes_usuario_final_doc.py
```

Saída:

- `data/output/orientacoes_usuario_final.doc`

### 9. Geração dos exemplos em Excel

```powershell
.\venv\Scripts\python.exe .\scripts\documentacao\create_example_workbooks.py
```

Saídas:

- `examples/sample_remuneracao_media_ano_industrial_sc_cbo_ocupacao_2024.xlsx`
- `examples/sample_validacao.xlsx`

## Metodologia resumida

### RAIS estabelecimentos

- filtro territorial para SC;
- separação de RAIS negativa;
- agregações por município, mesorregião, divisão CNAE e agrupamentos auxiliares;
- exportação em Excel e SQLite.

### RAIS vínculo industrial

Na extração industrial mais recente, os filtros aplicados foram:

- somente municípios de Santa Catarina;
- somente CNAEs divisão `05` a `43`;
- `Ind Vínculo Abandonado = 0`;
- `Ind Vínculo Ativo 31/12 = 1`;
- `Vl Rem Média Nom > 0`.

Níveis de desagregação:

- território: município e mesorregião;
- CBO: `CBO 2002 Ocupação` de 6 dígitos;
- CNAE: subclasse e divisão.

### Deflacionamento (remuneração real)

`run_remuneracao_3112_cbo_subclasse_sc.py` obtém o índice INPC via API do SIDRA/IBGE (tabela 1736, variável 2289) e aplica a correção sobre a remuneração nominal de dezembro, produzindo colunas nominal e real lado a lado por CBO e subclasse CNAE.

## Interface de linha de comando

Os scripts foram organizados para usar uma CLI previsível sempre que possível. Em termos práticos, isso significa:

- aceitar parâmetros por terminal;
- permitir sobrescrever caminhos de entrada e saída;
- usar nomes de argumentos fáceis de entender, como `--input-path`, `--output-path`, `--year` e `--years`.

Todo script aceita `--help` para listar seus argumentos.

Exemplo:

```powershell
.\venv\Scripts\python.exe .\scripts\remuneracao\create_validacao_remuneracao_media_ano.py --input-path .\data\output\remuneracao_media_ano_industrial_sc_cbo_ocupacao_2024.xlsx --output-path .\data\output\validacao_custom.xlsx
```

## Limitações e observações

- Os microdados da RAIS podem mudar de localização, nomenclatura ou encoding entre anos.
- O projeto hoje está orientado ao ambiente Windows e PowerShell.
- Parte dos arquivos auxiliares depende de estrutura local em `data/`.
- Os dados brutos não devem ser versionados no GitHub por volume e licenciamento da fonte.
- `data/output/` é regenerável a partir dos scripts: em caso de dúvida sobre um arquivo de saída, prefira reexecutar o script correspondente a manter cópias manuais renomeadas.

## Publicação no GitHub

Recomendação de publicação:

- versionar apenas código, documentação e arquivos auxiliares permitidos;
- não subir `data/raw/`, `data/output/` nem `venv/`;
- incluir no repositório apenas exemplos mínimos ou arquivos sintéticos, se necessário;
- descrever claramente a origem dos dados e o passo a passo de execução.

## Convenções de nomenclatura

Para facilitar automação, terminal e uso em diferentes sistemas, a documentação e as saídas do projeto seguem:

- nomes de arquivos e scripts em ASCII (sem acento ou espaço);
- `snake_case` em todos os nomes;
- prefixo `run_` para scripts que executam um pipeline principal e `create_` para scripts que geram um artefato derivado (validação, documento, mapa, exemplo);
- argumentos de linha de comando explícitos.

Exemplos:

- `validacao.xlsx`
- `orientacoes_usuario_final.doc`
- `remuneracao_media_ano_industrial_sc_cbo_ocupacao_2024.xlsx`

## Documentação adicional

Consulte o guia do usuário em [docs/GUIA_USUARIO.md](docs/GUIA_USUARIO.md).
Consulte exemplos prontos em [docs/EXEMPLOS.md](docs/EXEMPLOS.md).
Consulte o checklist de release em [docs/RELEASE_CHECKLIST.md](docs/RELEASE_CHECKLIST.md).
Consulte exemplos sintéticos em [examples/README.md](examples/README.md).
Consulte o histórico de mudanças em [CHANGELOG.md](CHANGELOG.md).

Para contribuições e manutenção do repositório, consulte [CONTRIBUTING.md](CONTRIBUTING.md).

## Licença

Este projeto está licenciado sob a licença MIT. Veja [LICENSE](LICENSE).
