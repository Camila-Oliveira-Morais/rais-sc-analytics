# Coleta RAIS SC

Projeto em Python para extração, tratamento e agregação de microdados da RAIS com foco em Santa Catarina.

O repositório foi organizado para permitir reprodução local, documentação de uso e publicação no GitHub, preservando a separação entre:

- código-fonte;
- scripts operacionais;
- arquivos auxiliares de dimensão e dicionários;
- dados brutos e saídas geradas localmente.

## Objetivos

Este projeto atende principalmente dois tipos de entrega:

1. Consolidação da RAIS de estabelecimentos para Santa Catarina com agregações analíticas por município, mesorregião, divisão CNAE e agrupamentos auxiliares.
2. Extrações específicas de remuneração para RAIS vínculo, incluindo recortes industriais e detalhamento por CBO.

## Funcionalidades principais

- download da RAIS a partir do FTP oficial;
- extração de arquivos `.7z`;
- leitura em chunks para lidar com arquivos grandes;
- filtro territorial para Santa Catarina;
- enriquecimento com dimensão de CNAE;
- enriquecimento com nomes de ocupação via dicionário de CBO;
- geração de saídas em Excel, CSV e SQLite;
- geração de arquivos auxiliares de validação e documentação operacional.

## Estrutura do projeto

```text
coleta_rais_ftp/
├── data/
│   ├── cnae_dimensao.xlsx
│   ├── dict/
│   │   └── dicionario_cbo.xlsx
│   ├── layouts/
│   ├── output/
│   ├── raw/
│   └── reference/
├── docs/
├── examples/
├── scripts/
├── src/
├── .gitignore
├── CHANGELOG.md
├── LICENSE
├── README.md
└── requirements.txt
```

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

### 1. Pipeline de estabelecimentos

Executa a consolidação de estabelecimentos para um ou mais anos:

```powershell
.\venv\Scripts\python.exe .\scripts\run_rais_sc.py --years 2022 2023 2024
```

Ou apenas para um ano:

```powershell
.\venv\Scripts\python.exe .\scripts\run_rais_sc.py --year 2024 --cnae-dimension-path .\data\cnae_dimensao.xlsx
```

Saídas principais:

- `data/output/rais_estabelecimentos_sc_2024.xlsx`
- `data/output/rais_estabelecimentos_sc_2022_2024.xlsx`
- `data/output/rais_estabelecimentos_sc_2022_2024.sqlite`

### 2. Remuneração média de dezembro por grande grupo CBO

```powershell
.\venv\Scripts\python.exe .\scripts\run_remuneracao_media_sc.py
```

Também suporta filtro adicional de vínculo ativo em 31/12:

```powershell
.\venv\Scripts\python.exe .\scripts\run_remuneracao_media_sc.py --active-3112-only
```

### 3. Remuneração média anual nominal industrial por CBO ocupação

```powershell
.\venv\Scripts\python.exe .\scripts\run_remuneracao_industrial_cbo_ocupacao_sc.py
```

Saídas principais:

- `data/output/remuneracao_media_ano_industrial_sc_cbo_ocupacao_2024.xlsx`
- `data/output/remuneracao_media_ano_industrial_abdon_batista_cbo_ocupacao_2024.xlsx`

### 4. Arquivo de validação

```powershell
.\venv\Scripts\python.exe .\scripts\create_validacao_remuneracao_media_ano.py
```

Saída:

- `data/output/validacao.xlsx`

### 5. Documento orientativo ao usuário final

```powershell
.\venv\Scripts\python.exe .\scripts\create_orientacoes_usuario_final_doc.py
```

Saída:

- `data/output/orientacoes_usuario_final.doc`

### 6. Geração dos exemplos em Excel

```powershell
.\venv\Scripts\python.exe .\scripts\create_example_workbooks.py
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

## Scripts disponíveis

- `scripts/run_rais_sc.py`
- `scripts/run_remuneracao_media_sc.py`
- `scripts/run_remuneracao_industrial_cbo_ocupacao_sc.py`
- `scripts/create_validacao_remuneracao_media_ano.py`
- `scripts/create_orientacoes_usuario_final_doc.py`
- `scripts/create_example_workbooks.py`

## Interface de linha de comando

Os scripts foram organizados para usar uma CLI previsível sempre que possível. Em termos práticos, isso significa:

- aceitar parâmetros por terminal;
- permitir sobrescrever caminhos de entrada e saída;
- usar nomes de argumentos fáceis de entender, como `--input-path`, `--output-path`, `--year` e `--years`.

Exemplos:

```powershell
.\venv\Scripts\python.exe .\scripts\create_validacao_remuneracao_media_ano.py --input-path .\data\output\remuneracao_media_ano_industrial_sc_cbo_ocupacao_2024.xlsx --output-path .\data\output\validacao_custom.xlsx
```

```powershell
.\venv\Scripts\python.exe .\scripts\create_orientacoes_usuario_final_doc.py --output-path .\data\output\guia_usuario.doc
```

## Limitações e observações

- Os microdados da RAIS podem mudar de localização, nomenclatura ou encoding entre anos.
- O projeto hoje está orientado ao ambiente Windows e PowerShell.
- Parte dos arquivos auxiliares depende de estrutura local em `data/`.
- Os dados brutos não devem ser versionados no GitHub por volume e licenciamento da fonte.

## Publicação no GitHub

Recomendação de publicação:

- versionar apenas código, documentação e arquivos auxiliares permitidos;
- não subir `data/raw/`, `data/output/` nem `venv/`;
- incluir no repositório apenas exemplos mínimos ou arquivos sintéticos, se necessário;
- descrever claramente a origem dos dados e o passo a passo de execução.

## Convenções de nomenclatura

Para facilitar automação, terminal e uso em diferentes sistemas, a documentação e as novas saídas padrão do projeto passam a preferir:

- nomes de arquivos em ASCII;
- `snake_case` ou nomes sem espaços;
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
