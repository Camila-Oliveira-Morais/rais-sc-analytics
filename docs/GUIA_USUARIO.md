# Guia do Usuário

## Objetivo

Este guia explica como preparar, executar e interpretar os arquivos gerados pelo projeto `coleta_rais_ftp`.

## Quando usar este projeto

Use este projeto quando precisar:

- consolidar dados da RAIS para Santa Catarina;
- produzir tabelas analíticas por município e CNAE;
- extrair remuneração média por recortes específicos;
- enriquecer resultados com nomes de ocupação CBO;
- gerar planilhas prontas para validação e consumo por usuário final.

## Fluxo operacional recomendado

### 1. Preparar ambiente

```powershell
python -m venv venv
.\venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

### 2. Conferir arquivos auxiliares

Verifique se estes arquivos estão disponíveis:

- `data/cnae_dimensao.xlsx`
- `data/dict/dicionario_cbo.xlsx`
- `data/reference/municipios_sc_mesorregioes.csv`

### 3. Executar o tipo de rotina desejada

#### Consolidação de estabelecimentos

```powershell
.\venv\Scripts\python.exe .\scripts\run_rais_sc.py --years 2022 2023 2024
```

#### Remuneração industrial por CBO ocupação

```powershell
.\venv\Scripts\python.exe .\scripts\run_remuneracao_industrial_cbo_ocupacao_sc.py
```

#### Geração de validação

```powershell
.\venv\Scripts\python.exe .\scripts\create_validacao_remuneracao_media_ano.py
```

## Arquivos entregues ao usuário final

### Arquivo geral industrial

`remuneracao_media_ano_industrial_sc_cbo_ocupacao_2024.xlsx`

Contém:

- ano de referência;
- município;
- mesorregião;
- código e nome da ocupação CBO;
- código e nome da divisão CNAE;
- código e nome da subclasse CNAE;
- quantidade de vínculos;
- remuneração média anual nominal.

### Arquivo de teste municipal

`remuneracao_media_ano_industrial_abdon_batista_cbo_ocupacao_2024.xlsx`

Serve para:

- conferência metodológica;
- validação pontual;
- apoio a homologação.

### Arquivo de validação

`validacao.xlsx`

Resume indicadores globais do arquivo principal:

- municípios únicos;
- mesorregiões únicas;
- total de vínculos;
- média salarial geral.

## Regras metodológicas usadas na extração industrial

### Recorte geográfico

- somente Santa Catarina;
- identificação territorial por código do município.

### Recorte setorial

- somente CNAEs divisão `05` a `43`.

### Regras de filtro

- `Ind Vínculo Abandonado = 0`
- `Ind Vínculo Ativo 31/12 = 1`
- `Vl Rem Média Nom > 0`

### Nível de desagregação

- CBO: ocupação de 6 dígitos;
- CNAE: subclasse e divisão;
- território: município e mesorregião.

## Como interpretar as colunas principais

### `qtd_vinculos`

Representa quantos vínculos válidos compõem aquele agrupamento.

### `remuneracao_media_nom_ano`

Representa a remuneração média anual nominal para o agrupamento após aplicação dos filtros.

### `cbo_ocupacao_nome`

É o nome da ocupação, preenchido a partir do dicionário CBO por chave de 6 dígitos.

## Boas práticas de uso

- Sempre interpretar a média junto com `qtd_vinculos`.
- Evitar conclusões fortes em células com muito poucos vínculos.
- Validar resultados agregados com o arquivo de validação.
- Em caso de conferência externa, comparar se os mesmos filtros foram aplicados.

## Problemas comuns

### Arquivo não encontrado

Verifique se os arquivos em `data/` existem localmente.

### Dicionário CBO não preenche nomes

Verifique:

- se `data/dict/dicionario_cbo.xlsx` existe;
- se a aba da planilha contém a lista de ocupações;
- se os códigos do dicionário estão em formato de 6 dígitos.

### Diferença em relação a outras extrações

Diferenças normalmente vêm de:

- filtro de vínculo ativo em 31/12;
- uso de remuneração média anual versus remuneração de dezembro;
- nível de agregação do CBO;
- recorte setorial por divisão ou subclasse CNAE.

## Recomendação para uso em GitHub

- manter apenas código e documentação no repositório;
- deixar dados brutos e saídas fora do versionamento;
- documentar claramente qualquer dependência local;
- registrar exemplos de execução no `README`.

## Exemplos sintéticos

Se você ainda não possui os microdados reais, use a pasta `examples/` apenas para entender:

- a estrutura das colunas;
- o formato das saídas;
- o tipo de validação produzida.

Esses arquivos não substituem os dados reais.
