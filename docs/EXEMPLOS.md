# Exemplos de Execução

## 1. Instalação

```powershell
python -m venv venv
.\venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

## 2. Pipeline de estabelecimentos

### Um único ano

```powershell
.\venv\Scripts\python.exe .\scripts\run_rais_sc.py --year 2024 --cnae-dimension-path .\data\cnae_dimensao.xlsx
```

### Múltiplos anos

```powershell
.\venv\Scripts\python.exe .\scripts\run_rais_sc.py --years 2022 2023 2024
```

## 3. Remuneração média de dezembro

### Saída padrão

```powershell
.\venv\Scripts\python.exe .\scripts\run_remuneracao_media_sc.py
```

### Com filtro de vínculo ativo em 31/12

```powershell
.\venv\Scripts\python.exe .\scripts\run_remuneracao_media_sc.py --active-3112-only
```

### Com caminhos customizados

```powershell
.\venv\Scripts\python.exe .\scripts\run_remuneracao_media_sc.py --input-path .\data\raw\2024_vinc_sul\extracted\RAIS_VINC_PUB_SUL.COMT --cnae-dimension-path .\data\cnae_dimensao.xlsx --municipality-reference-path .\data\reference\municipios_sc_mesorregioes.csv --output-path .\data\output\remuneracao_media_sc_custom.xlsx
```

## 4. Remuneração média anual industrial por CBO ocupação

### Saída padrão

```powershell
.\venv\Scripts\python.exe .\scripts\run_remuneracao_industrial_cbo_ocupacao_sc.py
```

### Com diretório de saída customizado

```powershell
.\venv\Scripts\python.exe .\scripts\run_remuneracao_industrial_cbo_ocupacao_sc.py --output-dir .\data\output
```

### Com caminhos totalmente explícitos

```powershell
.\venv\Scripts\python.exe .\scripts\run_remuneracao_industrial_cbo_ocupacao_sc.py --input-path .\data\raw\2024_vinc_sul\extracted\RAIS_VINC_PUB_SUL.COMT --cnae-dimension-path .\data\cnae_dimensao.xlsx --municipality-reference-path .\data\reference\municipios_sc_mesorregioes.csv --cbo-dictionary-path .\data\dict\dicionario_cbo.xlsx --main-output-path .\data\output\industrial_sc.xlsx --abdon-output-path .\data\output\industrial_abdon.xlsx --chunk-size 300000
```

## 5. Geração de validação

```powershell
.\venv\Scripts\python.exe .\scripts\create_validacao_remuneracao_media_ano.py --input-path .\data\output\remuneracao_media_ano_industrial_sc_cbo_ocupacao_2024.xlsx --output-path .\data\output\validacao.xlsx
```

## 6. Documento ao usuário final

```powershell
.\venv\Scripts\python.exe .\scripts\create_orientacoes_usuario_final_doc.py --output-path .\data\output\orientacoes_usuario_final.doc
```
