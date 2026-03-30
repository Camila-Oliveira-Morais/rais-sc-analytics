# Contribuindo

## Objetivo

Este projeto foi estruturado para facilitar extrações reproduzíveis da RAIS para Santa Catarina. Se você for contribuir, a meta é preservar:

- rastreabilidade dos dados;
- clareza metodológica;
- execução local previsível;
- documentação alinhada ao código.

## Ambiente recomendado

- Windows com PowerShell
- Python 3.11+
- ambiente virtual em `venv/`

## Setup rápido

```powershell
python -m venv venv
.\venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

## Princípios para mudanças

- Não versionar dados brutos nem saídas analíticas.
- Preferir nomes de arquivos e parâmetros estáveis.
- Documentar qualquer nova regra metodológica no `README` e no guia do usuário.
- Quando alterar filtros, atualizar também as explicações dos scripts e das saídas.

## Padrões para scripts

Os scripts devem, sempre que possível:

- aceitar parâmetros por linha de comando;
- permitir sobrescrever caminhos de entrada e saída;
- gerar mensagens claras de erro;
- manter leitura em chunks para arquivos grandes;
- usar codificação explicitamente quando ler arquivos externos.

## O que revisar antes de abrir um PR

- O script executa do zero com os arquivos esperados?
- O `README` continua válido?
- O guia do usuário precisa ser atualizado?
- O `.gitignore` continua cobrindo dados locais e saídas?
- A mudança altera nomes de arquivos produzidos? Se sim, isso foi documentado?

## Escopo de versionamento

Devem ficar no GitHub:

- código-fonte;
- scripts;
- documentação;
- licença;
- arquivos auxiliares permitidos.

Não devem ficar no GitHub:

- `data/raw/`
- `data/output/`
- ambiente virtual
- caches e arquivos temporários
