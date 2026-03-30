from __future__ import annotations

import argparse
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
OUTPUT_FILE = PROJECT_ROOT / "data" / "output" / "orientacoes_usuario_final.doc"


HTML_CONTENT = """\
<html>
<head>
<meta charset="utf-8">
<title>Orientações ao Usuário Final</title>
<style>
body { font-family: Calibri, Arial, sans-serif; font-size: 11pt; line-height: 1.4; margin: 32px; }
h1, h2 { color: #1f4e79; }
table { border-collapse: collapse; width: 100%; margin-top: 12px; margin-bottom: 16px; }
th, td { border: 1px solid #999; padding: 8px; vertical-align: top; }
th { background: #ddebf7; text-align: left; }
p { margin: 8px 0; }
ul { margin: 6px 0 12px 22px; }
</style>
</head>
<body>
<h1>Orientações ao Usuário Final</h1>

<p>Este documento orienta a leitura e o uso dos arquivos analíticos de remuneração média anual da RAIS 2024 para Santa Catarina, com foco nos CNAEs industriais e desagregação por ocupação CBO.</p>

<h2>1. Arquivos gerados</h2>
<ul>
<li><b>Arquivo geral SC:</b> <i>remuneracao_media_ano_industrial_sc_cbo_ocupacao_2024.xlsx</i></li>
<li><b>Arquivo de teste:</b> <i>remuneracao_media_ano_industrial_abdon_batista_cbo_ocupacao_2024.xlsx</i></li>
<li><b>Arquivo de validação:</b> <i>validação.xlsx</i></li>
</ul>

<h2>2. Base de origem</h2>
<p>Os dados foram produzidos a partir do microdado de vínculos da RAIS 2024 para a região Sul, restringindo a análise aos registros de Santa Catarina.</p>

<h2>3. Filtros aplicados</h2>
<table>
<tr><th>Filtro</th><th>Regra utilizada</th></tr>
<tr><td>Unidade geográfica</td><td>Somente vínculos de Santa Catarina, identificados pelo código do município iniciado por 42.</td></tr>
<tr><td>RAIS industrial</td><td>Somente CNAEs divisão de 05 até 43.</td></tr>
<tr><td>Abandono de vínculo</td><td>Somente registros com <b>Ind Vínculo Abandonado = NÃO</b>.</td></tr>
<tr><td>Vínculo ativo</td><td>Somente registros com <b>Ind Vínculo Ativo 31/12 = SIM</b>.</td></tr>
<tr><td>Remuneração</td><td>Somente registros com <b>Vl Rem Média Nom &gt; 0</b>.</td></tr>
</table>

<h2>4. Variável de remuneração utilizada</h2>
<p>A métrica usada foi <b>Vl Rem Média Nom</b>, que representa a remuneração média do ano em valores nominais.</p>
<p>O valor final apresentado nos arquivos corresponde à <b>média da remuneração média nominal anual</b> dentro de cada agrupamento, ponderada pela quantidade de vínculos observados em cada combinação.</p>

<h2>5. Nível de desagregação</h2>
<table>
<tr><th>Dimensão</th><th>Nível utilizado</th></tr>
<tr><td>CBO</td><td><b>CBO 2002 Ocupação</b>, no maior nível disponível no microdado, com código de 6 dígitos.</td></tr>
<tr><td>Nome do CBO</td><td>Obtido do dicionário <i>dicionario_cbo.xlsx</i>, vinculado pelo código de 6 dígitos da ocupação.</td></tr>
<tr><td>CNAE</td><td><b>CNAE Subclasse</b> como nível analítico mais detalhado, com apoio da <b>CNAE Divisão</b> para o recorte industrial.</td></tr>
<tr><td>Território</td><td>Município e mesorregião de Santa Catarina.</td></tr>
</table>

<h2>6. Colunas principais dos arquivos</h2>
<table>
<tr><th>Coluna</th><th>Descrição</th></tr>
<tr><td>ano_referencia</td><td>Ano de referência da RAIS utilizado na extração.</td></tr>
<tr><td>municipio_codigo</td><td>Código do município.</td></tr>
<tr><td>municipio_nome</td><td>Nome do município.</td></tr>
<tr><td>mesorregiao_nome</td><td>Nome da mesorregião do município.</td></tr>
<tr><td>cbo_ocupacao_codigo</td><td>Código CBO da ocupação, com 6 dígitos.</td></tr>
<tr><td>cbo_ocupacao_nome</td><td>Nome da ocupação, preenchido a partir do dicionário de CBO.</td></tr>
<tr><td>cnae_divisao_codigo</td><td>Código da divisão CNAE.</td></tr>
<tr><td>cnae_divisao_nome</td><td>Nome da divisão CNAE.</td></tr>
<tr><td>cnae_subclasse_codigo</td><td>Código da subclasse CNAE.</td></tr>
<tr><td>cnae_subclasse_nome</td><td>Nome da subclasse CNAE.</td></tr>
<tr><td>qtd_vinculos</td><td>Quantidade de vínculos válidos no agrupamento.</td></tr>
<tr><td>remuneracao_media_nom_ano</td><td>Remuneração média anual nominal do agrupamento.</td></tr>
</table>

<h2>7. Boas práticas de leitura</h2>
<ul>
<li>Ao comparar remuneração entre grupos, observar sempre a coluna <b>qtd_vinculos</b>, pois médias baseadas em poucos vínculos podem ser mais instáveis.</li>
<li>O recorte industrial foi definido pela divisão CNAE, e não por interpretação manual do nome da atividade.</li>
<li>O arquivo de Abdon Batista foi gerado para conferência e teste de consistência metodológica.</li>
<li>O arquivo de validação resume indicadores globais do arquivo geral e pode ser usado como checagem rápida de consistência.</li>
</ul>

<h2>8. Observações metodológicas</h2>
<ul>
<li>Os resultados consideram apenas vínculos com remuneração positiva.</li>
<li>Os nomes das ocupações dependem do dicionário informado pelo usuário e podem variar conforme a versão do dicionário utilizada.</li>
<li>Quando necessário, recomenda-se usar o arquivo geral para análises completas e o arquivo municipal para verificações pontuais.</li>
</ul>

<h2>9. Resumo executivo</h2>
<p>Em síntese, os arquivos entregues apresentam a remuneração média anual nominal da RAIS 2024 para Santa Catarina, no recorte industrial de CNAE divisão 05 a 43, com detalhamento máximo por ocupação CBO, município, mesorregião, divisão e subclasse CNAE.</p>

</body>
</html>
"""


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Gera um documento .doc com orientacoes de uso para o usuario final."
    )
    parser.add_argument(
        "--output-path",
        default=str(OUTPUT_FILE),
        help="Caminho de saida do documento .doc.",
    )
    return parser


def main() -> None:
    args = build_parser().parse_args()
    output_path = Path(args.output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(HTML_CONTENT, encoding="utf-8")


if __name__ == "__main__":
    main()
