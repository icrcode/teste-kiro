"""
Renderizador de Arquiteturas — Previsão de Demanda de Vagas por Curso
Lê um arquivo YAML de arquitetura e gera um diagrama visual (PNG/SVG).
"""

import yaml
import json
import os
import argparse
from datetime import datetime


# ---------------------------------------------------------------------------
# Mapeamento de ícones por serviço AWS
# ---------------------------------------------------------------------------
ICONE_AWS = {
    "S3":           "icones/s3.png",
    "Glue":         "icones/glue.png",
    "SageMaker":    "icones/sagemaker.png",
    "Lambda":       "icones/lambda.png",
    "RDS":          "icones/rds.png",
    "Aurora":       "icones/aurora.png",
    "EMR":          "icones/emr.png",
    "CloudWatch":   "icones/cloudwatch.png",
    "EventBridge":  "icones/eventbridge.png",
    "QuickSight":   "icones/quicksight.png",
}


def carregar_arquitetura(caminho: str) -> dict:
    """Carrega e parseia o arquivo YAML de arquitetura."""
    with open(caminho, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


def gerar_dot(arquitetura: dict) -> str:
    """
    Gera representação DOT (Graphviz) da arquitetura.
    Pode ser convertida para PNG/SVG com: dot -Tpng arquivo.dot -o saida.png
    """
    linhas = ['digraph arquitetura {', '    rankdir=LR;', '    node [shape=box, style=filled, fillcolor="#E8F4FD"];']

    # Nós (componentes)
    for comp in arquitetura.get("componentes", []):
        label = f'{comp["nome"]}\\n({comp["tipo"]})'
        linhas.append(f'    {comp["id"]} [label="{label}"];')

    linhas.append("")

    # Arestas (conexões)
    for conn in arquitetura.get("conexoes", []):
        desc = conn.get("descricao", "")
        linhas.append(f'    {conn["origem"]} -> {conn["destino"]} [label="{desc}"];')

    linhas.append("}")
    return "\n".join(linhas)


def salvar_dot(conteudo_dot: str, caso: str) -> str:
    """Salva o arquivo .dot em saida/."""
    os.makedirs("saida", exist_ok=True)
    caminho = f"saida/{caso}.dot"
    with open(caminho, "w", encoding="utf-8") as f:
        f.write(conteudo_dot)
    return caminho


def renderizar(caminho_yaml: str) -> dict:
    """
    Pipeline principal de renderização.
    Retorna dict com caminhos dos arquivos gerados.
    """
    arquitetura = carregar_arquitetura(caminho_yaml)
    caso = arquitetura.get("caso", "sem-nome")

    dot = gerar_dot(arquitetura)
    caminho_dot = salvar_dot(dot, caso)

    resultado = {
        "caso": caso,
        "dot": caminho_dot,
        "timestamp": datetime.utcnow().isoformat() + "Z",
        "aviso": (
            "Para gerar PNG, execute: "
            f"dot -Tpng {caminho_dot} -o saida/{caso}.png"
        ),
    }

    print(json.dumps(resultado, indent=2, ensure_ascii=False))
    return resultado


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Renderiza arquitetura YAML em diagrama DOT/PNG"
    )
    parser.add_argument("yaml", help="Caminho para o arquivo de arquitetura YAML")
    args = parser.parse_args()

    renderizar(args.yaml)
