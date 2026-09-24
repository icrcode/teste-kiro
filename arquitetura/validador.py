"""
Validador de Arquiteturas AWS — Previsão de Demanda de Vagas por Curso
Verifica schema, componentes e conexões de arquivos YAML de arquitetura.
"""

import yaml
import json
import os
import argparse
from datetime import datetime, timezone


# ---------------------------------------------------------------------------
# Serviços AWS suportados pelo modelo de previsão de matrículas
# ---------------------------------------------------------------------------
SERVICOS_AWS_SUPORTADOS = {
    "S3", "Glue", "SageMaker", "Lambda", "RDS", "Aurora",
    "EMR", "CloudWatch", "EventBridge", "QuickSight",
    "Athena", "Kinesis", "StepFunctions", "SNS", "SQS",
}

CAMPOS_OBRIGATORIOS_RAIZ = ["caso", "descricao", "componentes", "conexoes"]
CAMPOS_OBRIGATORIOS_COMPONENTE = ["id", "tipo", "nome"]


# ---------------------------------------------------------------------------
# Funções de validação
# ---------------------------------------------------------------------------

def carregar_yaml(caminho: str) -> tuple[dict | None, str | None]:
    """Carrega o YAML; retorna (dados, None) ou (None, mensagem_erro)."""
    try:
        with open(caminho, "r", encoding="utf-8") as f:
            dados = yaml.safe_load(f)
        if not isinstance(dados, dict):
            return None, "O arquivo YAML deve conter um mapeamento (dicionário) na raiz."
        return dados, None
    except FileNotFoundError:
        return None, f"Arquivo não encontrado: {caminho}"
    except yaml.YAMLError as e:
        return None, f"Erro de parsing YAML: {e}"


def validar_schema(dados: dict) -> list[str]:
    """Verifica campos obrigatórios na raiz."""
    erros = []
    for campo in CAMPOS_OBRIGATORIOS_RAIZ:
        if campo not in dados:
            erros.append(f"Campo obrigatório ausente na raiz: '{campo}'")
    return erros


def validar_componentes(dados: dict) -> tuple[list[str], list[str], set[str]]:
    """
    Valida cada componente.
    Retorna (erros, avisos, ids_encontrados).
    """
    erros, avisos, ids = [], [], set()
    componentes = dados.get("componentes", [])

    if not isinstance(componentes, list):
        return ["'componentes' deve ser uma lista."], [], set()

    for i, comp in enumerate(componentes):
        prefixo = f"componentes[{i}]"

        if not isinstance(comp, dict):
            erros.append(f"{prefixo}: deve ser um dicionário.")
            continue

        # Campos obrigatórios
        for campo in CAMPOS_OBRIGATORIOS_COMPONENTE:
            if campo not in comp:
                erros.append(f"{prefixo}: campo obrigatório ausente: '{campo}'")

        # ID duplicado
        comp_id = comp.get("id")
        if comp_id:
            if comp_id in ids:
                erros.append(f"{prefixo}: id duplicado: '{comp_id}'")
            else:
                ids.add(comp_id)

        # Tipo (serviço AWS) suportado
        tipo = comp.get("tipo", "")
        if tipo and tipo not in SERVICOS_AWS_SUPORTADOS:
            avisos.append(
                f"{prefixo} (id='{comp_id}'): serviço AWS '{tipo}' não está na lista "
                f"de suportados. Verifique o nome ou adicione ao validador."
            )

    return erros, avisos, ids


def validar_conexoes(dados: dict, ids_validos: set[str]) -> list[str]:
    """Verifica integridade das referências em conexoes."""
    erros = []
    conexoes = dados.get("conexoes", [])

    if not isinstance(conexoes, list):
        return ["'conexoes' deve ser uma lista."]

    for i, conn in enumerate(conexoes):
        prefixo = f"conexoes[{i}]"

        if not isinstance(conn, dict):
            erros.append(f"{prefixo}: deve ser um dicionário.")
            continue

        origem = conn.get("origem")
        destino = conn.get("destino")

        if not origem:
            erros.append(f"{prefixo}: campo 'origem' ausente ou vazio.")
        elif origem not in ids_validos:
            erros.append(f"{prefixo}: 'origem' referencia id inexistente: '{origem}'")

        if not destino:
            erros.append(f"{prefixo}: campo 'destino' ausente ou vazio.")
        elif destino not in ids_validos:
            erros.append(f"{prefixo}: 'destino' referencia id inexistente: '{destino}'")

        if origem and destino and origem == destino:
            erros.append(f"{prefixo}: conexão de um componente para ele mesmo ('{origem}').")

    return erros


def gerar_relatorio(caso: str, erros: list[str], avisos: list[str]) -> dict:
    """Monta e salva o relatório de validação em saida/."""
    status = "invalido" if erros else "valido"
    relatorio = {
        "caso": caso,
        "status": status,
        "erros": erros,
        "avisos": avisos,
        "timestamp": datetime.now(timezone.utc).isoformat(),
    }

    os.makedirs("saida", exist_ok=True)
    caminho = f"saida/{caso}-validacao.json"
    with open(caminho, "w", encoding="utf-8") as f:
        json.dump(relatorio, f, indent=2, ensure_ascii=False)

    return relatorio


# ---------------------------------------------------------------------------
# Ponto de entrada principal
# ---------------------------------------------------------------------------

def validar(caminho_yaml: str, strict: bool = False) -> dict:
    """
    Valida um arquivo de arquitetura YAML.

    Args:
        caminho_yaml: Caminho para o arquivo YAML.
        strict: Se True, trata avisos como erros.

    Returns:
        Dicionário com o relatório de validação.
    """
    todos_erros: list[str] = []
    todos_avisos: list[str] = []

    # 1. Carregar
    dados, erro_carga = carregar_yaml(caminho_yaml)
    if erro_carga:
        todos_erros.append(erro_carga)
        return gerar_relatorio("desconhecido", todos_erros, todos_avisos)

    caso = dados.get("caso", "desconhecido")

    # 2. Schema
    todos_erros.extend(validar_schema(dados))

    # 3. Componentes
    erros_comp, avisos_comp, ids_validos = validar_componentes(dados)
    todos_erros.extend(erros_comp)
    todos_avisos.extend(avisos_comp)

    # 4. Conexões (só valida se schema/componentes estiverem OK)
    if not todos_erros:
        todos_erros.extend(validar_conexoes(dados, ids_validos))

    # 5. Strict mode: promove avisos a erros
    if strict:
        todos_erros.extend(todos_avisos)
        todos_avisos = []

    return gerar_relatorio(caso, todos_erros, todos_avisos)


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Valida arquivos de arquitetura YAML para o modelo de previsão de matrículas."
    )
    parser.add_argument("yaml", help="Caminho para o arquivo de arquitetura YAML")
    parser.add_argument(
        "--strict",
        action="store_true",
        help="Trata avisos como erros (modo rigoroso)",
    )
    args = parser.parse_args()

    relatorio = validar(args.yaml, strict=args.strict)
    print(json.dumps(relatorio, indent=2, ensure_ascii=False))

    # Exit code 1 se inválido
    if relatorio["status"] == "invalido":
        raise SystemExit(1)
