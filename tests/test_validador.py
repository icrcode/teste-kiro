"""
Testes para arquitetura/validador.py
Execute: python -m pytest tests/test_validador.py -v
"""

import pytest
import yaml
import os
import sys

# Garante que o módulo seja encontrado
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from arquitetura.validador import (
    validar_schema,
    validar_componentes,
    validar_conexoes,
    carregar_yaml,
    validar,
)


# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------

@pytest.fixture
def arquitetura_valida():
    return {
        "caso": "teste",
        "descricao": "Arquitetura de teste",
        "componentes": [
            {"id": "s3_raw", "tipo": "S3", "nome": "Bucket Raw"},
            {"id": "glue_etl", "tipo": "Glue", "nome": "ETL Job"},
        ],
        "conexoes": [
            {"origem": "s3_raw", "destino": "glue_etl", "descricao": "dados brutos"},
        ],
    }


@pytest.fixture
def yaml_valido_path(tmp_path, arquitetura_valida):
    caminho = tmp_path / "valido.yaml"
    caminho.write_text(yaml.dump(arquitetura_valida, allow_unicode=True))
    return str(caminho)


@pytest.fixture
def yaml_caso_real_path():
    """Aponta para o caso real do projeto."""
    return os.path.join(
        os.path.dirname(__file__), "..", "arquiteturas", "previsao-matriculas.yaml"
    )


# ---------------------------------------------------------------------------
# Testes: validar_schema
# ---------------------------------------------------------------------------

class TestValidarSchema:
    def test_schema_valido(self, arquitetura_valida):
        assert validar_schema(arquitetura_valida) == []

    def test_campo_caso_ausente(self, arquitetura_valida):
        del arquitetura_valida["caso"]
        erros = validar_schema(arquitetura_valida)
        assert any("caso" in e for e in erros)

    def test_campo_componentes_ausente(self, arquitetura_valida):
        del arquitetura_valida["componentes"]
        erros = validar_schema(arquitetura_valida)
        assert any("componentes" in e for e in erros)

    def test_todos_campos_ausentes(self):
        erros = validar_schema({})
        assert len(erros) == 4  # caso, descricao, componentes, conexoes


# ---------------------------------------------------------------------------
# Testes: validar_componentes
# ---------------------------------------------------------------------------

class TestValidarComponentes:
    def test_componentes_validos(self, arquitetura_valida):
        erros, avisos, ids = validar_componentes(arquitetura_valida)
        assert erros == []
        assert ids == {"s3_raw", "glue_etl"}

    def test_servico_desconhecido_gera_aviso(self, arquitetura_valida):
        arquitetura_valida["componentes"][0]["tipo"] = "ServicoDesconhecido"
        erros, avisos, _ = validar_componentes(arquitetura_valida)
        assert erros == []
        assert any("ServicoDesconhecido" in a for a in avisos)

    def test_id_duplicado(self, arquitetura_valida):
        arquitetura_valida["componentes"][1]["id"] = "s3_raw"  # duplicado
        erros, _, _ = validar_componentes(arquitetura_valida)
        assert any("duplicado" in e for e in erros)

    def test_campo_obrigatorio_ausente(self, arquitetura_valida):
        del arquitetura_valida["componentes"][0]["nome"]
        erros, _, _ = validar_componentes(arquitetura_valida)
        assert any("nome" in e for e in erros)


# ---------------------------------------------------------------------------
# Testes: validar_conexoes
# ---------------------------------------------------------------------------

class TestValidarConexoes:
    def test_conexoes_validas(self, arquitetura_valida):
        ids = {"s3_raw", "glue_etl"}
        erros = validar_conexoes(arquitetura_valida, ids)
        assert erros == []

    def test_origem_inexistente(self, arquitetura_valida):
        arquitetura_valida["conexoes"][0]["origem"] = "id_inexistente"
        ids = {"s3_raw", "glue_etl"}
        erros = validar_conexoes(arquitetura_valida, ids)
        assert any("origem" in e for e in erros)

    def test_destino_inexistente(self, arquitetura_valida):
        arquitetura_valida["conexoes"][0]["destino"] = "id_inexistente"
        ids = {"s3_raw", "glue_etl"}
        erros = validar_conexoes(arquitetura_valida, ids)
        assert any("destino" in e for e in erros)

    def test_conexao_para_si_mesmo(self, arquitetura_valida):
        arquitetura_valida["conexoes"][0]["destino"] = "s3_raw"
        ids = {"s3_raw", "glue_etl"}
        erros = validar_conexoes(arquitetura_valida, ids)
        assert any("ele mesmo" in e for e in erros)


# ---------------------------------------------------------------------------
# Testes: integração com arquivo real
# ---------------------------------------------------------------------------

class TestIntegracao:
    def test_yaml_valido_retorna_status_valido(self, yaml_valido_path):
        relatorio = validar(yaml_valido_path)
        assert relatorio["status"] == "valido"
        assert relatorio["erros"] == []

    def test_arquivo_inexistente(self, tmp_path):
        relatorio = validar(str(tmp_path / "nao_existe.yaml"))
        assert relatorio["status"] == "invalido"
        assert any("não encontrado" in e for e in relatorio["erros"])

    def test_caso_real_previsao_matriculas(self, yaml_caso_real_path):
        """Valida o caso de uso real do projeto."""
        if not os.path.exists(yaml_caso_real_path):
            pytest.skip("Arquivo de arquitetura real não encontrado.")
        relatorio = validar(yaml_caso_real_path)
        assert relatorio["status"] == "valido", (
            f"Erros encontrados: {relatorio['erros']}"
        )

    def test_strict_mode_promove_avisos_a_erros(self, tmp_path):
        """No modo strict, serviço desconhecido vira erro."""
        arq = {
            "caso": "teste-strict",
            "descricao": "teste",
            "componentes": [
                {"id": "x", "tipo": "ServicoFake", "nome": "Fake"},
            ],
            "conexoes": [],
        }
        caminho = tmp_path / "strict.yaml"
        caminho.write_text(yaml.dump(arq, allow_unicode=True))
        relatorio = validar(str(caminho), strict=True)
        assert relatorio["status"] == "invalido"
