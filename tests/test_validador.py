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
    gerar_relatorio,
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
        erros, avisos = validar_conexoes(arquitetura_valida, ids)
        assert erros == []
        assert avisos == []

    def test_origem_inexistente(self, arquitetura_valida):
        arquitetura_valida["conexoes"][0]["origem"] = "id_inexistente"
        ids = {"s3_raw", "glue_etl"}
        erros, _ = validar_conexoes(arquitetura_valida, ids)
        assert any("origem" in e for e in erros)

    def test_destino_inexistente(self, arquitetura_valida):
        arquitetura_valida["conexoes"][0]["destino"] = "id_inexistente"
        ids = {"s3_raw", "glue_etl"}
        erros, _ = validar_conexoes(arquitetura_valida, ids)
        assert any("destino" in e for e in erros)

    def test_conexao_para_si_mesmo(self, arquitetura_valida):
        arquitetura_valida["conexoes"][0]["destino"] = "s3_raw"
        ids = {"s3_raw", "glue_etl"}
        erros, _ = validar_conexoes(arquitetura_valida, ids)
        assert any("ele mesmo" in e for e in erros)

    def test_ciclo_simples_a_b_a(self):
        """A → B → A deve ser detectado como aviso de ciclo."""
        dados = {
            "conexoes": [
                {"origem": "A", "destino": "B"},
                {"origem": "B", "destino": "A"},
            ]
        }
        ids = {"A", "B"}
        erros, avisos = validar_conexoes(dados, ids)
        assert erros == []
        assert any("Ciclo" in a for a in avisos)

    def test_ciclo_longo_a_b_c_a(self):
        """A → B → C → A deve ser detectado como aviso de ciclo."""
        dados = {
            "conexoes": [
                {"origem": "A", "destino": "B"},
                {"origem": "B", "destino": "C"},
                {"origem": "C", "destino": "A"},
            ]
        }
        ids = {"A", "B", "C"}
        erros, avisos = validar_conexoes(dados, ids)
        assert erros == []
        assert any("Ciclo" in a for a in avisos)

    def test_dag_sem_ciclos(self):
        """Grafo acíclico dirigido não deve reportar ciclos."""
        dados = {
            "conexoes": [
                {"origem": "A", "destino": "B"},
                {"origem": "B", "destino": "C"},
                {"origem": "A", "destino": "C"},
            ]
        }
        ids = {"A", "B", "C"}
        erros, avisos = validar_conexoes(dados, ids)
        assert not any("Ciclo" in a for a in avisos)
        assert erros == []


# ---------------------------------------------------------------------------
# Testes: gerar_relatorio
# ---------------------------------------------------------------------------

class TestGerarRelatorio:
    def test_status_valido_quando_sem_erros(self, tmp_path, monkeypatch):
        monkeypatch.chdir(tmp_path)
        relatorio = gerar_relatorio("meu-caso", [], [])
        assert relatorio["status"] == "valido"

    def test_status_invalido_quando_ha_erros(self, tmp_path, monkeypatch):
        monkeypatch.chdir(tmp_path)
        relatorio = gerar_relatorio("meu-caso", ["erro qualquer"], [])
        assert relatorio["status"] == "invalido"

    def test_campos_obrigatorios_presentes(self, tmp_path, monkeypatch):
        monkeypatch.chdir(tmp_path)
        relatorio = gerar_relatorio("meu-caso", ["erro"], ["aviso"])
        assert "status" in relatorio
        assert "erros" in relatorio
        assert "avisos" in relatorio
        assert "timestamp" in relatorio

    def test_arquivo_json_criado_em_saida(self, tmp_path, monkeypatch):
        monkeypatch.chdir(tmp_path)
        gerar_relatorio("meu-caso", [], [])
        assert (tmp_path / "saida" / "meu-caso-validacao.json").exists()

    def test_conteudo_json_correto(self, tmp_path, monkeypatch):
        import json
        monkeypatch.chdir(tmp_path)
        gerar_relatorio("meu-caso", ["erro1"], ["aviso1"])
        with open(tmp_path / "saida" / "meu-caso-validacao.json", encoding="utf-8") as f:
            dados = json.load(f)
        assert dados["status"] == "invalido"
        assert dados["erros"] == ["erro1"]
        assert dados["avisos"] == ["aviso1"]

    def test_timestamp_formato_iso8601(self, tmp_path, monkeypatch):
        from datetime import datetime
        monkeypatch.chdir(tmp_path)
        relatorio = gerar_relatorio("meu-caso", [], [])
        # Deve ser parseável como datetime ISO 8601 com timezone
        ts = datetime.fromisoformat(relatorio["timestamp"])
        assert ts.tzinfo is not None

    def test_cria_diretorio_saida_se_nao_existir(self, tmp_path, monkeypatch):
        monkeypatch.chdir(tmp_path)
        assert not (tmp_path / "saida").exists()
        gerar_relatorio("meu-caso", [], [])
        assert (tmp_path / "saida").is_dir()

    def test_nome_arquivo_usa_campo_caso(self, tmp_path, monkeypatch):
        monkeypatch.chdir(tmp_path)
        gerar_relatorio("previsao-matriculas", [], [])
        assert (tmp_path / "saida" / "previsao-matriculas-validacao.json").exists()


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
