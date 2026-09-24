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
# Testes: carregar_yaml — YAML inválido / não-dicionário
# ---------------------------------------------------------------------------

class TestCarregarYaml:
    def test_yaml_malformado_retorna_erro(self, tmp_path):
        """YAML com sintaxe inválida deve retornar mensagem de erro (não lançar exceção)."""
        caminho = tmp_path / "malformado.yaml"
        caminho.write_text("chave: [sem fechar\n", encoding="utf-8")
        dados, erro = carregar_yaml(str(caminho))
        assert dados is None
        assert erro is not None
        assert "YAML" in erro or "parsing" in erro.lower() or "erro" in erro.lower()

    def test_yaml_raiz_lista_retorna_erro(self, tmp_path):
        """YAML válido mas com lista na raiz (não dicionário) deve retornar erro."""
        caminho = tmp_path / "lista.yaml"
        caminho.write_text("- item1\n- item2\n", encoding="utf-8")
        dados, erro = carregar_yaml(str(caminho))
        assert dados is None
        assert erro is not None
        assert "mapeamento" in erro or "dicionário" in erro or "dict" in erro.lower()

    def test_yaml_valido_retorna_dados_sem_erro(self, tmp_path):
        """YAML válido com dicionário na raiz retorna dados e erro None."""
        caminho = tmp_path / "ok.yaml"
        caminho.write_text("caso: teste\n", encoding="utf-8")
        dados, erro = carregar_yaml(str(caminho))
        assert dados == {"caso": "teste"}
        assert erro is None

    def test_arquivo_inexistente_retorna_erro(self, tmp_path):
        dados, erro = carregar_yaml(str(tmp_path / "nao_existe.yaml"))
        assert dados is None
        assert erro is not None
        assert "não encontrado" in erro or "not found" in erro.lower()


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

    def test_campo_descricao_ausente(self, arquitetura_valida):
        del arquitetura_valida["descricao"]
        erros = validar_schema(arquitetura_valida)
        assert any("descricao" in e for e in erros)

    def test_campo_conexoes_ausente(self, arquitetura_valida):
        del arquitetura_valida["conexoes"]
        erros = validar_schema(arquitetura_valida)
        assert any("conexoes" in e for e in erros)

    def test_todos_campos_ausentes(self):
        erros = validar_schema({})
        assert len(erros) == 4  # caso, descricao, componentes, conexoes

    def test_mensagem_erro_contem_nome_do_campo(self):
        """RNF-02: mensagens de erro identificam claramente o campo ausente."""
        for campo in ["caso", "descricao", "componentes", "conexoes"]:
            erros = validar_schema({campo: "presente"})
            # Os outros 3 campos estão ausentes; cada mensagem deve conter seu nome
            ausentes = {"caso", "descricao", "componentes", "conexoes"} - {campo}
            for ausente in ausentes:
                assert any(ausente in e for e in erros), (
                    f"Mensagem de erro não menciona o campo ausente '{ausente}'"
                )


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

    def test_campo_id_ausente(self, arquitetura_valida):
        del arquitetura_valida["componentes"][0]["id"]
        erros, _, _ = validar_componentes(arquitetura_valida)
        assert any("id" in e for e in erros)

    def test_campo_tipo_ausente(self, arquitetura_valida):
        del arquitetura_valida["componentes"][0]["tipo"]
        erros, _, _ = validar_componentes(arquitetura_valida)
        assert any("tipo" in e for e in erros)

    def test_campo_nome_ausente(self, arquitetura_valida):
        del arquitetura_valida["componentes"][0]["nome"]
        erros, _, _ = validar_componentes(arquitetura_valida)
        assert any("nome" in e for e in erros)

    def test_todos_campos_componente_ausentes(self, arquitetura_valida):
        """Componente vazio gera erros para id, tipo e nome."""
        arquitetura_valida["componentes"][0] = {}
        erros, _, _ = validar_componentes(arquitetura_valida)
        for campo in ("id", "tipo", "nome"):
            assert any(campo in e for e in erros), (
                f"Mensagem de erro não menciona o campo ausente '{campo}'"
            )

    def test_mensagem_erro_componente_contem_nome_do_campo(self, arquitetura_valida):
        """RNF-02: mensagem de erro do componente identifica o campo ausente."""
        del arquitetura_valida["componentes"][0]["tipo"]
        erros, _, _ = validar_componentes(arquitetura_valida)
        assert any("tipo" in e for e in erros)

    def test_aviso_servico_desconhecido_contem_nome_do_servico(self, arquitetura_valida):
        """RNF-02: aviso para serviço desconhecido deve incluir o nome do serviço."""
        arquitetura_valida["componentes"][0]["tipo"] = "ServicoFalso"
        _, avisos, _ = validar_componentes(arquitetura_valida)
        assert any("ServicoFalso" in a for a in avisos)

    def test_aviso_servico_desconhecido_sugere_correcao(self, arquitetura_valida):
        """RF-02 / Critério: aviso deve sugerir correção ao usuário."""
        arquitetura_valida["componentes"][0]["tipo"] = "DynamoDB"
        _, avisos, _ = validar_componentes(arquitetura_valida)
        # A mensagem deve orientar o usuário a verificar ou corrigir
        assert any(
            "verifique" in a.lower() or "adicione" in a.lower() or "suportado" in a.lower()
            for a in avisos
        )

    @pytest.mark.parametrize("servico", [
        "S3", "Glue", "SageMaker", "Lambda", "RDS",
        "Aurora", "EMR", "CloudWatch", "EventBridge", "QuickSight",
    ])
    def test_todos_servicos_suportados_sem_aviso(self, arquitetura_valida, servico):
        """RF-02: todos os 10 serviços AWS suportados não devem gerar avisos."""
        arquitetura_valida["componentes"][0]["tipo"] = servico
        erros, avisos, _ = validar_componentes(arquitetura_valida)
        assert erros == []
        assert not any(servico in a for a in avisos), (
            f"Serviço suportado '{servico}' gerou aviso inesperado."
        )

    def test_servico_case_sensitive_minusculo_gera_aviso(self, arquitetura_valida):
        """RF-02: verificação é case-sensitive — 's3' não é igual a 'S3'."""
        arquitetura_valida["componentes"][0]["tipo"] = "s3"
        _, avisos, _ = validar_componentes(arquitetura_valida)
        assert any("s3" in a for a in avisos), (
            "Serviço 's3' (minúsculo) deveria gerar aviso pois a checagem é case-sensitive."
        )

    def test_servico_case_sensitive_parcial_gera_aviso(self, arquitetura_valida):
        """RF-02: variantes de capitalização como 'sagemaker' geram aviso."""
        arquitetura_valida["componentes"][0]["tipo"] = "sagemaker"
        _, avisos, _ = validar_componentes(arquitetura_valida)
        assert any("sagemaker" in a for a in avisos)


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

    def test_origem_e_destino_inexistentes_geram_dois_erros(self, arquitetura_valida):
        """Quando tanto origem quanto destino referenciam IDs inexistentes, dois erros são retornados."""
        arquitetura_valida["conexoes"][0]["origem"] = "id_falso_origem"
        arquitetura_valida["conexoes"][0]["destino"] = "id_falso_destino"
        ids = {"s3_raw", "glue_etl"}
        erros, _ = validar_conexoes(arquitetura_valida, ids)
        assert len([e for e in erros if "origem" in e or "destino" in e]) >= 2

    def test_mensagem_erro_origem_contem_id_desconhecido(self, arquitetura_valida):
        """RNF-02: mensagem de erro para origem inexistente deve conter o ID desconhecido."""
        arquitetura_valida["conexoes"][0]["origem"] = "id_nao_cadastrado"
        ids = {"s3_raw", "glue_etl"}
        erros, _ = validar_conexoes(arquitetura_valida, ids)
        assert any("id_nao_cadastrado" in e for e in erros), (
            "Mensagem de erro deve conter o ID referenciado que não existe."
        )

    def test_mensagem_erro_destino_contem_id_desconhecido(self, arquitetura_valida):
        """RNF-02: mensagem de erro para destino inexistente deve conter o ID desconhecido."""
        arquitetura_valida["conexoes"][0]["destino"] = "id_nao_cadastrado"
        ids = {"s3_raw", "glue_etl"}
        erros, _ = validar_conexoes(arquitetura_valida, ids)
        assert any("id_nao_cadastrado" in e for e in erros), (
            "Mensagem de erro deve conter o ID referenciado que não existe."
        )

    def test_campo_origem_ausente_gera_erro(self, arquitetura_valida):
        """Campo 'origem' ausente na conexão deve gerar erro bloqueante."""
        del arquitetura_valida["conexoes"][0]["origem"]
        ids = {"s3_raw", "glue_etl"}
        erros, _ = validar_conexoes(arquitetura_valida, ids)
        assert any("origem" in e for e in erros)

    def test_campo_destino_ausente_gera_erro(self, arquitetura_valida):
        """Campo 'destino' ausente na conexão deve gerar erro bloqueante."""
        del arquitetura_valida["conexoes"][0]["destino"]
        ids = {"s3_raw", "glue_etl"}
        erros, _ = validar_conexoes(arquitetura_valida, ids)
        assert any("destino" in e for e in erros)


# ---------------------------------------------------------------------------
# Testes: integração — conexão com ID inexistente via validar()
# ---------------------------------------------------------------------------

class TestValidarConexoesIntegracao:
    def test_conexao_origem_inexistente_retorna_status_invalido(self, tmp_path):
        """RF-03 / Critério: conexão com ID inexistente deve retornar status invalido."""
        arq = {
            "caso": "teste-conn-invalida",
            "descricao": "teste",
            "componentes": [
                {"id": "s3_raw", "tipo": "S3", "nome": "Bucket Raw"},
                {"id": "glue_etl", "tipo": "Glue", "nome": "ETL Job"},
            ],
            "conexoes": [
                {"origem": "id_que_nao_existe", "destino": "glue_etl"},
            ],
        }
        import yaml as _yaml
        caminho = tmp_path / "conn_invalida.yaml"
        caminho.write_text(_yaml.dump(arq, allow_unicode=True))
        relatorio = validar(str(caminho))
        assert relatorio["status"] == "invalido"
        assert any("id_que_nao_existe" in e for e in relatorio["erros"])


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

    def test_strict_mode_avisos_movidos_para_erros(self, tmp_path):
        """Em strict mode, avisos são movidos para erros e avisos fica vazio."""
        arq = {
            "caso": "teste-strict-avisos",
            "descricao": "teste",
            "componentes": [
                {"id": "x", "tipo": "ServicoFake", "nome": "Fake"},
            ],
            "conexoes": [],
        }
        caminho = tmp_path / "strict2.yaml"
        caminho.write_text(yaml.dump(arq, allow_unicode=True))
        relatorio = validar(str(caminho), strict=True)
        assert relatorio["avisos"] == []
        assert any("ServicoFake" in e for e in relatorio["erros"])

    def test_sem_strict_mode_avisos_permanecem_como_avisos(self, tmp_path):
        """Sem strict mode, serviço desconhecido gera aviso (não erro) e status é válido."""
        arq = {
            "caso": "teste-nao-strict",
            "descricao": "teste",
            "componentes": [
                {"id": "x", "tipo": "ServicoFake", "nome": "Fake"},
            ],
            "conexoes": [],
        }
        caminho = tmp_path / "nao_strict.yaml"
        caminho.write_text(yaml.dump(arq, allow_unicode=True))
        relatorio = validar(str(caminho), strict=False)
        assert relatorio["status"] == "valido"
        assert relatorio["erros"] == []
        assert any("ServicoFake" in a for a in relatorio["avisos"])

    def test_strict_mode_ciclo_vira_erro(self, tmp_path):
        """Em strict mode, ciclo detectado (aviso) deve tornar o relatório inválido."""
        arq = {
            "caso": "teste-strict-ciclo",
            "descricao": "teste",
            "componentes": [
                {"id": "A", "tipo": "S3", "nome": "Bucket A"},
                {"id": "B", "tipo": "Glue", "nome": "ETL B"},
            ],
            "conexoes": [
                {"origem": "A", "destino": "B"},
                {"origem": "B", "destino": "A"},
            ],
        }
        caminho = tmp_path / "strict_ciclo.yaml"
        caminho.write_text(yaml.dump(arq, allow_unicode=True))
        relatorio = validar(str(caminho), strict=True)
        assert relatorio["status"] == "invalido"
        assert relatorio["avisos"] == []
        assert any("Ciclo" in e for e in relatorio["erros"])
