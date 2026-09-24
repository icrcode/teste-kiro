# Design: Validador de Arquitetura

## Componentes

### `arquitetura/validador.py`
Módulo principal de validação. Expõe a função `validar(caminho_yaml: str) -> dict`.

### Fluxo de Execução
```
arquiteturas/<caso>.yaml
        ↓
  [carregar_yaml()]       — lê e parseia o arquivo
        ↓
  [validar_schema()]      — verifica campos obrigatórios
        ↓
  [validar_componentes()] — checa serviços AWS válidos
        ↓
  [validar_conexoes()]    — verifica integridade das referências
        ↓
  [gerar_relatorio()]     — salva saida/<caso>-validacao.json
```

## Estrutura de Dados

### Resultado da Validação
```python
{
    "status": "valido" | "invalido",
    "erros": ["mensagem de erro 1", ...],
    "avisos": ["aviso 1", ...],
    "timestamp": "2026-09-24T10:00:00Z"
}
```
