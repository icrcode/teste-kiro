# Tasks: Validador de Arquitetura

## Fase 1 — Core do Validador
- [ ] Implementar `carregar_yaml()` com tratamento de erros de parsing
- [ ] Implementar `validar_schema()` com checagem de campos obrigatórios
- [ ] Implementar `validar_componentes()` com lista de serviços AWS suportados
- [ ] Implementar `validar_conexoes()` com verificação de referências

## Fase 2 — Relatório e CLI
- [ ] Implementar `gerar_relatorio()` gravando JSON em `saida/`
- [ ] Criar interface CLI: `python validador.py <caminho_yaml>`
- [ ] Adicionar flag `--strict` para tratar avisos como erros

## Fase 3 — Testes
- [ ] Escrever testes para schema inválido
- [ ] Escrever testes para serviço AWS desconhecido
- [ ] Escrever testes para conexão com ID inexistente
- [ ] Escrever teste de arquitetura válida completa (caso previsao-matriculas)
