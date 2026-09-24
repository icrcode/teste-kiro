# Spec: Validador de Arquitetura — Previsão de Matrículas

## Visão Geral
Especificação para o componente de validação de arquiteturas AWS geradas pelo modelo de previsão de demanda de vagas por curso.

## Requisitos Funcionais

### RF-01: Validação de Schema
- O validador deve verificar se o arquivo `arquiteturas/<caso>.yaml` está no formato correto
- Campos obrigatórios: `caso`, `descricao`, `componentes`, `conexoes`
- Cada componente deve ter: `id`, `tipo`, `nome`

### RF-02: Validação de Componentes AWS
- Verificar se os `tipo` dos componentes são serviços AWS válidos e suportados
- Lista de serviços suportados: S3, Glue, SageMaker, Lambda, RDS, Aurora, EMR, CloudWatch, EventBridge, QuickSight

### RF-03: Validação de Conexões
- Todas as `origem` e `destino` em `conexoes` devem referenciar `id` existentes em `componentes`
- Não deve haver conexões circulares que causem loops infinitos no pipeline

### RF-04: Relatório de Validação
- Gerar relatório em `saida/<caso>-validacao.json` com:
  - `status`: `valido` | `invalido`
  - `erros`: lista de erros encontrados
  - `avisos`: lista de avisos (não bloqueantes)
  - `timestamp`: data/hora da validação

## Requisitos Não Funcionais

### RNF-01: Performance
- Validação deve completar em menos de 5 segundos para arquiteturas com até 50 componentes

### RNF-02: Mensagens de Erro
- Mensagens de erro devem ser claras e indicar exatamente o campo/componente com problema

## Critérios de Aceite
- [ ] Schema YAML inválido retorna erro com linha problemática
- [ ] Serviço AWS desconhecido retorna aviso com sugestão de correção
- [ ] Conexão com ID inexistente retorna erro bloqueante
- [ ] Arquitetura válida gera relatório com `status: valido`
