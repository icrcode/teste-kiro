# Skill: Gerador de Arquitetura — Previsão de Matrículas

## Descrição
Gera diagramas e descrições de arquitetura AWS para o modelo de previsão de demanda de vagas por curso.

## Como Usar
Invoque esta skill quando precisar:
- Criar ou atualizar o diagrama de arquitetura do pipeline de ML
- Descrever os componentes AWS envolvidos em cada etapa
- Documentar fluxo de dados entre serviços

## Inputs Esperados
- `caso`: nome do caso de uso (ex: `previsao-matriculas`)
- `componentes`: lista de serviços AWS a incluir
- `formato`: `diagrama` | `yaml` | `markdown`

## Outputs Gerados
- Arquivo de arquitetura em `arquiteturas/<caso>.yaml`
- Diagrama visual em `saida/<caso>.png` (via renderizador)

## Template de Arquitetura

```yaml
caso: <nome-do-caso>
descricao: <descricao-curta>
componentes:
  - id: <id-unico>
    tipo: <servico-aws>
    nome: <nome-descritivo>
    config: {}
conexoes:
  - origem: <id>
    destino: <id>
    descricao: <fluxo>
```

## Instruções ao Agente
1. Leia o caso de uso em `casos/<caso>.yaml`
2. Identifique os serviços AWS necessários com base no `guia-visual-aws.md`
3. Gere o arquivo `arquiteturas/<caso>.yaml` seguindo o template acima
4. Chame o renderizador para gerar o diagrama visual
5. Valide com `arquitetura/validador.py` antes de salvar em `saida/`
