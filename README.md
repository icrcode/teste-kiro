# Previsão de Demanda de Vagas por Curso

Estrutura base para um modelo de **previsão de matrículas** por curso, construído com serviços AWS e orquestrado via Kiro (workshop-arquiteturas-kiro).

## Objetivo
Prever a demanda de vagas por curso e período letivo usando histórico de matrículas, perfil de alunos e dados socioeconômicos.

## Estrutura do Projeto

```
.
├── .kiro/
│   ├── steering/
│   │   └── guia-visual-aws.md          # Contexto AWS e padrões do projeto
│   ├── skills/
│   │   ├── gerador-arquitetura/        # Skill para gerar arquiteturas
│   │   └── gerador-slides/             # Skill para gerar slides SENAI
│   ├── specs/
│   │   └── validador-arquitetura/      # Spec do validador (req + design + tasks)
│   └── hooks/                          # Hooks de automação
├── renderizador/
│   └── renderizador.py                 # Gera diagrama DOT a partir do YAML
├── icones/                             # Ícones PNG dos serviços AWS
├── arquitetura/
│   └── validador.py                    # Valida arquivos YAML de arquitetura
├── arquiteturas/
│   └── previsao-matriculas.yaml        # Arquitetura do caso de uso
├── casos/
│   └── previsao-matriculas.yaml        # Caso de uso detalhado com features e modelo
├── slides/
│   ├── gerar_slides.py                 # Slides SENAI do caso previsão de matrículas
│   ├── gerar_tema.py                   # Slides SENAI sobre qualquer tema (Claude/Kiro)
│   ├── exportar_previa.py              # Exporta slides .pptx como imagens PNG
│   └── assets/                         # Modelo SENAI, logos, fontes e ícones Google
├── resultado-slide/                    # Apresentações geradas (.pptx) e imagens de prévia
├── tests/
│   └── test_validador.py               # Testes do validador
└── saida/                              # Diagramas e relatórios gerados
```

## Início Rápido

```bash
pip install -r requirements.txt
```

### 1. Validar a arquitetura
```bash
python arquitetura/validador.py arquiteturas/previsao-matriculas.yaml
```

### 2. Renderizar o diagrama
```bash
python renderizador/renderizador.py arquiteturas/previsao-matriculas.yaml
# Gera: saida/previsao-matriculas.dot

# Para converter para PNG (requer Graphviz instalado):
dot -Tpng saida/previsao-matriculas.dot -o saida/previsao-matriculas.png
```

### 3. Rodar os testes
```bash
python -m pytest tests/ -v
```

### 4. Gerar apresentações SENAI
```bash
python slides/gerar_slides.py    # caso previsão de matrículas
python slides/gerar_tema.py      # qualquer tema, conteúdo escrito pelo Claude ou pelo Kiro
```
Detalhes, opções e galeria de resultados em [slides/README.md](slides/README.md).

![Visão geral da apresentação gerada para o caso previsão de matrículas](resultado-slide/previa/previsao-matriculas/grade.png)

## Modelo de ML

| Atributo            | Valor                          |
|---------------------|-------------------------------|
| Algoritmo principal | XGBoost                        |
| Alternativas        | Prophet, LSTM                  |
| Horizonte           | 2 semestres à frente           |
| Métrica principal   | MAPE (threshold: 15%)          |
| Retreinamento       | A cada semestre (EventBridge)  |

## Pipeline AWS

```
S3 (raw) → Glue ETL → S3 (processed) → SageMaker Feature Store
         → SageMaker Training → SageMaker Endpoint
         → Model Monitor → CloudWatch → EventBridge (retreat)
         → Lambda → Endpoint → QuickSight Dashboard
```

## Configuração MCP
Acesse `.kiro/settings/mcp.json` para configurar os servidores MCP.
Use o comando `Open Kiro MCP` na paleta de comandos do Kiro para gerenciar.
