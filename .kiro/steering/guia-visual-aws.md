# Guia Visual AWS — Previsão de Demanda de Vagas por Curso

## Objetivo
Este projeto implementa um modelo de **previsão de matrículas** por curso, utilizando serviços AWS para ingestão, processamento, treinamento e inferência de dados educacionais.

## Contexto do Domínio
- **Problema:** Prever a demanda de vagas por curso em instituições de ensino
- **Entrada:** Histórico de matrículas, perfil de alunos, calendário acadêmico, dados socioeconômicos
- **Saída:** Estimativa de matrículas por curso e período letivo

## Serviços AWS Utilizados

### Ingestão e Armazenamento
- **S3**: Armazenamento de dados brutos (histórico de matrículas, catálogo de cursos)
- **Glue**: Catalogação e ETL dos dados educacionais
- **RDS / Aurora**: Banco relacional para dados estruturados (alunos, cursos, períodos)

### Processamento e Feature Engineering
- **Glue ETL Jobs**: Transformação e limpeza dos dados
- **EMR (opcional)**: Processamento distribuído para grandes volumes
- **Lambda**: Funções serverless para pré-processamento em tempo real

### Machine Learning
- **SageMaker**: Treinamento, validação e hospedagem do modelo de previsão
  - Algoritmos sugeridos: XGBoost, LSTM (séries temporais), Prophet
- **SageMaker Feature Store**: Repositório centralizado de features
- **SageMaker Pipelines**: Orquestração do fluxo de ML

### Monitoramento e Operação
- **CloudWatch**: Monitoramento de métricas do modelo e infraestrutura
- **SageMaker Model Monitor**: Detecção de drift de dados e modelo
- **EventBridge**: Agendamento de retreinamento periódico

## Fluxo de Arquitetura

```
[Fontes de Dados]
  Histórico Matrículas → S3 (raw)
  Catálogo de Cursos   → S3 (raw)
  Dados Socioeconômicos → S3 (raw)
        ↓
[ETL - Glue]
  Limpeza + Transformação → S3 (processed)
        ↓
[Feature Store - SageMaker]
  Features: curso_id, periodo, historico_demanda,
            taxa_evasao, vagas_ofertadas, etc.
        ↓
[Treinamento - SageMaker]
  Modelo: XGBoost / Prophet / LSTM
        ↓
[Endpoint de Inferência - SageMaker]
  Previsão de matrículas por curso/período
        ↓
[Visualização]
  QuickSight Dashboard → Gestores Acadêmicos
```

## Padrões de Nomenclatura
- Buckets S3: `prevmatriculas-{env}-{tipo}` (ex: `prevmatriculas-prod-raw`)
- Jobs Glue: `job-prevmatriculas-{etapa}` (ex: `job-prevmatriculas-feature-engineering`)
- Modelos SageMaker: `modelo-prevmatriculas-v{versao}` (ex: `modelo-prevmatriculas-v1`)

## Variáveis de Ambiente Esperadas
```
AWS_REGION=us-east-1
S3_BUCKET_RAW=prevmatriculas-dev-raw
S3_BUCKET_PROCESSED=prevmatriculas-dev-processed
SAGEMAKER_ROLE_ARN=arn:aws:iam::ACCOUNT_ID:role/SageMakerExecutionRole
FEATURE_GROUP_NAME=prevmatriculas-features
```
