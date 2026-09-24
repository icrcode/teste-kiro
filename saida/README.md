# Saída

Pasta de saída gerada automaticamente pelo renderizador e validador.

## Arquivos Gerados
| Padrão                         | Gerado por          |
|-------------------------------|---------------------|
| `<caso>.dot`                  | renderizador.py     |
| `<caso>.png`                  | Graphviz (manual)   |
| `<caso>-validacao.json`       | validador.py        |

> Esta pasta é ignorada pelo git (adicione `saida/*.dot` e `saida/*.json` ao `.gitignore`).
