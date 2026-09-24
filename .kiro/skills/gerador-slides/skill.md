# Skill: Gerador de Slides SENAI

## Descrição
Gera apresentações PowerPoint no padrão visual SENAI 2026 (Open Sans, ícones
Material Symbols, azul `#164194` e laranja `#E84910`, gráficos em HD).

## Como Usar

### Caso previsão de matrículas
```bash
python slides/gerar_slides.py --slides 12 --modo aprofundado
```
Saída: `slides/previsao-matriculas.pptx`

### Qualquer tema
```bash
python slides/gerar_tema.py --tema "<tema>" --uc "<unidade curricular>" --slides 10 --modo resumido
```
Saída: `slides/apresentacoes/<tema>.pptx` e `<tema>.json` (conteúdo editável).

Quando o script pede ao Kiro para escrever o conteúdo, ele cria
`slides/apresentacoes/<tema>.prompt.md`. Nesse caso:
1. Leia o `.prompt.md` e siga as instruções dele.
2. Salve o conteúdo no campo `"conteudo"` do `<tema>.json` indicado, sem alterar `"meta"`.
3. Não rode comandos: o script no terminal valida e monta o `.pptx` sozinho.
   Se o script não estiver aberto, monte com
   `python slides/gerar_tema.py --de-json slides/apresentacoes/<tema>.json`.

## Inputs
- `tema`: assunto da apresentação
- `uc`: unidade curricular ou contexto (opcional)
- `modo`: `resumido` | `aprofundado`
- `slides`: quantidade total (3 a 30; `gerar_slides.py` aceita 3 a 18)

## Outputs
- `.pptx` com capa, agenda, divisórias de seção, conteúdo e encerramento SENAI
- Notas do apresentador em cada slide

Documentação completa: `slides/README.md`.
