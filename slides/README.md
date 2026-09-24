# Slides

Geradores de apresentações PowerPoint no padrão visual **SENAI 2026**, baseados no
modelo oficial em `assets/Modelo Senai 2026 - Indentidade Nova.pptx`:

- Tipografia **Open Sans** e ícones **Material Symbols** (Google Fonts)
- Azul SENAI `#164194` e laranja `#E84910`, logos `senai.png` e `técnico.png`
- Gráficos matplotlib renderizados em HD (300 dpi) no tamanho exato do slide
- Transição fade e entrada animada dos gráficos
- Notas do apresentador em cada slide

Instale as dependências a partir da raiz do projeto:

```bash
pip install -r requirements.txt
```

## 1. Previsão de matrículas — `gerar_slides.py`

Apresentação fixa do caso de uso deste projeto (XGBoost + AWS).

```bash
python slides/gerar_slides.py                                 # pergunta no terminal
python slides/gerar_slides.py --slides 10 --modo resumido     # sem perguntas
```

Saída: `slides/previsao-matriculas.pptx`

## 2. Qualquer tema — `gerar_tema.py`

Uma IA escreve o conteúdo (textos, ícones, dados de gráficos e notas) e o script
monta os slides com o mesmo tema SENAI.

```bash
python slides/gerar_tema.py                                   # pergunta no terminal
python slides/gerar_tema.py --tema "Energia solar" --slides 10 --modo aprofundado
```

O script pergunta o **tema**, a **unidade curricular** (opcional), o **nível**
(resumido ou aprofundado) e **quantos slides** (3 a 30).

### Quem escreve o conteúdo

| Situação | O que acontece |
| --- | --- |
| `ANTHROPIC_API_KEY` definida | Usa a API do Claude (`claude-opus-5`) |
| Sem chave, com **Kiro CLI** (`kiro-cli`) instalado | Roda o agente do Kiro no terminal e monta o `.pptx` em seguida (requer `KIRO_API_KEY`) |
| Sem chave e sem Kiro CLI | Abre o chat do Kiro na IDE e aguarda o agente salvar o conteúdo |

Use `--kiro` para forçar o Kiro mesmo com a chave da API definida.

### Saída e reuso

As apresentações vão para `slides/apresentacoes/` (ignorada pelo git):

- `<tema>.pptx` — a apresentação
- `<tema>.json` — o conteúdo, editável; o esquema é validado ao renderizar
- `<tema>.prompt.md` — instruções entregues ao Kiro (quando ele é usado)

Para renderizar de novo depois de editar o JSON, sem nova chamada à IA:

```bash
python slides/gerar_tema.py --de-json slides/apresentacoes/<tema>.json
```

### Layouts disponíveis

`topicos` · `cartoes` · `numeros` · `processo` · `linha_do_tempo` · `comparacao` ·
`grafico` (barras, barras horizontais, linhas, rosca) · `destaque` · `tabela`

Números e gráficos sempre trazem a fonte; valores aproximados aparecem como
"Dados ilustrativos". Revise os dados antes de apresentar.

## Opções comuns

| Opção | Descrição |
| --- | --- |
| `--dpi N` | Resolução dos gráficos (padrão 300) |
| `--saida caminho.pptx` | Caminho do arquivo gerado |
| `--sem-instalar-fontes` | Não instala a Open Sans no Windows do usuário |

Na primeira execução no Windows, a Open Sans é instalada só para o usuário atual
(sem administrador), para o PowerPoint exibir a tipografia correta. As fontes e
ícones faltantes são baixados do Google Fonts automaticamente.
