"""
Gerador de slides PowerPoint — Previsão de Demanda de Vagas por Curso
Identidade visual: Modelo SENAI 2026 (assets/Modelo Senai 2026 - Indentidade Nova.pptx)
  - Tipografia: Open Sans (Google Fonts)
  - Ícones:     Material Symbols Rounded (Google Fonts)
  - Cores:      azul SENAI #164194 · laranja SENAI #E84910
  - Gráficos:   matplotlib renderizado no tamanho exato do slide, em HD (300 dpi)

Execute:
    python slides/gerar_slides.py                      (pergunta no terminal)
    python slides/gerar_slides.py --slides 10 --modo resumido
Saída:
    slides/previsao-matriculas.pptx
"""

import argparse
import io
import os
import shutil
import sys
import urllib.request
import zipfile
from dataclasses import dataclass, field

import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import Arc, Circle
from matplotlib.font_manager import fontManager
from PIL import Image, ImageDraw, ImageFont

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from lxml import etree


# ══════════════════════════════════════════════════════
# CAMINHOS
# ══════════════════════════════════════════════════════
BASE      = os.path.dirname(os.path.abspath(__file__))
ASSETS    = os.path.join(BASE, "assets")
FONT_DIR  = os.path.join(ASSETS, "fonts")
CACHE     = os.path.join(ASSETS, ".cache")
ICON_DIR  = os.path.join(CACHE, "icones")

LOGO_SENAI = os.path.join(ASSETS, "senai.png")
LOGO_CURSO = os.path.join(ASSETS, "técnico.png")
MODELO     = os.path.join(ASSETS, "Modelo Senai 2026 - Indentidade Nova.pptx")

# Google Fonts — baixadas automaticamente se não existirem em assets/fonts
_GF_OPENSANS = "https://github.com/googlefonts/opensans/raw/main/fonts/ttf/{}"
_GF_SYMBOLS  = ("https://github.com/google/material-design-icons/raw/master/variablefont/"
                "MaterialSymbolsRounded%5BFILL%2CGRAD%2Copsz%2Cwght%5D.{}")
FONTES_OPENSANS = ["OpenSans-Regular.ttf", "OpenSans-Italic.ttf", "OpenSans-Light.ttf",
                   "OpenSans-SemiBold.ttf", "OpenSans-Bold.ttf", "OpenSans-BoldItalic.ttf"]
SYMBOLS_TTF = os.path.join(FONT_DIR, "MaterialSymbolsRounded.ttf")
SYMBOLS_CP  = os.path.join(FONT_DIR, "MaterialSymbolsRounded.codepoints")

FONTE = "Open Sans"


# ══════════════════════════════════════════════════════
# PALETA SENAI
# ══════════════════════════════════════════════════════
def _rgb(h): return RGBColor.from_string(h.lstrip("#"))

H_AZUL      = "#164194"   # azul SENAI (títulos, painéis)
H_LARANJA   = "#E84910"   # laranja SENAI (acentos)
H_BRANCO    = "#FFFFFF"
H_TINTA     = "#1E2A44"   # texto de dados
H_CINZA     = "#5B6475"   # texto secundário
H_LINHA     = "#DDE3EE"   # grades, divisores
H_SUPERF    = "#F3F5FA"   # cartões
H_AZUL_BG   = "#E6ECF7"   # fundo de ícones
H_LAR_BG    = "#FDE9E1"
H_AZUL_TXT  = "#C9D6F0"   # texto secundário sobre azul
# Série categórica dos gráficos (validada p/ daltonismo, na ordem fixa)
H_SERIE_1   = "#7FA3E3"
H_SERIE_2   = "#2458B8"
H_SERIE_3   = H_LARANJA
# Status (reservado — sempre acompanhado de rótulo)
H_BOM       = "#2E9D5B"
H_ATENCAO   = "#F2B233"
H_CRITICO   = "#D64545"

AZUL, LARANJA, BRANCO = _rgb(H_AZUL), _rgb(H_LARANJA), _rgb(H_BRANCO)
CINZA, SUPERF, LINHA  = _rgb(H_CINZA), _rgb(H_SUPERF), _rgb(H_LINHA)
AZUL_BG, LAR_BG       = _rgb(H_AZUL_BG), _rgb(H_LAR_BG)
AZUL_TXT, TINTA       = _rgb(H_AZUL_TXT), _rgb(H_TINTA)

# Slide 16:9 (mesma proporção do modelo SENAI 20 x 11,25 in)
W, H = 13.333, 7.5
X0, X1 = 0.9, 12.58          # margens de conteúdo
Y0, Y1 = 2.0, 6.75           # área útil abaixo do cabeçalho

DPI = 300                    # resolução dos gráficos (sobrescrevível via --dpi)


# ══════════════════════════════════════════════════════
# RECURSOS: FONTES, ÍCONES, LOGOS
# ══════════════════════════════════════════════════════
def _baixar(url, destino):
    os.makedirs(os.path.dirname(destino), exist_ok=True)
    with urllib.request.urlopen(url, timeout=60) as r, open(destino, "wb") as f:
        shutil.copyfileobj(r, f)


def garantir_fontes():
    """Baixa Open Sans e Material Symbols do Google Fonts se faltarem."""
    for nome in FONTES_OPENSANS:
        destino = os.path.join(FONT_DIR, nome)
        if not os.path.exists(destino):
            print(f"  baixando {nome} (Google Fonts)...")
            _baixar(_GF_OPENSANS.format(nome), destino)
    for ext, destino in (("ttf", SYMBOLS_TTF), ("codepoints", SYMBOLS_CP)):
        if not os.path.exists(destino):
            print(f"  baixando Material Symbols .{ext} (Google Fonts)...")
            _baixar(_GF_SYMBOLS.format(ext), destino)
    for nome in FONTES_OPENSANS:
        fontManager.addfont(os.path.join(FONT_DIR, nome))


def instalar_fontes_usuario():
    """Instala Open Sans para o usuário atual no Windows (sem admin),
    para o PowerPoint exibir a tipografia correta. Retorna as fontes instaladas."""
    if sys.platform != "win32":
        return []
    import winreg
    sistema = os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts")
    usuario = os.path.join(os.environ["LOCALAPPDATA"], "Microsoft", "Windows", "Fonts")
    chave   = r"Software\Microsoft\Windows NT\CurrentVersion\Fonts"
    novas = []
    for nome in FONTES_OPENSANS:
        if os.path.exists(os.path.join(sistema, nome)) or os.path.exists(os.path.join(usuario, nome)):
            continue
        os.makedirs(usuario, exist_ok=True)
        destino = os.path.join(usuario, nome)
        shutil.copy2(os.path.join(FONT_DIR, nome), destino)
        familia, estilo = ImageFont.truetype(destino, 12).getname()
        rotulo = familia if estilo == "Regular" else f"{familia} {estilo}"
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, chave, 0, winreg.KEY_SET_VALUE) as k:
            winreg.SetValueEx(k, f"{rotulo} (TrueType)", 0, winreg.REG_SZ, destino)
        novas.append(nome)
    return novas


_codepoints = None
_fonte_icone = {}

def icone_png(nome, cor=H_AZUL, px=384):
    """Renderiza um Material Symbol como PNG transparente (supersampling 2x)."""
    global _codepoints
    caminho = os.path.join(ICON_DIR, f"{nome}_{cor.lstrip('#')}_{px}.png")
    if os.path.exists(caminho):
        return caminho
    if _codepoints is None:
        with open(SYMBOLS_CP, encoding="utf-8") as f:
            _codepoints = dict(l.split() for l in f if l.strip())
    s = px * 2
    if s not in _fonte_icone:
        fnt = ImageFont.truetype(SYMBOLS_TTF, s)
        fnt.set_variation_by_axes([0, 0, 48, 400])     # FILL, GRAD, opsz, wght
        _fonte_icone[s] = fnt
    im = Image.new("RGBA", (s, s), (0, 0, 0, 0))
    ImageDraw.Draw(im).text((s / 2, s / 2), chr(int(_codepoints[nome], 16)),
                            font=_fonte_icone[s], fill=cor, anchor="mm")
    os.makedirs(ICON_DIR, exist_ok=True)
    im.resize((px, px), Image.LANCZOS).save(caminho)
    return caminho


def _logo_recortado():
    """Remove a margem transparente do logo do curso."""
    destino = os.path.join(CACHE, "tecnico_recortado.png")
    if not os.path.exists(destino):
        im = Image.open(LOGO_CURSO).convert("RGBA")
        im = im.crop(im.getchannel("A").getbbox())
        im.thumbnail((1600, 1600), Image.LANCZOS)
        os.makedirs(CACHE, exist_ok=True)
        im.save(destino)
    return destino


def _logo_branco():
    """Versão monocromática branca do logo SENAI (fundos azuis do modelo)."""
    destino = os.path.join(CACHE, "senai_branco.png")
    if not os.path.exists(destino):
        im = Image.open(LOGO_SENAI).convert("RGBA")
        branco = Image.new("RGBA", im.size, (255, 255, 255, 0))
        branco.putalpha(im.getchannel("A"))
        os.makedirs(CACHE, exist_ok=True)
        branco.save(destino)
    return destino


def _fundo_azul():
    """Fundo azul SENAI com a colagem de fotos do modelo a 10% (como no original)."""
    destino = os.path.join(CACHE, "fundo_azul_08.jpg")
    if os.path.exists(destino):
        return destino
    base = Image.new("RGB", (1920, 1080), H_AZUL)
    if os.path.exists(MODELO):
        with zipfile.ZipFile(MODELO) as z:
            midias = [i for i in z.infolist()
                      if i.filename.startswith("ppt/media/") and i.filename.endswith(".png")]
            if midias:
                maior = max(midias, key=lambda i: i.file_size)
                foto = Image.open(io.BytesIO(z.read(maior))).convert("RGB").resize(base.size)
                if foto.size == base.size:
                    base = Image.blend(base, foto, 0.08)
    os.makedirs(CACHE, exist_ok=True)
    base.save(destino, quality=92)
    return destino


# ══════════════════════════════════════════════════════
# HELPERS PPTX  (todas as medidas em polegadas)
# ══════════════════════════════════════════════════════
I = Inches

def set_bg(sl, cor):
    f = sl.background.fill; f.solid(); f.fore_color.rgb = cor


def rect(sl, x, y, w, h, fill=None, line=None, lw=1.0, raio=None, forma=None):
    tipo = forma or (MSO_SHAPE.ROUNDED_RECTANGLE if raio else MSO_SHAPE.RECTANGLE)
    s = sl.shapes.add_shape(tipo, I(x), I(y), I(w), I(h))
    if raio:
        s.adjustments[0] = raio
    if fill is not None:
        s.fill.solid(); s.fill.fore_color.rgb = fill
    else:
        s.fill.background()
    if line is not None:
        s.line.color.rgb = line; s.line.width = Pt(lw)
    else:
        s.line.fill.background()
    s.shadow.inherit = False
    return s


def circulo(sl, x, y, d, fill):
    return rect(sl, x, y, d, d, fill=fill, forma=MSO_SHAPE.OVAL)


def poligono(sl, pontos, fill):
    fb = sl.shapes.build_freeform(I(pontos[0][0]), I(pontos[0][1]), scale=1.0)
    fb.add_line_segments([(I(px), I(py)) for px, py in pontos[1:]], close=True)
    s = fb.convert_to_shape()
    s.fill.solid(); s.fill.fore_color.rgb = fill
    s.line.fill.background(); s.shadow.inherit = False
    return s


_ALIGN  = {"l": PP_ALIGN.LEFT, "c": PP_ALIGN.CENTER, "r": PP_ALIGN.RIGHT}
_ANCHOR = {"t": MSO_ANCHOR.TOP, "m": MSO_ANCHOR.MIDDLE, "b": MSO_ANCHOR.BOTTOM}

def txt(sl, texto, x, y, w, h, size=16, bold=False, italic=False, cor=AZUL,
        align="l", anchor="t", espaco=1.1, depois=0):
    """Caixa de texto Open Sans sem margens internas; '\\n' separa parágrafos."""
    tb = sl.shapes.add_textbox(I(x), I(y), I(w), I(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = _ANCHOR[anchor]
    for i, linha in enumerate(texto.split("\n")):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = _ALIGN[align]
        p.line_spacing = espaco
        if depois:
            p.space_after = Pt(depois)
        r = p.add_run()
        r.text = linha
        f = r.font
        f.name, f.size, f.bold, f.italic = FONTE, Pt(size), bold, italic
        f.color.rgb = cor
    return tb


def icone(sl, nome, x, y, tam, cor=H_AZUL):
    return sl.shapes.add_picture(icone_png(nome, cor), I(x), I(y), I(tam), I(tam))


def icone_badge(sl, nome, x, y, d, fundo=AZUL_BG, cor=H_AZUL):
    """Ícone Google dentro de um círculo suave."""
    circulo(sl, x, y, d, fundo)
    t = d * 0.58
    icone(sl, nome, x + (d - t) / 2, y + (d - t) / 2, t, cor)


def pill(sl, rotulo, x, y, w, h, fundo=AZUL_BG, cor=AZUL, size=12):
    s = rect(sl, x, y, w, h, fill=fundo, raio=0.5)
    tf = s.text_frame
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = rotulo
    r.font.name, r.font.size, r.font.bold = FONTE, Pt(size), True
    r.font.color.rgb = cor
    return s


def numero(sl, n, x, y, d, fundo=AZUL):
    s = circulo(sl, x, y, d, fundo)
    tf = s.text_frame
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = str(n)
    r.font.name, r.font.size, r.font.bold = FONTE, Pt(d * 28), True
    r.font.color.rgb = BRANCO
    return s


# ── Transição e animação (fluidez na apresentação) ──
_NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"

def _transicao(sl):
    sl._element.append(etree.fromstring(
        f'<p:transition xmlns:p="{_NS_P}" spd="med"><p:fade/></p:transition>'))


def _animar_entrada(sl, ids):
    """Gráficos surgem com fade automático logo após a transição do slide."""
    if not ids:
        return
    efeitos, n = [], 5
    for i, spid in enumerate(ids):
        tipo = "afterEffect" if i == 0 else "withEffect"
        efeitos.append(
            f'<p:par><p:cTn id="{n}" presetID="10" presetClass="entr" presetSubtype="0" '
            f'fill="hold" grpId="0" nodeType="{tipo}"><p:stCondLst><p:cond delay="{i * 150}"/></p:stCondLst>'
            f'<p:childTnLst>'
            f'<p:set><p:cBhvr><p:cTn id="{n + 1}" dur="1" fill="hold"><p:stCondLst>'
            f'<p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl>'
            f'<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr>'
            f'<p:to><p:strVal val="visible"/></p:to></p:set>'
            f'<p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="{n + 2}" dur="700"/>'
            f'<p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl></p:cBhvr></p:animEffect>'
            f'</p:childTnLst></p:cTn></p:par>')
        n += 3
    sl._element.append(etree.fromstring(
        f'<p:timing xmlns:p="{_NS_P}"><p:tnLst><p:par>'
        f'<p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>'
        f'<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq">'
        f'<p:childTnLst><p:par><p:cTn id="3" fill="hold"><p:stCondLst><p:cond delay="indefinite"/>'
        f'<p:cond evt="onBegin" delay="0"><p:tn val="2"/></p:cond></p:stCondLst><p:childTnLst>'
        f'<p:par><p:cTn id="4" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst>'
        f'<p:childTnLst>{"".join(efeitos)}</p:childTnLst></p:cTn></p:par>'
        f'</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn>'
        f'<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>'
        f'<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst>'
        f'</p:seq></p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>'))


# ══════════════════════════════════════════════════════
# CONTEXTO DE GERAÇÃO
# ══════════════════════════════════════════════════════
@dataclass
class Ctx:
    prs: Presentation
    modo: str                       # "resumido" | "aprofundado"
    total: int
    atual: int = 0
    agenda: list = field(default_factory=list)
    graficos: list = field(default_factory=list)   # ids a animar no slide atual

    @property
    def fundo(self):
        return self.modo == "aprofundado"

    def v(self, resumido, aprofundado):
        """Escolhe o conteúdo conforme o nível de detalhe."""
        return aprofundado if self.fundo else resumido


def novo_slide(ctx, notas_r="", notas_a=""):
    ctx.atual += 1
    ctx.graficos = []
    sl = ctx.prs.slides.add_slide(ctx.prs.slide_layouts[6])
    set_bg(sl, BRANCO)
    notas = ctx.v(notas_r, notas_a or notas_r)
    if notas:
        sl.notes_slide.notes_text_frame.text = notas
    return sl


def finalizar(ctx, sl):
    _transicao(sl)
    _animar_entrada(sl, ctx.graficos)


def pagina(ctx, sl, cor=AZUL):
    txt(sl, f"Página {ctx.atual} de {ctx.total}", X1 - 3, 6.95, 3, 0.3,
        size=10.5, italic=True, cor=cor, align="r")


def barra_lateral(sl):
    """Faixa vertical do modelo SENAI: topo laranja, corte diagonal cinza, corpo azul."""
    b = 0.5
    poligono(sl, [(0, 0), (b, 0), (b, 0.85), (0, 1.2)], LARANJA)
    poligono(sl, [(0, 1.2), (b, 0.85), (b, 1.45), (0, 1.8)], _rgb("#CDCDCD"))
    poligono(sl, [(0, 1.8), (b, 1.45), (b, H), (0, H)], AZUL)


def cabecalho(ctx, sl, titulo, subtitulo=""):
    barra_lateral(sl)
    sl.shapes.add_picture(LOGO_SENAI, I(X1 - 1.6), I(0.62), I(1.6))
    txt(sl, titulo, X0, 0.55, 9.6, 0.75, size=32, bold=True, anchor="m")
    if subtitulo:
        txt(sl, subtitulo, X0, 1.3, 9.6, 0.4, size=15, cor=CINZA)
    pagina(ctx, sl)


def grafico(ctx, sl, fn, x, y, w, h, *args):
    """Renderiza o gráfico no tamanho físico exato da área (sem distorção)."""
    pic = sl.shapes.add_picture(fn(w, h, *args), I(x), I(y), I(w), I(h))
    ctx.graficos.append(pic.shape_id)
    return pic


def cartao(sl, x, y, w, h, fill=SUPERF):
    return rect(sl, x, y, w, h, fill=fill, raio=min(0.12, 0.12 / max(min(w, h), 0.5)))


# ══════════════════════════════════════════════════════
# GRÁFICOS (matplotlib · Open Sans · HD)
# ══════════════════════════════════════════════════════
def _mpl_setup():
    plt.rcParams.update({
        "font.family":        FONTE,
        "font.size":          11,
        "axes.labelsize":     11,
        "xtick.labelsize":    11,
        "ytick.labelsize":    10.5,
        "axes.labelcolor":    H_CINZA,
        "xtick.color":        H_CINZA,
        "ytick.color":        H_CINZA,
        "axes.spines.top":    False,
        "axes.spines.right":  False,
        "axes.spines.left":   False,
        "axes.edgecolor":     H_LINHA,
        "axes.linewidth":     1.0,
        "axes.facecolor":     H_BRANCO,
        "figure.facecolor":   H_BRANCO,
        "grid.color":         "#EDF0F6",
        "grid.linewidth":     0.9,
        "xtick.major.size":   0,
        "ytick.major.size":   0,
        "legend.frameon":     False,
        "legend.fontsize":    10.5,
        "lines.antialiased":  True,
        "patch.antialiased":  True,
        "lines.solid_capstyle": "round",
    })


def _fig(w, h, **kw):
    _mpl_setup()
    return plt.subplots(figsize=(w, h), layout="constrained", **kw)


def _png(fig):
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=DPI, facecolor=fig.get_facecolor())
    plt.close(fig)
    buf.seek(0)
    return buf


def _suave(x, y, pontos=240):
    """Interpolação cúbica monótona (Fritsch–Carlson): curvas fluidas sem
    ultrapassar os dados reais. Ignora trechos com NaN."""
    x, y = np.asarray(x, float), np.asarray(y, float)
    ok = ~np.isnan(y)
    x, y = x[ok], y[ok]
    if len(x) < 3:
        return x, y
    dx, dy = np.diff(x), np.diff(y)
    m = dy / dx
    t = np.zeros_like(y)
    t[0], t[-1] = m[0], m[-1]
    for i in range(1, len(x) - 1):
        t[i] = 0 if m[i - 1] * m[i] <= 0 else (
            3 * (dx[i - 1] + dx[i]) /
            ((2 * dx[i] + dx[i - 1]) / m[i - 1] + (dx[i] + 2 * dx[i - 1]) / m[i]))
    xs = np.linspace(x[0], x[-1], pontos)
    k = np.clip(np.searchsorted(x, xs) - 1, 0, len(x) - 2)
    hh = dx[k]; s = (xs - x[k]) / hh
    h00 = 2*s**3 - 3*s**2 + 1; h10 = s**3 - 2*s**2 + s
    h01 = -2*s**3 + 3*s**2;    h11 = s**3 - s**2
    return xs, h00*y[k] + h10*hh*t[k] + h01*y[k+1] + h11*hh*t[k+1]


def g_historico(w, h):
    cursos = ["ADS", "Redes", "Dev. Sist.", "Infra", "Ciberseg."]
    dados  = np.array([[210, 195, 230], [180, 175, 190],
                       [155, 170, 200], [130, 145, 160], [90, 110, 135]])
    anos   = ["2022", "2023", "2024"]
    cores  = [H_SERIE_1, H_SERIE_2, H_SERIE_3]
    x = np.arange(len(cursos)); bw = 0.26

    fig, ax = _fig(w, h)
    for i, (a, c) in enumerate(zip(anos, cores)):
        bs = ax.bar(x + (i - 1) * bw, dados[:, i], bw, label=a, color=c,
                    edgecolor=H_BRANCO, linewidth=1.5, zorder=3)
        ax.bar_label(bs, padding=3, fontsize=9.5, color=H_TINTA, fontweight="semibold")
    ax.set_xticks(x, cursos)
    ax.set_ylabel("Matrículas")
    ax.set_ylim(0, dados.max() * 1.18)
    ax.yaxis.grid(True); ax.set_axisbelow(True)
    ax.legend(loc="upper right", ncols=3, handlelength=1.1, columnspacing=1.2)
    return _png(fig)


def g_previsao(w, h):
    sems = ["2022/1", "2022/2", "2023/1", "2023/2", "2024/1", "2024/2", "2025/1", "2025/2"]
    x    = np.arange(len(sems), dtype=float)
    real = np.array([210, 195, 230, 218, 245, 238, np.nan, np.nan])
    prev = np.array([np.nan] * 5 + [242, 260, 275])
    ic_s = np.array([np.nan] * 5 + [255, 278, 295])
    ic_i = np.array([np.nan] * 5 + [229, 242, 255])

    fig, ax = _fig(w, h)
    ax.axvspan(5.5, 7.4, color=H_SUPERF, zorder=0)
    xb, sup = _suave(x, ic_s); _, inf = _suave(x, ic_i)
    ax.fill_between(xb, inf, sup, color=H_LARANJA, alpha=0.12, lw=0,
                    label="Intervalo de confiança 90%", zorder=1)
    ax.plot(*_suave(x, real), color=H_SERIE_2, lw=2.4, label="Real", zorder=3)
    ax.plot(*_suave(x, prev), color=H_LARANJA, lw=2.4, ls=(0, (5, 3)),
            label="Previsto (XGBoost)", zorder=3)
    ax.scatter(x, real, s=42, color=H_SERIE_2, edgecolor=H_BRANCO, lw=1.6, zorder=4)
    ax.scatter(x, prev, s=42, color=H_LARANJA, edgecolor=H_BRANCO, lw=1.6, zorder=4)
    for xi, v, dy in ((5, 238, -17), (7, 275, 9)):
        ax.annotate(f"{v}", (xi, v), xytext=(0, dy), textcoords="offset points",
                    ha="center", fontsize=10, color=H_TINTA, fontweight="semibold")
    ax.text(6.45, 196, "previsão", ha="center", fontsize=10, color=H_CINZA, style="italic")
    ax.set_xticks(x, sems)
    ax.set_xlim(-0.4, 7.4); ax.set_ylim(185, 305)
    ax.set_ylabel("Matrículas — ADS")
    ax.yaxis.grid(True); ax.set_axisbelow(True)
    ax.legend(loc="upper left")
    return _png(fig)


def g_donut(w, h):
    sizes, labels = [58, 27, 15], ["Presencial", "EAD", "Híbrido"]
    fig, ax = _fig(w, h, facecolor=H_SUPERF)
    ax.set_facecolor(H_SUPERF)
    wedges, _ = ax.pie(sizes, colors=[H_SERIE_2, H_SERIE_3, H_SERIE_1], startangle=90,
                       counterclock=False,
                       wedgeprops=dict(width=0.34, edgecolor=H_SUPERF, linewidth=3))
    for wd, lab, v in zip(wedges, labels, sizes):
        ang = np.deg2rad((wd.theta1 + wd.theta2) / 2)
        cx, cy = np.cos(ang) * 1.36, np.sin(ang) * 1.3
        ax.text(cx, cy, f"{v}%\n{lab}", ha="center", va="center",
                fontsize=10.5, color=H_TINTA, fontweight="semibold", linespacing=1.15)
    ax.text(0, 0.06, "58%", ha="center", va="center", fontsize=22,
            fontweight="bold", color=H_AZUL)
    ax.text(0, -0.24, "presencial", ha="center", va="center", fontsize=10, color=H_CINZA)
    ax.set_xlim(-1.75, 1.75); ax.set_ylim(-1.6, 1.6); ax.set_aspect("equal")
    return _png(fig)


def g_radar(w, h):
    cats = ["Histórico\nde demanda", "Taxa de\nevasão", "Vagas\nanteriores",
            "Nota de\ncorte", "IDH", "Modalidade"]
    vals = [0.92, 0.78, 0.85, 0.70, 0.60, 0.55]
    ang  = np.linspace(0, 2 * np.pi, len(cats), endpoint=False)
    ang_f, vals_f = np.append(ang, ang[0]), vals + [vals[0]]

    fig, ax = _fig(w, h, subplot_kw=dict(polar=True))
    ax.set_theta_offset(np.pi / 2); ax.set_theta_direction(-1)
    ax.fill(ang_f, vals_f, color=H_SERIE_2, alpha=0.14, zorder=2)
    ax.plot(ang_f, vals_f, color=H_SERIE_2, lw=2.2, zorder=3, solid_joinstyle="round")
    ax.scatter(ang, vals, s=60, color=H_SERIE_2, edgecolor=H_BRANCO, lw=2, zorder=4)
    for a, v in zip(ang, vals):
        ax.text(a, v + 0.13, f"{v:.0%}", ha="center", va="center",
                fontsize=9.5, color=H_TINTA, fontweight="semibold")
    ax.set_xticks(ang, cats, fontsize=10.5, color=H_TINTA)
    ax.tick_params(axis="x", pad=14)
    ax.set_ylim(0, 1.08)
    ax.set_yticks([0.25, 0.5, 0.75, 1.0], [])
    ax.yaxis.grid(color=H_LINHA); ax.xaxis.grid(color=H_LINHA)
    ax.spines["polar"].set_visible(False)
    return _png(fig)


def g_gauge(w, h, valor=11.3):
    fig, ax = _fig(w, h)
    ax.set_xlim(-1.35, 1.35); ax.set_ylim(-0.95, 1.15); ax.set_aspect("equal"); ax.axis("off")
    r, esp = 1.0, 22
    ax.add_patch(Arc((0, 0), 2 * r, 2 * r, theta1=0, theta2=180, color="#EDF0F6", lw=esp))
    faixas = [(0, 15, H_BOM), (15, 22, H_ATENCAO), (22, 30, H_CRITICO)]
    for a, b, c in faixas:
        ax.add_patch(Arc((0, 0), 2 * r, 2 * r, theta1=180 - b / 30 * 180,
                         theta2=180 - a / 30 * 180, color=c, lw=esp, alpha=0.9))
    rad = np.radians(180 - valor / 30 * 180)
    ax.plot([0, 0.72 * np.cos(rad)], [0, 0.72 * np.sin(rad)], color=H_TINTA, lw=3.2,
            solid_capstyle="round", zorder=5)
    ax.add_patch(Circle((0, 0), 0.07, color=H_TINTA, zorder=6))
    ax.text(-1.0, -0.2, "0%", ha="center", fontsize=9.5, color=H_CINZA)
    ax.text(1.0, -0.2, "30%", ha="center", fontsize=9.5, color=H_CINZA)
    ax.text(0, -0.36, f"{valor:.1f}%", ha="center", va="center", fontsize=24,
            fontweight="bold", color=H_AZUL)
    ax.text(0, -0.62, "MAPE", ha="center", va="center", fontsize=10.5, color=H_CINZA)
    ax.text(0, -0.86, "Dentro da meta (limite 15%)", ha="center", va="center",
            fontsize=10, color=H_BOM, fontweight="bold")
    return _png(fig)


def g_kpi_h(w, h):
    labels = ["Acurácia (MAPE < 15%)", "Cursos atendidos", "Uptime do endpoint",
              "Redução de custo ocioso"]
    vals = [92, 88, 99.5, 35]
    fig, ax = _fig(w, h)
    y = np.arange(len(labels))
    ax.barh(y, [100] * len(vals), color="#EDF0F6", height=0.56, zorder=2)
    ax.barh(y, vals, color=H_SERIE_2, height=0.56, zorder=3)
    for yi, v in zip(y, vals):
        ax.text(101.5, yi, f"{v:g}%", va="center", fontsize=10.5,
                color=H_TINTA, fontweight="bold")
    ax.set_yticks(y, labels, color=H_TINTA)
    ax.set_xlim(0, 110); ax.invert_yaxis()
    ax.set_xticks([]); ax.spines["bottom"].set_visible(False)
    return _png(fig)


def g_drift(w, h):
    sem = np.arange(1, 27)
    drift = np.array([.02, .03, .025, .04, .035, .05, .048, .06, .07, .065, .08, .09,
                      .085, .095, .11, .10, .105, .12, .09, .07, .06, .055, .05, .045, .04, .038])
    lim = 0.10
    xs, ys = _suave(sem, drift, 500)

    fig, ax = _fig(w, h)
    ax.fill_between(xs, ys, lim, where=ys > lim, color=H_CRITICO, alpha=0.14, lw=0,
                    interpolate=True)
    ax.plot(xs, ys, color=H_SERIE_2, lw=2.2, label="PSI (drift)", zorder=3)
    ax.axhline(lim, color=H_CRITICO, lw=1.6, ls=(0, (5, 3)), label="Limite 0,10", zorder=2)
    for s in (14, 18):
        ax.axvline(s, color=H_LARANJA, lw=1.5, ls=":", zorder=2)
        ax.text(s + 0.3, 0.128, "retreino", fontsize=9.5, color=H_LARANJA,
                style="italic", fontweight="semibold")
    ax.set_xlabel("Semana"); ax.set_ylabel("PSI")
    ax.set_xlim(1, 26); ax.set_ylim(0, 0.14)
    ax.yaxis.grid(True); ax.set_axisbelow(True)
    ax.legend(loc="upper left", ncols=2)
    return _png(fig)


# ══════════════════════════════════════════════════════
# SLIDES
# ══════════════════════════════════════════════════════
def slide_capa(ctx):
    sl = novo_slide(ctx,
        "Apresentação do projeto de previsão de demanda de vagas por curso.",
        "Apresente o objetivo: usar dados históricos de matrículas e um modelo XGBoost "
        "na AWS para antecipar a demanda de vagas por curso em até dois semestres.")
    sl.shapes.add_picture(_logo_recortado(), I(0.75), I(0.6), height=I(1.7))
    sl.shapes.add_picture(LOGO_SENAI, I(W - 0.75 - 2.2), I(0.75), I(2.2))

    rect(sl, 0.75, 3.0, 1.1, 0.08, fill=LARANJA)
    txt(sl, "Previsão de Demanda de\nVagas por Curso", 0.75, 3.2, 11.8, 1.8,
        size=50, bold=True, espaco=0.95)
    txt(sl, "UC: Arquitetura de Soluções com AWS", 0.75, 5.15, 11.8, 0.5,
        size=22, bold=True, italic=True)
    x = 0.75
    for tag in ["AWS", "XGBoost", "SageMaker", "Machine Learning"]:
        larg = 0.4 + 0.1 * len(tag)
        pill(sl, tag, x, 5.85, larg, 0.38)
        x += larg + 0.15

    rect(sl, 0, H - 0.12, W, 0.12, fill=AZUL)
    rect(sl, 0, H - 0.12, 2.6, 0.12, fill=LARANJA)
    finalizar(ctx, sl)


def slide_agenda(ctx):
    sl = novo_slide(ctx, "Roteiro da apresentação.",
                    "Mostre o roteiro e quanto tempo pretende gastar em cada bloco.")
    cabecalho(ctx, sl, "Agenda", "O que vamos ver hoje")
    itens = ctx.agenda
    colunas = 1 if len(itens) <= 5 else 2
    por_col = -(-len(itens) // colunas)
    cw = (X1 - X0 - 0.6 * (colunas - 1)) / colunas
    rh = min(0.82, (Y1 - Y0) / por_col)
    for i, titulo in enumerate(itens):
        c, r = divmod(i, por_col)
        x, y = X0 + c * (cw + 0.6), Y0 + r * rh
        txt(sl, f"{i + 1:02d}", x, y, 0.8, rh - 0.12, size=24, bold=True,
            cor=LARANJA, anchor="m")
        txt(sl, titulo, x + 0.9, y, cw - 0.9, rh - 0.12, size=17, bold=True, anchor="m")
        rect(sl, x, y + rh - 0.06, cw, 0.012, fill=LINHA)
    finalizar(ctx, sl)


def slide_divisoria(ctx, n, titulo, descricao):
    sl = novo_slide(ctx, f"Início da seção {n}: {titulo}.")
    sl.shapes.add_picture(_fundo_azul(), 0, 0, I(W), I(H))
    rect(sl, 1.27, 1.35, 1.0, 0.09, fill=LARANJA)
    txt(sl, "SEÇÃO", 1.27, 1.55, 4, 0.4, size=16, bold=True, cor=AZUL_TXT)
    txt(sl, f"{n:02d}", 1.15, 1.85, 6, 2.3, size=120, bold=True, cor=BRANCO, anchor="m")
    txt(sl, titulo, 1.27, 4.35, 10.5, 0.85, size=44, bold=True, cor=BRANCO)
    txt(sl, descricao, 1.27, 5.25, 10.5, 0.5, size=18, cor=AZUL_TXT)
    sl.shapes.add_picture(_logo_branco(), I(X1 - 1.9), I(0.7), I(1.9))
    pagina(ctx, sl, BRANCO)
    finalizar(ctx, sl)


def slide_problema(ctx):
    sl = novo_slide(ctx,
        "Quatro dores da abertura de vagas sem dados.",
        "Explore cada dor com um exemplo real da instituição: turma cancelada, lista de "
        "espera, decisão tomada em reunião sem base histórica e calendário apertado.")
    cabecalho(ctx, sl, "Qual é o problema?",
              "Instituições erram ao abrir vagas — e pagam caro por isso")
    itens = [
        ("money_off", "Salas vazias", "Vagas abertas sem alunos",
         "Turmas abertas sem demanda real geram custo com docentes, salas e laboratórios ociosos."),
        ("hourglass_top", "Fila de espera", "Alunos sem vaga no curso desejado",
         "Cursos procurados esgotam cedo e o aluno migra para outra instituição."),
        ("casino", "Puro achismo", "Intuição em vez de dados",
         "A oferta é definida por percepção da coordenação, sem série histórica estruturada."),
        ("schedule", "Gestão reativa", "Sem antecedência por semestre",
         "A demanda só é percebida no período de matrícula, tarde demais para ajustar a oferta."),
    ]
    cw = (X1 - X0 - 3 * 0.3) / 4
    ch = ctx.v(3.0, 3.75)
    for i, (ic, tit, curto, longo) in enumerate(itens):
        x = X0 + i * (cw + 0.3)
        cartao(sl, x, Y0, cw, ch)
        laranja = i % 2 == 1
        icone_badge(sl, ic, x + 0.3, Y0 + 0.32, 0.9,
                    LAR_BG if laranja else AZUL_BG, H_LARANJA if laranja else H_AZUL)
        txt(sl, tit, x + 0.3, Y0 + 1.45, cw - 0.5, 0.45, size=18, bold=True)
        txt(sl, ctx.v(curto, longo), x + 0.3, Y0 + 2.0, cw - 0.5, ch - 2.1,
            size=ctx.v(14, 13), cor=CINZA, espaco=1.2)

    by = Y0 + ch + 0.35
    rect(sl, X0, by, X1 - X0, 0.7, fill=AZUL, raio=0.18)
    icone(sl, "lightbulb", X0 + 0.3, by + 0.16, 0.38, H_BRANCO)
    txt(sl, "Solução: XGBoost treinado com histórico real — prevê a demanda 2 semestres à frente",
        X0 + 0.85, by, X1 - X0 - 1.1, 0.7, size=15, bold=True, cor=BRANCO, anchor="m")
    finalizar(ctx, sl)


def slide_solucao(ctx):
    sl = novo_slide(ctx,
        "Visão geral do pipeline de ML na AWS.",
        "Percorra o fluxo de cima para baixo: os dados chegam no S3, são tratados no Glue, "
        "viram features no SageMaker Feature Store, alimentam o treino do XGBoost e "
        "a previsão chega ao gestor pelo QuickSight.")
    cabecalho(ctx, sl, "Como resolvemos", "Pipeline de Machine Learning ponta a ponta na AWS")

    pw, ph = 4.6, Y1 - Y0
    rect(sl, X0, Y0, pw, ph, fill=AZUL, raio=0.05)
    txt(sl, "XGBoost", X0 + 0.4, Y0 + 0.35, pw - 0.8, 0.7, size=32, bold=True, cor=BRANCO)
    rect(sl, X0 + 0.4, Y0 + 1.1, 1.2, 0.07, fill=LARANJA)
    pontos = ctx.v(
        ["Treinado com histórico real", "Prevê 2 semestres à frente", "MAPE abaixo de 15%"],
        ["Treinado com histórico real", "Prevê 2 semestres à frente",
         "Retreino automático semestral", "Erro médio (MAPE) < 15%",
         "Dashboard ao vivo para gestores"])
    passo = (ph - 1.55) / len(pontos)
    for i, p in enumerate(pontos):
        y = Y0 + 1.45 + i * passo
        icone(sl, "check_circle", X0 + 0.4, y + 0.02, 0.32, H_BRANCO)
        txt(sl, p, X0 + 0.9, y, pw - 1.1, passo - 0.1, size=14.5, cor=BRANCO)

    etapas = [
        ("cloud_upload", "Ingestão de dados", "Amazon S3 (raw)", "CSVs de matrículas"),
        ("cleaning_services", "ETL e limpeza", "AWS Glue", "Padroniza e trata nulos"),
        ("extension", "Feature engineering", "SageMaker Feature Store", "Variáveis versionadas"),
        ("neurology", "Treinamento XGBoost", "Amazon SageMaker", "Treino e validação"),
        ("monitoring", "Previsão e dashboard", "Amazon QuickSight", "Painel de demanda por curso"),
    ]
    sx, sw = X0 + pw + 0.4, X1 - (X0 + pw + 0.4)
    rh, gap = 0.78, (ph - 5 * 0.78) / 4
    for i, (ic, nome, serv, desc) in enumerate(etapas):
        y = Y0 + i * (rh + gap)
        final = i == len(etapas) - 1
        if not final:
            rect(sl, sx + 0.46, y + rh, 0.025, gap, fill=LINHA)
        cartao(sl, sx, y, sw, rh)
        icone_badge(sl, ic, sx + 0.18, y + 0.11, 0.56,
                    LAR_BG if final else AZUL_BG, H_LARANJA if final else H_AZUL)
        txt(sl, nome, sx + 0.95, y + 0.1, 3.0, 0.35, size=15, bold=True)
        txt(sl, ctx.v(serv, f"{serv} · {desc}"), sx + 0.95, y + 0.43, sw - 1.7, 0.3,
            size=11.5, cor=CINZA)
        numero(sl, i + 1, sx + sw - 0.6, y + 0.19, 0.4, LARANJA if final else AZUL)
    finalizar(ctx, sl)


def slide_historico(ctx):
    sl = novo_slide(ctx,
        "Base histórica usada pelo modelo.",
        "Destaque o crescimento de ADS em 2024 e a tendência de alta em Ciberseg.: "
        "é essa série de 3 anos que o XGBoost aprende. Comente que Redes é o curso mais estável.")
    cabecalho(ctx, sl, "Histórico de matrículas", "5 cursos · 3 anos de dados · base para o modelo")
    grafico(ctx, sl, g_historico, X0, Y0 - 0.1, 8.35, Y1 - Y0 + 0.1)
    kpis = [("trending_up", "+23%", ctx.v("crescimento em ADS", "crescimento de ADS entre 2022 e 2024")),
            ("school", "5", ctx.v("cursos monitorados", "cursos técnicos monitorados")),
            ("calendar_month", "3 anos", ctx.v("de histórico", "de histórico semestral analisado"))]
    kx, kw = X0 + 8.6, X1 - (X0 + 8.6)
    kh = (Y1 - Y0 - 2 * 0.22) / 3
    for i, (ic, v, l) in enumerate(kpis):
        y = Y0 + i * (kh + 0.22)
        cartao(sl, kx, y, kw, kh)
        icone(sl, ic, kx + kw - 0.62, y + 0.22, 0.4, H_LARANJA if i == 0 else H_AZUL)
        txt(sl, v, kx + 0.28, y + 0.18, kw - 1.0, 0.65, size=30, bold=True,
            cor=LARANJA if i == 0 else AZUL)
        txt(sl, l, kx + 0.28, y + 0.85, kw - 0.5, kh - 0.9, size=12.5, cor=CINZA)
    finalizar(ctx, sl)


def slide_previsao(ctx):
    sl = novo_slide(ctx,
        "Comparação entre o real e o previsto.",
        "Mostre que a curva prevista acompanha a real até 2024/2 e explique a faixa laranja: "
        "o intervalo de confiança de 90%. Os semestres de 2025 são projeções do modelo.")
    cabecalho(ctx, sl, "Previsão vs. real", "Modelo XGBoost — curso de ADS, intervalo de confiança de 90%")
    grafico(ctx, sl, g_previsao, X0, Y0 - 0.1, 7.9, 4.45)
    txt(sl, "* 2025/1 e 2025/2 são previsões geradas pelo modelo", X0, 6.45, 7.9, 0.3,
        size=10.5, italic=True, cor=CINZA)
    dx = X0 + 8.15
    cartao(sl, dx, Y0, X1 - dx, Y1 - Y0)
    txt(sl, "Distribuição por modalidade", dx + 0.25, Y0 + 0.2, X1 - dx - 0.5, 0.4,
        size=14, bold=True)
    grafico(ctx, sl, g_donut, dx + 0.15, Y0 + 0.6, X1 - dx - 0.3, 3.2)
    if ctx.fundo:
        txt(sl, "A modalidade de ensino entra como variável do modelo e muda o "
                "perfil de demanda de cada curso.",
            dx + 0.25, Y0 + 3.85, X1 - dx - 0.5, 0.85, size=11.5, cor=CINZA, espaco=1.2)
    finalizar(ctx, sl)


def slide_features(ctx):
    sl = novo_slide(ctx,
        "Variáveis de entrada do modelo.",
        "Explique que histórico de demanda e vagas anteriores são as variáveis mais fortes, "
        "e que fatores externos como IDH e nota de corte refinam a previsão por região.")
    cabecalho(ctx, sl, "Features do modelo", "9 variáveis de entrada e sua importância relativa")
    feats = [("curso_id", "Identificador único do curso"),
             ("periodo_letivo", "Semestre e ano"),
             ("historico_demanda", "Média móvel de 3 anos"),
             ("taxa_evasao", "Evasão histórica (%)"),
             ("vagas_anteriores", "Vagas do semestre anterior"),
             ("nota_corte_enem", "Nota de corte no ENEM"),
             ("regiao_geografica", "Macrorregião"),
             ("modalidade_ensino", "Presencial, EAD ou híbrido"),
             ("idh_municipio", "IDH do município")]
    cw, gx = 3.2, 0.2
    rh = ctx.v(0.64, 0.82); gy = (Y1 - Y0 - 5 * rh) / 4
    for i, (f, d) in enumerate(feats):
        c, r = i % 2, i // 2
        x, y = X0 + c * (cw + gx), Y0 + r * (rh + gy)
        cartao(sl, x, y, cw, rh)
        rect(sl, x, y + 0.14, 0.06, rh - 0.28, fill=AZUL if (r + c) % 2 == 0 else LARANJA)
        if ctx.fundo:
            txt(sl, f, x + 0.25, y + 0.1, cw - 0.35, 0.35, size=13.5, bold=True)
            txt(sl, d, x + 0.25, y + 0.45, cw - 0.35, 0.3, size=11, cor=CINZA)
        else:
            txt(sl, f, x + 0.25, y, cw - 0.35, rh, size=13.5, bold=True, anchor="m")
    grafico(ctx, sl, g_radar, X0 + 6.75, Y0 - 0.2, X1 - (X0 + 6.75), Y1 - Y0 + 0.2)
    finalizar(ctx, sl)


def slide_arquitetura(ctx):
    sl = novo_slide(ctx,
        "As seis camadas da arquitetura AWS.",
        "Mostre que as camadas azuis cuidam do dado e do modelo, e as laranjas da operação: "
        "monitorar, alertar e entregar a previsão para os gestores.")
    cabecalho(ctx, sl, "Arquitetura AWS", "Do dado bruto à decisão — todas as camadas em um olhar")
    camadas = [
        ("database", "Ingestão", "S3 Raw", "S3 Raw · bucket prevmatriculas-dev-raw recebe os CSVs"),
        ("cleaning_services", "ETL", "Glue ETL · S3 Processed", "Glue ETL limpa e grava no S3 Processed"),
        ("extension", "Feature Store", "SageMaker Feature Store", "SageMaker Feature Store · grupo prevmatriculas-features"),
        ("neurology", "ML Pipeline", "SageMaker Train · Endpoint REST", "SageMaker treina o XGBoost e publica um endpoint REST"),
        ("notifications_active", "Operação", "Model Monitor · CloudWatch · EventBridge", "Model Monitor, CloudWatch e EventBridge vigiam e retreinam"),
        ("dashboard", "Visualização", "QuickSight · gestores", "QuickSight entrega o painel aos gestores acadêmicos"),
    ]
    rh, gap = 0.66, (Y1 - Y0 - 6 * 0.66) / 5
    for i, (ic, nome, curto, longo) in enumerate(camadas):
        y = Y0 + i * (rh + gap)
        ops = i >= 4
        cartao(sl, X0, y, X1 - X0, rh)
        rect(sl, X0, y + 0.12, 0.06, rh - 0.24, fill=LARANJA if ops else AZUL)
        icone_badge(sl, ic, X0 + 0.25, y + 0.1, 0.46,
                    LAR_BG if ops else AZUL_BG, H_LARANJA if ops else H_AZUL)
        txt(sl, nome, X0 + 0.9, y, 2.6, rh, size=16, bold=True, anchor="m")
        txt(sl, ctx.v(curto, longo), X0 + 3.6, y, X1 - X0 - 3.9, rh, size=13.5,
            cor=CINZA, anchor="m")
    finalizar(ctx, sl)


def slide_pipeline(ctx):
    sl = novo_slide(ctx,
        "Fluxo de dados em seis passos.",
        "Siga o fluxo da esquerda para a direita e reforce o gatilho de retreino: "
        "se o MAPE passar de 15%, o EventBridge dispara um novo treinamento sem intervenção manual.")
    cabecalho(ctx, sl, "Fluxo de dados", "Do CSV ao dashboard em 6 passos")
    etapas = [("database", "S3 Raw", "Fontes brutas", "CSVs de matrículas chegam ao bucket raw"),
              ("cleaning_services", "Glue ETL", "Limpeza", "Limpeza, tipagem e junção das fontes"),
              ("extension", "Feature Store", "Features", "Variáveis calculadas e versionadas"),
              ("neurology", "SageMaker", "Treino XGBoost", "Treino com validação cruzada"),
              ("api", "Endpoint", "Inferência REST", "Previsões sob demanda via API"),
              ("dashboard", "QuickSight", "Dashboard", "Painel para os gestores")]
    n, gap = len(etapas), 0.3
    nw = (X1 - X0 - (n - 1) * gap) / n
    ny, nh = Y0 + 0.3, 2.05
    for i, (ic, nome, curto, longo) in enumerate(etapas):
        x = X0 + i * (nw + gap)
        laranja = i % 2 == 1
        cartao(sl, x, ny, nw, nh)
        icone_badge(sl, ic, x + (nw - 0.8) / 2, ny + 0.32, 0.8,
                    LAR_BG if laranja else AZUL_BG, H_LARANJA if laranja else H_AZUL)
        txt(sl, nome, x + 0.08, ny + 1.3, nw - 0.16, 0.55, size=14.5, bold=True,
            align="c", anchor="m")
        numero(sl, i + 1, x + nw - 0.32, ny - 0.14, 0.4, LARANJA if laranja else AZUL)
        txt(sl, ctx.v(curto, longo), x, ny + nh + 0.15, nw, 0.9, size=ctx.v(12.5, 11.5),
            cor=CINZA, align="c", espaco=1.15)
        if i < n - 1:
            icone(sl, "arrow_forward", x + nw + (gap - 0.26) / 2, ny + nh / 2 - 0.13,
                  0.26, H_CINZA)
    by = Y1 - 0.75
    cartao(sl, X0, by, X1 - X0, 0.75)
    rect(sl, X0, by + 0.14, 0.06, 0.47, fill=LARANJA)
    icone(sl, "bolt", X0 + 0.3, by + 0.18, 0.4, H_LARANJA)
    txt(sl, "EventBridge dispara retreinamento automático quando o MAPE passa de 15%",
        X0 + 0.9, by, X1 - X0 - 1.2, 0.75, size=15, bold=True, anchor="m")
    finalizar(ctx, sl)


def slide_modelo(ctx):
    sl = novo_slide(ctx,
        "Por que escolhemos XGBoost.",
        "Compare com as alternativas: Prophet capta sazonalidade, LSTM exige muito mais dados "
        "e ARIMA serve de baseline. O XGBoost equilibra precisão, velocidade e interpretabilidade.")
    cabecalho(ctx, sl, "O modelo: XGBoost", "Rápido, preciso e interpretável")
    pw = 4.3
    rect(sl, X0, Y0, pw, Y1 - Y0, fill=AZUL, raio=0.05)
    txt(sl, "Parâmetros", X0 + 0.4, Y0 + 0.35, pw - 0.8, 0.6, size=26, bold=True, cor=BRANCO)
    rect(sl, X0 + 0.4, Y0 + 1.0, 1.2, 0.07, fill=LARANJA)
    params = ctx.v(
        [("Horizonte", "2 semestres"), ("Erro máximo", "MAPE < 15%"), ("Retreino", "A cada 180 dias")],
        [("Horizonte", "2 semestres à frente"), ("Erro máximo", "MAPE < 15%"),
         ("Métricas", "RMSE · MAE"), ("Instância AWS", "ml.m5.xlarge"),
         ("Retreinamento", "A cada 180 dias")])
    passo = (Y1 - Y0 - 1.4) / len(params)
    for i, (k, v) in enumerate(params):
        y = Y0 + 1.3 + i * passo
        txt(sl, k, X0 + 0.4, y, pw - 0.8, 0.28, size=11.5, cor=AZUL_TXT)
        txt(sl, v, X0 + 0.4, y + 0.26, pw - 0.8, 0.36, size=16, bold=True, cor=BRANCO)

    gx = X0 + pw + 0.35
    grafico(ctx, sl, g_gauge, gx, Y0 - 0.05, 3.7, 2.55)
    tx = gx + 3.9
    txt(sl, "Por que XGBoost?", tx, Y0 + 0.1, X1 - tx, 0.4, size=16, bold=True)
    motivos = ctx.v(["Treina em minutos", "Explica cada previsão"],
                    ["Treina em minutos", "Ótimo com dados tabulares",
                     "Features explicáveis"])
    for i, m in enumerate(motivos):
        y = Y0 + 0.65 + i * 0.52
        icone(sl, "check_circle", tx, y + 0.02, 0.28, H_LARANJA)
        txt(sl, m, tx + 0.4, y, X1 - tx - 0.4, 0.5, size=12.5, cor=CINZA)

    txt(sl, "Alternativas avaliadas", gx, Y0 + 2.7, X1 - gx, 0.4, size=16, bold=True)
    alts = [("Prophet", "Sazonalidade acadêmica semestral"),
            ("LSTM", "Padrões complexos de longo prazo"),
            ("ARIMA", "Baseline estatístico de comparação")]
    aw = (X1 - gx - 2 * 0.25) / 3
    ay, ah = Y0 + 3.2, Y1 - (Y0 + 3.2)
    for i, (nm, d) in enumerate(alts):
        x = gx + i * (aw + 0.25)
        cartao(sl, x, ay, aw, ah)
        txt(sl, nm, x + 0.25, ay + 0.2, aw - 0.5, 0.4, size=15, bold=True)
        txt(sl, d, x + 0.25, ay + 0.65, aw - 0.5, ah - 0.75, size=12, cor=CINZA, espaco=1.15)
    finalizar(ctx, sl)


def slide_monitoramento(ctx):
    sl = novo_slide(ctx,
        "Como o modelo é vigiado em produção.",
        "Aponte as semanas 15 e 18, quando o PSI passou do limite de 0,10 e o retreino foi "
        "disparado. Em seguida mostre que os KPIs de produção seguem dentro da meta.")
    cabecalho(ctx, sl, "Monitoramento 24/7", "O modelo fica em produção e precisa ser vigiado")
    grafico(ctx, sl, g_drift, X0, Y0 - 0.1, 7.85, 3.35)
    servs = [("query_stats", "Model Monitor", "Detecta drift (limite 0,10)"),
             ("notifications_active", "CloudWatch", "Alarmes de MAPE e latência"),
             ("event_repeat", "EventBridge", "Retreino a cada 180 dias"),
             ("functions", "Lambda", "Pré-processamento Python 3.12")]
    sx = X0 + 8.05
    cw, chh = (X1 - sx - 0.2) / 2, (3.25 - 0.2) / 2
    for i, (ic, nm, d) in enumerate(servs):
        c, r = i % 2, i // 2
        x, y = sx + c * (cw + 0.2), Y0 + r * (chh + 0.2)
        cartao(sl, x, y, cw, chh)
        icone(sl, ic, x + 0.2, y + 0.2, 0.38, H_LARANJA if i >= 2 else H_AZUL)
        txt(sl, nm, x + 0.2, y + 0.68, cw - 0.3, 0.3, size=13, bold=True)
        if ctx.fundo:
            txt(sl, d, x + 0.2, y + 0.98, cw - 0.3, chh - 1.0, size=10.5, cor=CINZA)
    txt(sl, "KPIs em produção", X0, Y0 + 3.35, 4, 0.35, size=14, bold=True)
    grafico(ctx, sl, g_kpi_h, X0, Y0 + 3.65, X1 - X0, Y1 - (Y0 + 3.65))
    finalizar(ctx, sl)


def slide_resultados(ctx):
    sl = novo_slide(ctx,
        "O que o modelo entrega.",
        "Feche com os números-chave e, se houver perguntas técnicas, mostre os padrões de "
        "nomenclatura adotados nos recursos AWS.")
    cabecalho(ctx, sl, "Resultados esperados", "O que o modelo entrega na prática")
    kpis = [("target", "< 15%", "Erro médio (MAPE)"),
            ("date_range", "2 sem.", "Horizonte de previsão"),
            ("event_repeat", "180 dias", "Ciclo de retreino"),
            ("verified", "90%", "Intervalo de confiança")]
    kw = (X1 - X0 - 3 * 0.3) / 4
    kh = ctx.v(2.45, 1.75)
    for i, (ic, v, l) in enumerate(kpis):
        x = X0 + i * (kw + 0.3)
        destaque = i == 0
        cartao(sl, x, Y0, kw, kh)
        icone_badge(sl, ic, x + 0.25, Y0 + 0.25, ctx.v(0.75, 0.55),
                    LAR_BG if destaque else AZUL_BG, H_LARANJA if destaque else H_AZUL)
        txt(sl, v, x + 0.25, Y0 + ctx.v(1.1, 0.8), kw - 0.5, ctx.v(0.75, 0.6), size=ctx.v(34, 28),
            bold=True, cor=LARANJA if destaque else AZUL)
        txt(sl, l, x + 0.25, Y0 + ctx.v(1.85, 1.35), kw - 0.5, 0.35, size=12.5, cor=CINZA)

    if not ctx.fundo:
        by = Y0 + kh + 0.5
        rect(sl, X0, by, X1 - X0, 1.3, fill=AZUL, raio=0.1)
        icone(sl, "insights", X0 + 0.45, by + 0.4, 0.5, H_BRANCO)
        txt(sl, "Abertura de vagas baseada em dados,\ncom dois semestres de antecedência.",
            X0 + 1.3, by, X1 - X0 - 1.6, 1.3, size=20, bold=True, cor=BRANCO, anchor="m")
    else:
        ty = Y0 + kh + 0.3
        txt(sl, "Padrões de nomenclatura AWS", X0, ty, 8, 0.4, size=16, bold=True)
        padroes = [("S3 Buckets", "prevmatriculas-{env}-{tipo}", "prevmatriculas-prod-raw"),
                   ("Glue Jobs", "job-prevmatriculas-{etapa}", "job-prevmatriculas-feature-engineering"),
                   ("Modelos SageMaker", "modelo-prevmatriculas-v{versao}", "modelo-prevmatriculas-v1"),
                   ("Feature Group", "prevmatriculas-features", "SageMaker Feature Store")]
        ry = ty + 0.5
        rh = (Y1 - ry - 3 * 0.06) / 4
        for i, (t, p, ex) in enumerate(padroes):
            y = ry + i * (rh + 0.06)
            cartao(sl, X0, y, X1 - X0, rh)
            txt(sl, t, X0 + 0.25, y, 2.6, rh, size=12.5, bold=True, anchor="m")
            txt(sl, p, X0 + 3.0, y, 4.3, rh, size=12, cor=TINTA, anchor="m")
            txt(sl, ex, X0 + 7.4, y, X1 - X0 - 7.6, rh, size=11, italic=True, cor=CINZA, anchor="m")
    finalizar(ctx, sl)


def slide_proximos(ctx):
    sl = novo_slide(ctx, "Próximos passos do projeto.",
        "Apresente o roteiro de evolução: começar pequeno com um piloto, medir, "
        "expandir para a rede e só então incorporar novas fontes e simulações.")
    cabecalho(ctx, sl, "Próximos passos", "Roteiro de evolução do projeto")
    fases = [("flag", "Piloto", "5 cursos técnicos com acompanhamento semestral do erro"),
             ("hub", "Expansão", "Todas as unidades da rede usando o mesmo pipeline"),
             ("dataset", "Novas fontes", "Dados socioeconômicos e de mercado de trabalho"),
             ("tune", "Simulação", "Painel de cenários para testar ofertas de vagas")]
    n = len(fases)
    cw = (X1 - X0 - (n - 1) * 0.3) / n
    ly = Y0 + 1.25
    rect(sl, X0 + cw / 2, ly - 0.015, (X1 - X0) - cw, 0.03, fill=LINHA)
    for i, (ic, t, d) in enumerate(fases):
        x = X0 + i * (cw + 0.3)
        cx = x + cw / 2
        txt(sl, f"FASE {i + 1}", x, Y0 + 0.3, cw, 0.35, size=12, bold=True, cor=LARANJA, align="c")
        circulo(sl, cx - 0.45, ly - 0.45, 0.9, BRANCO)
        icone_badge(sl, ic, cx - 0.4, ly - 0.4, 0.8,
                    LAR_BG if i == 0 else AZUL_BG, H_LARANJA if i == 0 else H_AZUL)
        cartao(sl, x, ly + 0.75, cw, 2.35)
        txt(sl, t, x + 0.25, ly + 0.95, cw - 0.5, 0.45, size=18, bold=True)
        txt(sl, d, x + 0.25, ly + 1.5, cw - 0.5, 1.6, size=13, cor=CINZA, espaco=1.2)
    finalizar(ctx, sl)


def slide_encerramento(ctx):
    sl = novo_slide(ctx, "Encerramento e contatos.",
                    "Agradeça, abra para perguntas e deixe os contatos visíveis.")
    sl.shapes.add_picture(_fundo_azul(), 0, 0, I(W), I(H))
    sl.shapes.add_picture(_logo_branco(), I(1.3), I(3.05), I(4.6))
    rect(sl, W / 2, 1.95, 0.02, 3.6, fill=BRANCO)
    cx = W / 2 + 0.55
    txt(sl, "Obrigado!", cx, 1.75, 5.5, 0.8, size=40, bold=True, cor=BRANCO)
    contatos = [("location_on", "Rodovia Admar Gonzaga, 2765 | Itacorubi\n88034-001 | Florianópolis - SC", False),
                ("language", "sc.senai.br", True),
                ("mail", "faleconosco@fiesc.com.br", False),
                ("call", "(48) 3231 4100  ·  0800 048 1212", False)]
    y = 2.8
    for ic, t, destaque in contatos:
        linhas = t.count("\n") + 1
        icone(sl, ic, cx, y + 0.02, 0.32, H_LARANJA if destaque else H_BRANCO)
        txt(sl, t, cx + 0.5, y, 5.3, 0.33 * linhas, size=14, bold=destaque,
            italic=destaque, cor=BRANCO, espaco=1.15)
        y += 0.33 * linhas + 0.22
    pagina(ctx, sl, BRANCO)
    finalizar(ctx, sl)


# ══════════════════════════════════════════════════════
# ROTEIRO: quais slides entram, conforme quantidade e nível
# ══════════════════════════════════════════════════════
@dataclass
class Item:
    chave: str
    titulo: str
    secao: str | None
    prioridade: int
    fn: object
    so_aprofundado: bool = False


SECOES = {
    "contexto":   (1, "Contexto", "O problema e a proposta de solução"),
    "dados":      (2, "Dados e modelo", "Histórico, previsões, features e o XGBoost"),
    "aws":        (3, "Arquitetura AWS", "Camadas, fluxo de dados e monitoramento"),
    "resultados": (4, "Resultados", "O que o modelo entrega"),
}

# Ordem de exibição; prioridade menor = entra primeiro quando há poucos slides
CONTEUDO = [
    Item("problema",      "Qual é o problema?",      "contexto",   1, slide_problema),
    Item("solucao",       "Como resolvemos",         "contexto",   2, slide_solucao),
    Item("historico",     "Histórico de matrículas", "dados",      5, slide_historico),
    Item("previsao",      "Previsão vs. real",       "dados",      3, slide_previsao),
    Item("features",      "Features do modelo",      "dados",     10, slide_features),
    Item("modelo",        "O modelo: XGBoost",       "dados",      7, slide_modelo),
    Item("arquitetura",   "Arquitetura AWS",         "aws",        4, slide_arquitetura),
    Item("pipeline",      "Fluxo de dados",          "aws",        8, slide_pipeline),
    Item("monitoramento", "Monitoramento 24/7",      "aws",        9, slide_monitoramento),
    Item("resultados",    "Resultados esperados",    "resultados", 6, slide_resultados),
    Item("proximos",      "Próximos passos",         "resultados", 11, slide_proximos,
         so_aprofundado=True),
]


def limites(modo):
    conteudo = [c for c in CONTEUDO if modo == "aprofundado" or not c.so_aprofundado]
    return 3, 2 + len(conteudo) + 1 + len(SECOES)   # capa + encerramento + agenda + divisórias


def montar_roteiro(n, modo):
    """Retorna a lista ordenada de (tipo, dado) com exatamente n slides."""
    disponiveis = [c for c in CONTEUDO if modo == "aprofundado" or not c.so_aprofundado]
    resto = n - 2
    escolhidos = set(c.chave for c in sorted(disponiveis, key=lambda c: c.prioridade)[:resto])
    resto -= len(escolhidos)
    com_agenda = resto > 0
    resto -= int(com_agenda)
    selecionados = [c for c in disponiveis if c.chave in escolhidos]
    # Divisórias para as seções mais longas primeiro
    por_secao = {}
    for c in selecionados:
        por_secao.setdefault(c.secao, []).append(c)
    com_divisoria = set(sorted(por_secao, key=lambda s: (-len(por_secao[s]), SECOES[s][0]))[:max(resto, 0)])

    roteiro = [("capa", None)]
    if com_agenda:
        roteiro.append(("agenda", None))
    secao_atual = None
    for c in selecionados:
        if c.secao != secao_atual:
            secao_atual = c.secao
            if c.secao in com_divisoria:
                roteiro.append(("divisoria", c.secao))
        roteiro.append(("conteudo", c))
    roteiro.append(("encerramento", None))
    return roteiro


# ══════════════════════════════════════════════════════
# CLI
# ══════════════════════════════════════════════════════
def _perguntar(msg, padrao):
    try:
        r = input(msg).strip()
    except EOFError:
        r = ""
    return r or str(padrao)


def perguntar_modo():
    print("\n  Nível de detalhe")
    print("    1) Resumido     — frases curtas, foco nos visuais")
    print("    2) Aprofundado  — explicações, detalhes técnicos e notas do apresentador")
    while True:
        r = _perguntar("  Escolha [1]: ", 1).lower()
        if r in ("1", "r", "resumido"):
            return "resumido"
        if r in ("2", "a", "aprofundado"):
            return "aprofundado"
        print("  Opção inválida. Digite 1 ou 2.")


def perguntar_slides(modo):
    mn, mx = limites(modo)
    padrao = min(12, mx)
    while True:
        r = _perguntar(f"\n  Quantos slides? ({mn} a {mx}) [{padrao}]: ", padrao)
        if r.isdigit() and mn <= int(r) <= mx:
            return int(r)
        print(f"  Digite um número inteiro entre {mn} e {mx}.")


def _progresso(i, total, rotulo):
    larg = 28
    cheio = round(larg * i / total)
    barra = "█" * cheio + "░" * (larg - cheio)
    sys.stdout.write(f"\r  {barra}  {i:>2}/{total}  {rotulo:<28}")
    sys.stdout.flush()


def main():
    global DPI
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except AttributeError:
        pass

    ap = argparse.ArgumentParser(description="Gera a apresentação SENAI de previsão de matrículas.")
    ap.add_argument("--slides", type=int, help="quantidade de slides")
    ap.add_argument("--modo", choices=["resumido", "aprofundado"], help="nível de detalhe")
    ap.add_argument("--dpi", type=int, default=DPI, help="resolução dos gráficos (padrão 300)")
    ap.add_argument("--saida", default=os.path.join(BASE, "previsao-matriculas.pptx"))
    ap.add_argument("--sem-instalar-fontes", action="store_true",
                    help="não instala Open Sans no Windows do usuário")
    args = ap.parse_args()
    DPI = args.dpi

    print("\n  SENAI  |  Gerador de slides — Previsão de Demanda de Vagas por Curso")
    print("  " + "─" * 66)

    modo = args.modo or perguntar_modo()
    mn, mx = limites(modo)
    if args.slides is not None and not (mn <= args.slides <= mx):
        ap.error(f"--slides deve estar entre {mn} e {mx} no modo {modo}")
    n = args.slides or perguntar_slides(modo)

    print("\n  Preparando recursos (Google Fonts, ícones, logos)...")
    garantir_fontes()
    if not args.sem_instalar_fontes:
        novas = instalar_fontes_usuario()
        if novas:
            print(f"  Open Sans instalada para o usuário ({len(novas)} arquivos) — "
                  "reinicie o PowerPoint se estiver aberto.")

    roteiro = montar_roteiro(n, modo)
    prs = Presentation()
    prs.slide_width, prs.slide_height = I(W), I(H)
    ctx = Ctx(prs, modo, total=len(roteiro))
    ctx.agenda = [d.titulo for t, d in roteiro if t == "conteudo"]

    print(f"\n  Gerando {len(roteiro)} slides · modo {modo} · gráficos em {DPI} dpi\n")
    n_secao = 0
    for i, (tipo, dado) in enumerate(roteiro, 1):
        rotulo = {"capa": "Capa", "agenda": "Agenda", "encerramento": "Encerramento",
                  "divisoria": f"Seção: {SECOES[dado][1]}" if tipo == "divisoria" else "",
                  "conteudo": dado.titulo if tipo == "conteudo" else ""}[tipo]
        _progresso(i - 1, len(roteiro), rotulo)
        if tipo == "capa":
            slide_capa(ctx)
        elif tipo == "agenda":
            slide_agenda(ctx)
        elif tipo == "divisoria":
            n_secao += 1
            _, titulo, desc = SECOES[dado]
            slide_divisoria(ctx, n_secao, titulo, desc)
        elif tipo == "conteudo":
            dado.fn(ctx)
        else:
            slide_encerramento(ctx)
        _progresso(i, len(roteiro), rotulo)

    prs.save(args.saida)
    print(f"\n\n  Concluído. Arquivo salvo em:\n  {args.saida}\n")


if __name__ == "__main__":
    main()
