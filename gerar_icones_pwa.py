"""Gera os ícones do PWA (192 e 512 px) em static/ — só PIL, sem rede.

Visual: quadrado com cantos arredondados em gradiente índigo→ciano (mesmo
gradiente do app) e 3 barras brancas de gráfico + linha de tendência.
"""
from __future__ import annotations

from pathlib import Path

from PIL import Image, ImageDraw, ImageFilter

SAIDA = Path(__file__).resolve().parent / "static"
GRAD_INICIO = (99, 102, 241)  # #6366F1
GRAD_FIM = (34, 211, 238)  # #22D3EE


def _gradiente(tam: int) -> Image.Image:
    """Gradiente diagonal índigo→ciano num quadrado tam×tam."""
    base = Image.new("RGB", (tam, tam))
    px = base.load()
    for y in range(tam):
        for x in range(tam):
            f = (x + y) / (2 * tam)
            px[x, y] = tuple(
                int(GRAD_INICIO[i] + (GRAD_FIM[i] - GRAD_INICIO[i]) * f) for i in range(3)
            )
    return base


def gerar(tam: int) -> Image.Image:
    # 1) gradiente de fundo
    img = _gradiente(tam).convert("RGBA")

    # 2) arredonda os cantos (máscara com raio 22% do tamanho)
    raio = int(tam * 0.22)
    mascara = Image.new("L", (tam, tam), 0)
    ImageDraw.Draw(mascara).rounded_rectangle([0, 0, tam - 1, tam - 1], raio, fill=255)
    img.putalpha(mascara)

    # 3) barras brancas do gráfico (proporção fixa do tamanho)
    camada = Image.new("RGBA", (tam, tam), (0, 0, 0, 0))
    d = ImageDraw.Draw(camada)
    branco = (255, 255, 255, 255)
    larg = tam * 0.13  # largura de cada barra
    base_y = tam * 0.78  # linha do chão
    alturas = [0.24, 0.40, 0.56]  # 3 barras crescentes
    espaco = tam * 0.09
    total = 3 * larg + 2 * espaco
    x0 = (tam - total) / 2
    for i, alt in enumerate(alturas):
        x = x0 + i * (larg + espaco)
        y = base_y - tam * alt
        d.rounded_rectangle([x, y, x + larg, base_y], int(larg * 0.28), fill=branco)

    # 4) linha de tendência sobre as barras
    pontos = [
        (x0 - tam * 0.04, base_y - tam * 0.30),
        (x0 + total * 0.5, base_y - tam * 0.50),
        (x0 + total + tam * 0.04, base_y - tam * 0.68),
    ]
    d.line(pontos, fill=branco, width=max(2, int(tam * 0.035)), joint="curve")
    raio_pt = max(3, int(tam * 0.035))
    for cx, cy in pontos:
        d.ellipse([cx - raio_pt, cy - raio_pt, cx + raio_pt, cy + raio_pt], fill=branco)

    # leve sombra interna para dar profundidade nas barras
    camada = camada.filter(ImageFilter.GaussianBlur(0.4))
    img = Image.alpha_composite(img, camada)
    return img


def main() -> None:
    SAIDA.mkdir(exist_ok=True)
    for tam in (192, 512):
        caminho = SAIDA / f"icon-{tam}.png"
        gerar(tam).save(caminho, "PNG")
        print(f"OK {caminho} ({caminho.stat().st_size} bytes)")


if __name__ == "__main__":
    main()
