#!/usr/bin/env python3
"""
HUGO BOSS — baja las fotos de los productos de un excel de estructura.

Cubre las dos marcas del archivo: HUGO y BOSS. Cada una va a su carpeta y
tiene su propia hoja de revision.


QUE HACE ESTE PROGRAMA, EN CRIOLLO
==================================

Por cada producto del excel:

  1. arma la direccion de su ficha en hugoboss.com
  2. baja esa pagina y le saca la lista de fotos
  3. baja cada foto y la guarda con el codigo del excel como nombre

Nada mas. Los pasos 1 y 3 son especificos de esta marca y estan aca; leer
el excel, comparar fotos repetidas y armar la hoja de revision es igual
para todas las marcas y vive en comun.py.


1. LA DIRECCION DE LA FICHA
---------------------------
Se construye desde el excel, sin buscar nada:

    hugoboss.com/hbeu{MATERIAL}_{COD.COLOR}.html
                     \\____ las dos columnas del excel ____/

    material 50549243 + color 404  ->  hugoboss.com/hbeu50549243_404.html

Si esa pagina no existe, el producto no esta publicado.


2. LA LISTA DE FOTOS
--------------------
Estan en el HTML de la ficha, sin JavaScript de por medio, y en el orden
del carrusel. Se sacan con una busqueda de texto:

    images.hugoboss.com/is/image/boss/hbeu50549243_404_350?...
    images.hugoboss.com/is/image/boss/hbeu50549243_404_300?...
    images.hugoboss.com/is/image/boss/hbeu50549243_404_360?...
    ...

EL ORDEN SE RESPETA: los sufijos se toman en el orden en que aparecen en
el HTML, que es el del carrusel del sitio (para el Odeno2: 350, 300, 360,
340, 341, 100 — y el sitio marca "1/6"). No se reordenan por numero. La
foto que el sitio muestra primera es la que se guarda sin sufijo.

POR QUE ASI Y NO PROBANDO NUMEROS: la primera version de este script
adivinaba los sufijos probando de 10 en 10 del 0 al 500. Eran 51 pedidos
por producto y ademas fallaba, porque este producto tiene una foto con
sufijo 341, que no es multiplo de 10. Leer la ficha da la lista exacta,
en el orden correcto, con un solo pedido.


3. LA DESCARGA
--------------
Las fotos viven en Adobe Scene7, que las entrega al tamano que le pidas.
Se le pide con fit=constrain, que significa "que entre en esta caja pero
SIN agrandar". Como la caja es mas grande que cualquier original, lo que
devuelve es el archivo tal cual.

Medido sobre una foto real:

    wid=2000&fit=constrain   -> 1500 x 2275   <- el original
    wid=1600 (como el sitio) -> 1600 x 2427   <- agrandada un 7%
    hei=3000                 -> 1500 x 3000   <- estirada, todo relleno

O sea que la propia web entrega las fotos reescaladas. Nosotros bajamos
el original.


USO
---
    pip install pandas xlrd requests pillow

    python hugo.py Hugo.xls --limite 3     # probar con 3 productos
    python hugo.py Hugo.xls                # las dos marcas
    python hugo.py Hugo.xls --marca BOSS   # solo una
"""

import argparse
import re
from pathlib import Path

import pandas as pd
import requests

import comun


# ===========================================================================
# LO ESPECIFICO DE HUGO BOSS
# ===========================================================================

MARCAS = ["HUGO", "BOSS"]
CARPETAS = {"HUGO": "HUGO", "BOSS": "BOSS"}

# Prefijo del codigo web. Verificado en BOSS.
PREFIJO = "hbeu"

FICHA = "https://www.hugoboss.com/{prefijo}{material}_{color}.html"

# Para los productos que no aparecieron: el buscador del sitio con el codigo
# de barra del excel. En Hugo Boss el EAN cae directo en el producto, asi
# que es el mejor link para que el cliente compruebe por si mismo si el
# articulo existe o no.
BUSCADOR = "https://www.hugoboss.com/search?q={ean}"

IMG = ("https://images.hugoboss.com/is/image/boss/"
       "{prefijo}{material}_{color}_{sufijo}?wid={ancho}&fit=constrain")

# Caja para pedirle la foto a Scene7. Es un TECHO, no un objetivo: con
# fit=constrain el servidor no agranda, asi que devuelve el original.
# 2000 alcanza de sobra (los originales miden ~1500 de ancho) y si algun
# producto tuviera fotos mas grandes, tambien entran.
MAX_ANCHO = 2000

CABECERAS = {
    "User-Agent": ("Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
                   "(KHTML, like Gecko) Chrome/120.0 Safari/537.36"),
    "Accept-Language": "en-US,en;q=0.9",
}


def url_ficha(material, color, prefijo=PREFIJO):
    return FICHA.format(prefijo=prefijo, material=material, color=color)


def url_foto(material, color, sufijo, prefijo=PREFIJO, ancho=MAX_ANCHO):
    return IMG.format(prefijo=prefijo, material=material, color=color,
                      sufijo=sufijo, ancho=ancho)


def leer_ficha(material, color, sesion, prefijo=PREFIJO):
    """
    Baja la ficha y le saca dos cosas:

      sufijos : los numeros de foto de ESTE color, en orden de carrusel
      colores : los colores que el sitio publica de este material, sacados
                de los botoncitos de color (terminan en _SW). Sirve para
                explicar por que un producto no aparecio.

    Devuelve (sufijos, colores, encontrada).
    """
    try:
        r = sesion.get(url_ficha(material, color, prefijo),
                       headers=CABECERAS, timeout=30, allow_redirects=True)
    except requests.RequestException:
        return [], [], False
    if r.status_code != 200:
        return [], [], False

    html = r.text

    # Fotos de este color. OJO: no alcanza con buscar el codigo, porque el
    # mismo aparece antes en la etiqueta de compartir en redes:
    #
    #   <meta property="og:image" content=".../hbeu50549243_404_100?$social_sharing$">
    #
    # Esa esta ARRIBA de la galeria en el documento, asi que si se toma el
    # orden crudo, la _100 queda primera cuando la portada real es la _350.
    # Las de la galeria se reconocen por $re_fullPageZoom$ en la query.
    patron = re.compile(
        rf"/boss/{re.escape(prefijo)}{re.escape(str(material))}_"
        rf"{re.escape(str(color))}_(\d+)([^\"'\s>)]*)")

    galeria, otras = [], []
    for m in patron.finditer(html):
        sufijo, cola = m.group(1), m.group(2)
        if "social_sharing" in cola:
            continue
        (galeria if "re_fullPageZoom" in cola else otras).append(sufijo)

    def sin_repetir(lista):
        vistos, salida = set(), []
        for s in lista:
            if s not in vistos:
                vistos.add(s)
                salida.append(s)
        return salida

    # Si algun dia cambian el nombre del preset, se usa lo que haya.
    sufijos = sin_repetir(galeria) or sin_repetir(otras)

    # Colores publicados: hbeu50549243_001_SW
    swatch = re.compile(
        rf"{re.escape(prefijo)}{re.escape(str(material))}_(\d+)_SW")
    colores = sorted(set(swatch.findall(html)))

    return sufijos, colores, True


# ===========================================================================
# PRINCIPAL
# ===========================================================================

def main(argv=None):
    """argv=None usa la linea de comandos; una lista permite llamarlo
    desde codigo (lo hace runner.py cuando corre dentro del .exe)."""
    ap = argparse.ArgumentParser(description="Baja las fotos de HUGO y BOSS.")
    ap.add_argument("excel", type=Path)
    ap.add_argument("--salida", type=Path, default=Path("fotos"))
    ap.add_argument("--limite", type=int, default=None)
    ap.add_argument("--marca", default=None, help="HUGO o BOSS. Por defecto, las dos.")
    ap.add_argument("--prefijo", default=PREFIJO)
    ap.add_argument("--ancho", type=int, default=MAX_ANCHO,
                    help="techo de tamano; no agranda (por defecto 2000)")
    ap.add_argument("--forzar", action="store_true")
    args = ap.parse_args(argv)

    marcas = [args.marca.upper()] if args.marca else MARCAS
    sesion = requests.Session()
    hubo = False

    for marca in marcas:
        prod = comun.cargar(args.excel, marca=marca)
        if len(prod) == 0:
            continue
        hubo = True
        if args.limite:
            prod = prod.head(args.limite)

        carpeta = args.salida / CARPETAS.get(marca, marca)
        carpeta.mkdir(parents=True, exist_ok=True)
        print(f"\n{marca}: {len(prod)} productos -> {carpeta}/")
        print("un pedido a la ficha por producto\n")

        filas = []
        for i, (_, p) in enumerate(prod.iterrows(), 1):
            cod, mat, col = p["cod"], p["referencia"], p["cod_color"]

            if not args.forzar and comun.ya_bajado(carpeta, cod):
                existentes = comun.archivos_de(carpeta, cod)
                print(f"  {i:>3}/{len(prod)} [--] {cod:<16} ya estaba ({len(existentes)})")
                filas.append(_fila(p, existentes, 0, [], args))
                continue

            sufijos, colores, existe = leer_ficha(mat, col, sesion, args.prefijo)
            comun.time.sleep(comun.ESPERA)

            nota = ""
            if not existe:
                nota = "La ficha no existe en hugoboss.com. Producto no publicado."
            elif not sufijos:
                nota = (f"La ficha abre pero no tiene fotos del color {col}."
                        + (f" Colores publicados de este material: "
                           f"{', '.join(colores)}." if colores else ""))

            crudas = []
            for n, s in enumerate(sufijos):
                img = comun.bajar(url_foto(mat, col, s, args.prefijo, args.ancho),
                                  sesion)
                if img is not None:
                    crudas.append((n, img))          # n = orden del carrusel

            imagenes = comun.unicas(crudas)
            nombres = comun.guardar(imagenes, carpeta, cod)
            repetidas = len(crudas) - len(imagenes)

            detalle = f"{len(nombres)} fotos" if nombres else (nota or "sin fotos")
            if repetidas:
                detalle += f"  ({repetidas} repetidas descartadas)"
            print(f"  {i:>3}/{len(prod)} {'[ok]' if nombres else '[!!]'} {cod:<16} {detalle}")

            filas.append(_fila(p, nombres, repetidas, sufijos, args, nota))

        rep = pd.DataFrame(filas)
        destino, total, sin = comun.hoja_revision(rep, carpeta, carpeta.name)

        print(f"\n{'-' * 64}")
        print(f"  {marca}: {(rep.n_fotos > 0).sum()} de {len(rep)} con fotos, "
              f"{total} fotos, {rep.repetidas.sum()} repetidas")
        print(f"  Abri:  {destino.resolve()}")

    if not hubo:
        todas = comun.cargar(args.excel)["marca"].unique()
        print(f"El excel no tiene productos {marcas}. Marcas: {list(todas)}")


def _fila(p, nombres, repetidas, sufijos, args, nota=""):
    mat, col = p["referencia"], p["cod_color"]
    ean = str(p.get("ean", "") or "").strip()

    # Con fotos, el link va a la ficha. Sin fotos, la ficha no sirve de nada
    # (o no existe), asi que va al buscador con el codigo de barra.
    if nombres or not ean:
        url_producto = url_ficha(mat, col, args.prefijo)
        etiqueta = "Ver la ficha en hugoboss.com"
    else:
        url_producto = BUSCADOR.format(ean=ean)
        etiqueta = f"Buscar el codigo de barra {ean}"

    return {
        "cod": p["cod"],
        "nombre": p.get("nombre", ""),
        "color": p.get("color", ""),
        "referencia": mat,
        "cod_color": col,
        "tallas": p.get("tallas", ""),
        "n_fotos": len(nombres),
        "repetidas": repetidas,
        "archivos": " ".join(nombres),
        "sufijos": " ".join(sufijos),
        "nota": nota,
        "ean": ean,
        "url_producto": url_producto,
        "label_producto": etiqueta,
        "url_foto1": (url_foto(mat, col, sufijos[0], args.prefijo, args.ancho)
                      if sufijos else ""),
    }


if __name__ == "__main__":
    main()
