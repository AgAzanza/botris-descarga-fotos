#!/usr/bin/env python3
"""
GEOX — baja las fotos de los productos de un excel de estructura.


QUE HACE ESTE PROGRAMA, EN CRIOLLO
==================================

Por cada producto del excel:

  1. arma la direccion de su ficha en geox.com
  2. baja esa pagina y le saca la lista de fotos
  3. baja cada foto y la guarda con el codigo del excel como nombre

Leer el excel, comparar fotos repetidas y armar la hoja de revision es
igual para todas las marcas y vive en comun.py.


1. LA DIRECCION DE LA FICHA
---------------------------
El codigo del excel sin la talla ES la direccion. No hace falta el slug:

    geox.com/en-RU/D650ZB00022C4005.html

Geox redirige solo a la version larga con el nombre del producto. Si esa
pagina no existe, el articulo no esta publicado.


2. LA LISTA DE FOTOS
--------------------
Estan en el HTML, en atributos data-src (Geox las carga de forma diferida,
por eso no van en src):

    data-src="https://geox-cdn.thron.com/delivery/public/thumbnail/geox/
              EC_D650ZB00022C4005_100/crzcqo/std/1024x1024/EC_AB615_100.jpg"
                                \\__ el codigo del excel __/ \\_ tamano _/

OJO CON LA ETIQUETA DE REDES: el mismo codigo aparece antes en un
content="..." que es la miniatura para compartir. Si se toma el orden
crudo del documento, esa se cuela primero. Solo se leen los data-src.

POR QUE ASI Y NO PROBANDO NUMEROS: la primera version de este script
adivinaba los sufijos probando de 10 en 10 y explorando vecinos. Eran ~21
pedidos por producto y no habia forma de saber si se escapaba alguno: los
sufijos no son contiguos (se ven 00 10 30 40 50 60 100 en un producto y
90 91 en otros). Leer la ficha da la lista exacta, en el orden del
carrusel, con un solo pedido.


3. LA DESCARGA
--------------
La URL del CDN lleva el tamano adentro. La ficha pide 1024x1024 y nosotros
lo cambiamos por 2048x2048, que el CDN sirve sin problema.


USO
---
    pip install pandas xlrd requests pillow

    python geox.py Geox.xls --limite 3     # probar con 3 productos
    python geox.py Geox.xls                # bajar todo
    python geox.py Geox.xls --forzar       # rehacer los ya bajados

    # por que falta una foto de un producto:
    python geox.py Geox.xls --diagnostico D653KC000TUC1S9B
"""

import argparse
import re
from pathlib import Path

import pandas as pd
import requests

import comun


# ===========================================================================
# LO ESPECIFICO DE GEOX
# ===========================================================================

MARCA = "GEOX"
CARPETA = "Geox"

# Sufijo que llevan los archivos de imagen. Sale del formato que espera el
# archivo de carga de WooCommerce: {codigo}-{sufijo}.webp
SUFIJO = "geox-ecuador"

FICHA = "https://www.geox.com/{locale}/{cod}.html"
LOCALE = "en-RU"

# Tamano que se le pide al CDN. La ficha trae 1024x1024 en la URL y se
# reemplaza por este. Verificado a mano: 2048x2048 responde OK.
TAMANO = "2048x2048"

CABECERAS = {
    "User-Agent": ("Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
                   "(KHTML, like Gecko) Chrome/120.0 Safari/537.36"),
    "Accept-Language": "en-US,en;q=0.9",
}

# data-src="...EC_{codigo}_{sufijo}/..."   -> captura la URL entera
FOTO = re.compile(r'data-src="([^"]*geox-cdn\.thron\.com[^"]*)"')


def url_ficha(cod, locale=LOCALE):
    return FICHA.format(locale=locale, cod=cod)


def agrandar(url, tamano=TAMANO):
    """Cambia el /std/1024x1024/ de la URL por el tamano que queramos."""
    return re.sub(r"/std/\d+x\d+/", f"/std/{tamano}/", url)


def leer_ficha(cod, sesion, locale=LOCALE, tamano=TAMANO):
    """
    Baja la ficha y devuelve (urls_de_fotos, encontrada).

    Solo mira los data-src, que son los del carrusel, y respeta el orden en
    que aparecen. Descarta repetidas conservando la primera aparicion.
    """
    try:
        r = sesion.get(url_ficha(cod, locale), headers=CABECERAS,
                       timeout=30, allow_redirects=True)
    except requests.RequestException:
        return [], False
    if r.status_code != 200:
        return [], False

    urls, vistos = [], set()
    for url in FOTO.findall(r.text):
        # nos quedamos solo con las de ESTE producto
        m = re.search(rf"EC_{re.escape(cod)}_(\d+)", url)
        if not m:
            continue
        sufijo = m.group(1)
        if sufijo in vistos:
            continue
        vistos.add(sufijo)
        urls.append(agrandar(url, tamano))

    return urls, True


def sufijo_de(url):
    m = re.search(r"EC_[A-Z0-9]+_(\d+)/", url)
    return m.group(1) if m else "?"


def diagnostico(cod, sesion, locale=LOCALE, tamano=TAMANO):
    """
    Para un solo producto: que fotos trae la ficha, cuales bajaron y cuales
    se descartaron por parecidas. Sirve para responder "por que falta esta".
    """
    urls, existe = leer_ficha(cod, sesion, locale, tamano)
    print(f"\n  {cod}   ficha: {url_ficha(cod, locale)}")
    if not existe:
        print("  la ficha no responde\n")
        return
    print(f"  fotos en el HTML: {len(urls)}\n")

    bajadas = []
    for n, url in enumerate(urls):
        img = comun.bajar(url, sesion)
        marca = sufijo_de(url)
        if img is None:
            print(f"    _{marca:<4} NO SE PUDO BAJAR   {url}")
            continue
        bajadas.append((marca, img, comun.firma(img)))
        print(f"    _{marca:<4} {img.width}x{img.height}")

    print(f"\n  comparacion (se descarta si la diferencia es <= {comun.DIF_IGUAL}):")
    guardadas = []
    for marca, img, f in bajadas:
        gemela = None
        for gm, gf in guardadas:
            d = comun.diferencia(gf, f)
            if d <= comun.DIF_IGUAL:
                gemela = (gm, d)
                break
        if gemela:
            print(f"    _{marca:<4} DESCARTADA por parecida a _{gemela[0]} "
                  f"(diferencia {gemela[1]:.2f})")
        else:
            print(f"    _{marca:<4} se guarda")
            guardadas.append((marca, f))

    print(f"\n  resultado: {len(guardadas)} de {len(bajadas)} fotos\n")


# ===========================================================================
# PRINCIPAL
# ===========================================================================

def main(argv=None):
    """argv=None usa la linea de comandos; una lista permite llamarlo
    desde codigo (lo hace runner.py cuando corre dentro del .exe)."""
    ap = argparse.ArgumentParser(description="Baja las fotos de Geox.")
    ap.add_argument("excel", type=Path)
    ap.add_argument("--salida", type=Path, default=Path("fotos"))
    ap.add_argument("--limite", type=int, default=None)
    ap.add_argument("--locale", default=LOCALE)
    ap.add_argument("--tamano", default=TAMANO)
    ap.add_argument("--forzar", action="store_true")
    ap.add_argument("--diagnostico", metavar="CODIGO",
                    help="para un solo producto: que fotos trae la ficha, "
                         "cuales bajaron y cuales se descartaron por parecidas")
    args = ap.parse_args(argv)

    if args.diagnostico:
        diagnostico(args.diagnostico, requests.Session(), args.locale, args.tamano)
        return

    prod = comun.cargar(args.excel, marca=MARCA)
    if len(prod) == 0:
        todas = comun.cargar(args.excel)["marca"].unique()
        print(f"El excel no tiene productos {MARCA}. Marcas: {list(todas)}")
        return
    if args.limite:
        prod = prod.head(args.limite)

    carpeta = args.salida / CARPETA
    carpeta.mkdir(parents=True, exist_ok=True)
    print(f"\n{MARCA}: {len(prod)} productos -> {carpeta}/")
    print("un pedido a la ficha por producto\n")

    sesion = requests.Session()
    filas = []

    for i, (_, p) in enumerate(prod.iterrows(), 1):
        cod = p["cod"]

        if not args.forzar and comun.ya_bajado(carpeta, cod, SUFIJO):
            existentes = comun.archivos_de(carpeta, cod, SUFIJO)
            print(f"  {i:>3}/{len(prod)} [--] {cod:<20} ya estaba ({len(existentes)})")
            filas.append(_fila(p, existentes, 0, [], args))
            continue

        urls, existe = leer_ficha(cod, sesion, args.locale, args.tamano)
        comun.time.sleep(comun.ESPERA)

        nota = ""
        if not existe:
            nota = ("La ficha no existe en geox.com. Producto no publicado "
                    "o todavia no cargado.")
        elif not urls:
            nota = "La ficha abre pero no tiene fotos de este producto."

        crudas = []
        for n, url in enumerate(urls):
            img = comun.bajar(url, sesion)
            if img is not None:
                crudas.append((n, img))       # n = orden del carrusel

        imagenes = comun.unicas(crudas)
        nombres = comun.guardar(imagenes, carpeta, cod, SUFIJO)
        repetidas = len(crudas) - len(imagenes)

        detalle = f"{len(nombres)} fotos" if nombres else (nota or "sin fotos")
        if repetidas:
            detalle += f"  ({repetidas} repetidas descartadas)"
        print(f"  {i:>3}/{len(prod)} {'[ok]' if nombres else '[!!]'} {cod:<20} {detalle}")

        filas.append(_fila(p, nombres, repetidas, urls, args, nota))

    rep = pd.DataFrame(filas)
    destino, total, sin = comun.hoja_revision(rep, carpeta, CARPETA)
    xls, n_filas, n_prod = comun.excel_encontrados(args.excel, rep, carpeta,
                                                   marca=MARCA)

    print(f"\n{'-' * 64}")
    print(f"  productos con fotos  : {(rep.n_fotos > 0).sum()} de {len(rep)}")
    print(f"  fotos guardadas      : {total}")
    print(f"  repetidas descartadas: {rep.repetidas.sum()}")
    if xls:
        print(f"  excel de encontrados : {xls.name}  "
              f"({n_prod} productos, {n_filas} filas)")
    print(f"\n  Abri:  {destino.resolve()}")


def _fila(p, nombres, repetidas, urls, args, nota=""):
    return {
        "cod": p["cod"],
        "nombre": p.get("nombre", ""),
        "color": p.get("color", ""),
        "referencia": p.get("referencia", ""),
        "cod_color": p.get("cod_color", ""),
        "tallas": p.get("tallas", ""),
        "n_fotos": len(nombres),
        "repetidas": repetidas,
        "archivos": " ".join(nombres),
        "sufijos": " ".join(sufijo_de(u) for u in urls),
        "nota": nota,
        "url_producto": url_ficha(p["cod"], args.locale),
        "label_producto": "Ver la ficha en geox.com",
        "url_foto1": urls[0] if urls else "",
    }


if __name__ == "__main__":
    main()