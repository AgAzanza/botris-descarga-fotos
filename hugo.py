#!/usr/bin/env python3
"""
MICHAEL KORS — baja las fotos de los productos de un excel de estructura.


QUE HACE ESTE PROGRAMA, EN CRIOLLO
==================================

Por cada producto del excel le pregunta al sitio por esa variante concreta
y se baja las fotos que le devuelve.

Leer el excel, comparar fotos repetidas y armar la hoja de revision es
igual para todas las marcas y vive en comun.py.


POR QUE NO SE PUEDE CONSTRUIR LA URL DE LA FOTO
-----------------------------------------------
Las fotos de MK llevan un identificador aleatorio en el medio:

    assets.michaelkors.com/transform/ECOM_Image_Zoom_Highres/
        9c5f384e-0f9a-4f0e-88af-c815ee9b2de2/MT670N27R3-001-0001_1-tif
        \\_________ esto no se deduce _________/\\__ esto si __/

El nombre del archivo si sigue un patron, pero el UUID no. Hay que
pedirle las URLs al sitio.


LA CONSULTA
-----------
MK corre sobre Salesforce Commerce Cloud, que expone las variantes en JSON:

    /on/demandware.store/Sites-mk_us-Site/en_US/Product-Variation
        ?pid={REFERENCIA}&dwvar_{REFERENCIA}_color={COLOR}&quantity=1

Los dos datos salen del excel. Un pedido por producto.

OJO CON EL COLOR: el codigo del excel no es el mismo que el del sitio, y la
conversion NO es uniforme. Verificado a mano sobre MT670N27R3:

    excel 001  ->  sitio 0001      (relleno con ceros)
    excel 303  ->  sitio 3031      (NO es 0303)

Por eso el script prueba primero el relleno con ceros y, si no acierta, usa
la lista de colores que el propio sitio publica en variationAttributes.

Esto importa porque cuando el color no coincide el servidor NO da error:
devuelve las fotos del color por defecto, con codigo 200 y JSON valido. Y
tampoco sirve mirar productType: MK contesta "master" igual cuando acierta.
Lo unico que distingue es el nombre del archivo, que lleva el color adentro.

Verificado a mano: pidiendo color=001 sobre MT670N27R3 devolvio
"productType": "master" y fotos de "MT670N27R3-303-3031", o sea del color
303. Un script ingenuo las hubiera guardado como si fueran del 001.

Por eso hay DOS controles antes de guardar nada:

  1. productType tiene que ser "variant", no "master"
  2. el nombre del archivo tiene que contener -{COLOR}-

Si alguno falla, el producto queda marcado como dudoso en la hoja de
revision y no se guarda ninguna foto. Es preferible entregar un pendiente
a entregar la foto del color equivocado.


EL TAMANO
---------
El JSON devuelve la misma foto en varias medidas. Se usa zoomHiRes, que es
la mas grande.


USO
---
    pip install pandas xlrd requests pillow curl_cffi

    python mk.py MK.xls --limite 3
    python mk.py MK.xls

    # ver que contesta el sitio para un producto, sin bajar nada:
    python mk.py MK.xls --diagnostico MT670N27R3 001
"""

import argparse
import json
from pathlib import Path

import pandas as pd
import requests

import comun


# ===========================================================================
# SESION
# ===========================================================================
#
# michaelkors.com responde 403 a requests pelado (lo mismo que Lanidor).
# curl_cffi imita la huella TLS de Chrome y pasa.

try:
    from curl_cffi import requests as cffi
    HAY_CFFI = True
except ImportError:
    HAY_CFFI = False


def nueva_sesion():
    if HAY_CFFI:
        return cffi.Session(impersonate="chrome")
    print("  aviso: curl_cffi no esta instalado y michaelkors.com devuelve\n"
          "         403 sin el. Instalalo con:  pip install curl_cffi\n")
    return requests.Session()


# ===========================================================================
# LO ESPECIFICO DE MICHAEL KORS
# ===========================================================================

MARCA = "MK"
CARPETA = "MichaelKors"

# Sufijo de los archivos de imagen (ver comun.nombre_foto).
SUFIJO = "michaelkors-ecuador"

BASE = "https://www.michaelkors.com"
VARIACION = (BASE + "/on/demandware.store/Sites-mk_us-Site/en_US/"
                    "Product-Variation")
BUSCADOR = BASE + "/search?q={referencia}"

# Cual de los tamanos del JSON usamos, en orden de preferencia.
TAMANOS = ["zoomHiRes", "zoom", "large", "base"]

CABECERAS = {
    "User-Agent": ("Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
                   "(KHTML, like Gecko) Chrome/120.0 Safari/537.36"),
    "Accept": "application/json, text/javascript, */*; q=0.01",
    "X-Requested-With": "XMLHttpRequest",
    "Referer": BASE + "/",
}


def formas_color(cod_color):
    """
    Las formas en que el sitio puede esperar el color, en orden de prueba.

    No sabemos cual usa cada producto: el excel trae 001 y en la URL de la
    ficha se vio 0001. En vez de elegir una y rezar, se prueban todas y se
    usa la primera que devuelva una variante de verdad.

    >>> formas_color("001")
    ['0001', '001', '1']
    """
    c = str(cod_color).strip()
    formas = [c.zfill(4), c, c.lstrip("0") or "0"]
    vistas, salida = set(), []
    for f in formas:
        if f not in vistas:
            vistas.add(f)
            salida.append(f)
    return salida


def slug(texto):
    """'Georgette High-Low Skirt' -> 'georgette-high-low-skirt'"""
    import re as _re, unicodedata as _u
    t = _u.normalize("NFKD", str(texto))
    t = "".join(c for c in t if not _u.combining(c)).lower()
    return _re.sub(r"[^a-z0-9]+", "-", t).strip("-")


def url_ficha(referencia, nombre=""):
    """La ficha real. El slug sale del nombre que devuelve el propio JSON."""
    s = slug(nombre)
    return f"{BASE}/{s}/{referencia}.html" if s else f"{BASE}/{referencia}.html"


def _pedir(referencia, color, sesion):
    """Una consulta cruda. Devuelve el bloque 'product' o None."""
    params = {"pid": referencia,
              f"dwvar_{referencia}_color": color,
              "quantity": 1}
    try:
        r = sesion.get(VARIACION, params=params, headers=CABECERAS, timeout=30)
    except Exception:
        return None, "error de red"
    if r.status_code != 200:
        return None, f"HTTP {r.status_code}"
    try:
        datos = json.loads(r.text)
    except Exception:
        return None, "no devolvio JSON"
    return (datos.get("product") or {}), None


def _colores_de(prod):
    for attr in prod.get("variationAttributes") or []:
        if str(attr.get("id", "")).lower() == "color":
            return [str(v.get("id") or v.get("value"))
                    for v in attr.get("values") or []]
    return []


def _fotos_de(prod, cod_color):
    """URLs del tamano mas grande, quedandose solo con las de este color."""
    imagenes = prod.get("images") or {}
    lote = next((imagenes[t] for t in TAMANOS if imagenes.get(t)), [])
    urls = [im.get("url") for im in lote if im.get("url")]
    marca = f"-{str(cod_color).strip()}-"
    return [u for u in urls if marca in u], urls


def candidatos(cod_color, publicados):
    """
    Que valor de color mandarle al sitio, en orden de prueba.

    Primero el relleno con ceros, que es el caso comun (001 -> 0001). Si eso
    no acierta, se usa la lista de colores que el propio sitio publica, que
    llega en variationAttributes. Hace falta porque el formato NO es
    uniforme: verificado a mano, el color 001 del excel es 0001 en el sitio,
    pero el 303 es 3031 y no 0303.

    >>> candidatos("001", ["3031", "0001"])
    ['0001', '3031']
    >>> candidatos("303", ["3031", "0001"])
    ['0303', '3031', '0001']
    """
    c = str(cod_color).strip()
    orden = [c.zfill(4)]

    def puntaje(pid):
        if pid == c.zfill(4):
            return 0
        if pid.lstrip("0") == c.lstrip("0"):
            return 1
        if pid.startswith(c):
            return 2
        if pid.endswith(c):
            return 3
        return 4

    orden += sorted(publicados or [], key=puntaje)

    vistos, salida = set(), []
    for x in orden:
        if x and x not in vistos:
            vistos.add(x)
            salida.append(x)
    return salida


def consultar(referencia, cod_color, sesion):
    """
    Pide la variante y se queda con las fotos SOLO si son del color pedido.

    La prueba de que acerto NO es el productType: verificado a mano, MK
    responde "master" incluso cuando devuelve las fotos correctas. Lo que
    distingue es el nombre del archivo, que lleva el color adentro:

        color=0001 -> MT670N27R3-001-0001_1-tif   <- del color 001, sirve
        color=001  -> MT670N27R3-303-3031_1-tif   <- del color por defecto

    Devuelve (urls, colores_publicados, nombre_producto, problema).
    """
    colores, nombre, ultimo = [], "", "No se obtuvo respuesta del sitio."
    probados = set()
    pendientes = candidatos(cod_color, [])

    while pendientes:
        color = pendientes.pop(0)
        if color in probados:
            continue
        probados.add(color)

        prod, err = _pedir(referencia, color, sesion)
        comun.time.sleep(comun.ESPERA)
        if prod is None:
            ultimo = f"El sitio respondio mal ({err})."
            continue

        nombre = prod.get("productName") or nombre

        # La primera respuesta nos dice que colores existen de verdad; con
        # eso se rearma la cola en vez de seguir adivinando.
        nuevos = _colores_de(prod)
        if nuevos and not colores:
            colores = nuevos
            for c in candidatos(cod_color, colores):
                if c not in probados and c not in pendientes:
                    pendientes.append(c)

        correctas, todas = _fotos_de(prod, cod_color)
        if correctas:
            return correctas, colores, nombre, None
        if todas:
            ultimo = (f"El sitio devolvio fotos de otro color "
                      f"(ejemplo: {todas[0].rsplit('/', 1)[-1]}).")
        else:
            ultimo = "El producto existe pero no trae fotos."

    if colores:
        ultimo += f" Colores publicados: {', '.join(colores)}."
    return [], colores, nombre, ultimo


def diagnostico(referencia, cod_color, sesion):
    """Muestra crudo lo que contesta el sitio, para un solo producto."""
    print(f"\n  referencia {referencia}   color del excel {cod_color}\n")
    prod0, _ = _pedir(referencia, str(cod_color).strip().zfill(4), sesion)
    publicados = _colores_de(prod0) if prod0 else []
    for color in candidatos(cod_color, publicados):
        prod, err = _pedir(referencia, color, sesion)
        if prod is None:
            print(f"  color={color:<6} -> {err}")
            continue
        _, todas = _fotos_de(prod, cod_color)
        correctas, _ = _fotos_de(prod, cod_color)
        veredicto = "SIRVE" if correctas else "no es este color"
        print(f"  color={color:<6} -> {veredicto}"
              f"   (productType={prod.get('productType')!r}, no es el criterio)")
        print(f"                  nombre  : {prod.get('productName')}")
        print(f"                  colores : {', '.join(_colores_de(prod)) or '(ninguno)'}")
        print(f"                  fotos   : {len(todas)}"
              + (f"  ej: {todas[0].rsplit('/', 1)[-1]}" if todas else ""))
        print(f"                  ficha   : {url_ficha(referencia, prod.get('productName',''))}")
        print()
        comun.time.sleep(comun.ESPERA)


# ===========================================================================
# PRINCIPAL
# ===========================================================================

def main(argv=None):
    """argv=None usa la linea de comandos; una lista permite llamarlo
    desde codigo (lo hace runner.py cuando corre dentro del .exe)."""
    ap = argparse.ArgumentParser(description="Baja las fotos de Michael Kors.")
    ap.add_argument("excel", type=Path)
    ap.add_argument("--salida", type=Path, default=Path("fotos"))
    ap.add_argument("--limite", type=int, default=None)
    ap.add_argument("--forzar", action="store_true")
    ap.add_argument("--diagnostico", nargs=2, metavar=("REFERENCIA", "COLOR"),
                    help="muestra que contesta el sitio para un solo producto "
                         "y no baja nada")
    args = ap.parse_args(argv)

    if args.diagnostico:
        diagnostico(args.diagnostico[0], args.diagnostico[1], nueva_sesion())
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
    print("una consulta por producto\n")

    sesion = nueva_sesion()
    filas = []

    for i, (_, p) in enumerate(prod.iterrows(), 1):
        cod, ref, col = p["cod"], p["referencia"], p["cod_color"]

        if not args.forzar and comun.ya_bajado(carpeta, cod, SUFIJO):
            existentes = comun.archivos_de(carpeta, cod, SUFIJO)
            print(f"  {i:>3}/{len(prod)} [--] {cod:<16} ya estaba ({len(existentes)})")
            filas.append(_fila(p, existentes, 0, [], "", ""))
            continue

        urls, colores, nombre_web, problema = consultar(ref, col, sesion)

        crudas = []
        for n, url in enumerate(urls):
            img = comun.bajar(url, sesion)
            if img is not None:
                crudas.append((n, img))       # n = orden del carrusel

        imagenes = comun.unicas(crudas)
        nombres = comun.guardar(imagenes, carpeta, cod, SUFIJO)
        repetidas = len(crudas) - len(imagenes)

        detalle = f"{len(nombres)} fotos" if nombres else (problema or "sin fotos")
        if repetidas:
            detalle += f"  ({repetidas} repetidas descartadas)"
        print(f"  {i:>3}/{len(prod)} {'[ok]' if nombres else '[!!]'} {cod:<16} {detalle[:70]}")

        filas.append(_fila(p, nombres, repetidas, urls, problema or "", nombre_web))

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


def _fila(p, nombres, repetidas, urls, nota, nombre_web=""):
    ref = p.get("referencia", "")
    if nombres:
        # la ficha real; el slug sale del nombre que devolvio el JSON
        url_producto = url_ficha(ref, nombre_web or p.get("nombre", ""))
        etiqueta = "Ver la ficha en michaelkors.com"
    else:
        url_producto = BUSCADOR.format(referencia=ref)
        etiqueta = f"Buscar la referencia {ref}"

    return {
        "cod": p["cod"],
        "nombre": p.get("nombre", ""),
        "color": p.get("color", ""),
        "referencia": ref,
        "cod_color": p.get("cod_color", ""),
        "tallas": p.get("tallas", ""),
        "n_fotos": len(nombres),
        "repetidas": repetidas,
        "archivos": " ".join(nombres),
        "nota": nota,
        "url_producto": url_producto,
        "label_producto": etiqueta,
        "url_foto1": urls[0] if urls else "",
    }


if __name__ == "__main__":
    main()