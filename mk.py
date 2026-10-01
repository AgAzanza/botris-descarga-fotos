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
import re
from datetime import datetime
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

# MK tiene un catalogo DISTINTO POR PAIS. El mismo producto puede existir
# en el Reino Unido y no en Estados Unidos, o tener alli colores que aca no.
# Verificado: el 43F5DTFS6L color 150 no esta en el sitio de EE.UU. (que
# solo publica el 0632) pero si en michaelkors.co.uk. Consultando un solo
# pais, esos productos salian como "no existe" cuando si existen.
#
# Se prueban en orden y se corta en el primero que devuelva las fotos del
# color pedido.
SITIOS = [
    ("https://www.michaelkors.com",    "mk_us", "en_US"),
    ("https://www.michaelkors.co.uk",  "mk_uk", "en_GB"),
]

BASE = SITIOS[0][0]
VARIACION = "{base}/on/demandware.store/Sites-{sitio}-Site/{locale}/Product-Variation"
FICHA_WEB = "{base}/{referencia}.html"
BUSCADOR = BASE + "/search?q={referencia}"

# Cual de los tamanos usamos cuando las URLs vienen del JSON, en orden de
# preferencia. Son los nombres de los campos de la respuesta.
TAMANOS = ["zoomHiRes", "zoom", "large", "base"]

# El nombre del tamano DENTRO DE LA URL del CDN, que es OTRO. En la
# respuesta JSON el campo se llama "zoomHiRes" pero en la direccion la
# misma foto va como ECOM_Image_Zoom_Highres:
#
#   assets.michaelkors.com/transform/ECOM_Image_Zoom_Highres/{uuid}/{nombre}
#
# Confundirlos hace que las fotos sacadas de la ficha den 404.
PRESET_CDN = "ECOM_Image_Zoom_Highres"

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


def _pedir(referencia, color, sesion, sitio=None):
    """Una consulta cruda a un pais. Devuelve el bloque 'product' o None."""
    base, id_sitio, locale = sitio or SITIOS[0]
    url = VARIACION.format(base=base, sitio=id_sitio, locale=locale)
    params = {"pid": referencia,
              f"dwvar_{referencia}_color": color,
              "quantity": 1}
    try:
        r = sesion.get(url, params=params, headers=CABECERAS, timeout=30)
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


# En la ficha, las fotos van en src normal: .../{ref}-{c3}-{c4}_{n}-tif
_FOTO = r"assets\.michaelkors\.com/transform/([^/\"']+)/([0-9a-f-]{{8,}})/{ref}-{col}-(\d+)_(\d+)-tif"


def _de_la_ficha(referencia, cod_color, sesion, sitio):
    """
    Respaldo: leer la ficha del producto en vez del JSON.

    La pagina trae las fotos del color seleccionado en src normal, y los
    botones de color revelan que colores existen. Sirve cuando el endpoint
    de variantes no resuelve, que pasa segun el pais.

    Devuelve (urls, colores_vistos, mapa) donde mapa es {color3: color4}.

    EL MAPA ES LO MAS VALIOSO DE LA PAGINA. El nombre de cada imagen es
    {ref}-{c3}-{c4}_{n}-tif, o sea que la propia URL dice cual es el codigo
    de cuatro digitos de cada color:

        30R5G9IS6L-156-0250_1-tif   ->  156 es 0250
        32R6GY5W6B-252-1335_1-tif   ->  252 es 1335

    Sin eso hay que adivinar, y adivinar falla: medido sobre 15 productos
    reales, en 3 el c4 no era el color con ceros adelante. Con el mapa, al
    endpoint de variantes se le manda el codigo correcto a la primera.
    """
    base = sitio[0]
    url = FICHA_WEB.format(base=base, referencia=referencia)
    try:
        r = sesion.get(url, params={f"dwvar_{referencia}_color": cod_color},
                       headers={**CABECERAS, "Accept": "text/html"},
                       timeout=30, allow_redirects=True)
    except Exception:
        return [], [], {}
    if r.status_code != 200:
        return [], [], {}
    html = r.text

    # que colores muestra la ficha, y el codigo de 4 digitos de cada uno
    mapa = {}
    for c3, c4 in re.findall(
            rf"{re.escape(referencia)}-(\d+)-(\d+)_\d+-tif", html):
        mapa.setdefault(c3, c4)
    vistos = sorted(mapa)

    # las fotos grandes de NUESTRO color, sin repetir y en orden
    patron = re.compile(_FOTO.format(ref=re.escape(referencia),
                                     col=re.escape(str(cod_color).strip())))
    urls, claves = [], set()
    for m in patron.finditer(html):
        preset, uuid, c4, n = m.groups()
        if "Swatch" in preset or "Thumbnail" in preset:
            continue
        clave = (uuid, n)
        if clave in claves:
            continue
        claves.add(clave)
        urls.append(f"https://assets.michaelkors.com/transform/"
                    f"{PRESET_CDN}/{uuid}/{referencia}-{cod_color}-{c4}_{n}-tif")
    urls.sort(key=lambda u: int(re.search(r"_(\d+)-tif$", u).group(1)))
    return urls, vistos, mapa


def consultar(referencia, cod_color, sesion):
    """
    Busca las fotos del color pedido, en cada pais, por los dos caminos.

    ORDEN Y POR QUE:

    1. LA FICHA, para conseguir el MAPA de colores. El nombre de cada
       imagen dice el codigo de cuatro digitos de su color
       (30R5G9IS6L-156-0250_1-tif: el 156 es 0250). Sin ese dato hay que
       adivinar, y medido sobre 15 productos reales, en 3 el codigo no era
       el color con ceros adelante.

    2. EL ENDPOINT DE VARIANTES, ya con el codigo correcto. Es el que
       entrega la galeria completa del color.

    3. SE QUEDA CON LA LISTA MAS LARGA, entre los dos caminos Y ENTRE LOS
       DOS PAISES. La ficha suele traer UNA sola imagen del color pedido
       (el botoncito), y quedarse con eso daba 1 foto cuando habia 5.

       Por eso UNA SOLA FOTO NO CORTA LA BUSQUEDA: se guarda como lo mejor
       hasta ahora y se sigue con el otro pais. Caso real: el 43F5DTFS6L
       color 150 se vende en el Reino Unido, pero Estados Unidos muestra
       igual su botoncito; cortando ahi se entregaba 1 foto en vez de las
       4 que tiene.

    El criterio de acierto es siempre el mismo: que el NOMBRE del archivo
    contenga el color. MK responde 200 y "master" aunque acierte, asi que
    el tipo no sirve.

    Devuelve (urls, colores_publicados, nombre_producto, problema).
    """
    color = str(cod_color).strip()
    colores, nombre = [], ""
    ultimo = "No se obtuvo respuesta del sitio."
    mejor_global = []

    for base, pais, locale in SITIOS:
        sitio = (base, pais, locale)

        # --- 1. la ficha: fotos sueltas y, sobre todo, el mapa de colores ---
        urls_ficha, vistos, mapa = _de_la_ficha(referencia, color, sesion, sitio)
        comun.time.sleep(comun.ESPERA)
        if vistos:
            colores = sorted(set(colores) | set(vistos))
            if color not in vistos:
                ultimo = (f"{pais} no publica el color {color} de esta "
                          f"referencia (tiene: {', '.join(vistos)})")

        # --- 2. el endpoint, con el codigo que dijo la ficha ---
        #
        # Solo si hace falta: cuando la ficha ya trajo una galeria (2 o mas
        # fotos) no tiene sentido preguntar de nuevo. Con una sola foto si
        # se consulta, porque esa suele ser el botoncito de color.
        orden = candidatos(color, vistos)
        if color in mapa:
            orden.insert(0, mapa[color])      # el correcto, primero

        urls_endpoint, probados = [], set()
        for c in ([] if len(urls_ficha) > 1 else orden):
            if c in probados:
                continue
            probados.add(c)

            prod, err = _pedir(referencia, c, sesion, sitio)
            comun.time.sleep(comun.ESPERA)
            if prod is None:
                ultimo = f"{pais}: el sitio respondio mal ({err})."
                continue

            nombre = prod.get("productName") or nombre
            nuevos = _colores_de(prod)
            if nuevos:
                colores = sorted(set(colores) | set(nuevos))

            correctas, todas = _fotos_de(prod, color)
            if correctas:
                urls_endpoint = correctas
                break
            if todas:
                ultimo = (f"{pais}: devolvio fotos de otro color "
                          f"(ejemplo: {todas[0].rsplit('/', 1)[-1]}).")

        # --- 3. la lista mas larga gana ---
        mejor = max((urls_endpoint, urls_ficha), key=len)
        if len(mejor) > len(mejor_global):
            mejor_global = mejor
        # con 2 o mas es una galeria de verdad y no hace falta seguir;
        # con una sola, casi seguro es el botoncito: se prueba el otro pais
        if len(mejor_global) > 1:
            return mejor_global, colores, nombre, None

    if mejor_global:
        return mejor_global, colores, nombre, None

    if colores:
        ultimo += f". Colores publicados: {', '.join(colores)}."
    return [], colores, nombre, ultimo


def diagnostico(referencia, cod_color, sesion):
    """Muestra que contesta cada pais y cada camino, para un solo producto."""
    color = str(cod_color).strip()
    print(f"\n  referencia {referencia}   color del excel {color}\n")

    for base, pais, locale in SITIOS:
        sitio = (base, pais, locale)
        print(f"  --- {pais}  ({base}) ---")

        urls, vistos, mapa = _de_la_ficha(referencia, color, sesion, sitio)
        comun.time.sleep(comun.ESPERA)
        print(f"    ficha                 -> {len(urls)} fotos de este color")
        print(f"    codigo real del color -> {mapa.get(color, '(no aparece)')}")
        print(f"    colores que publica   -> {', '.join(vistos) or '-'}")
        for u in urls[:6]:
            print(f"       {u.rsplit('/', 1)[-1]}")

        orden = candidatos(color, vistos)
        if color in mapa:
            orden.insert(0, mapa[color])
        vistos_ya = set()
        for c in orden[:4]:
            if c in vistos_ya:
                continue
            vistos_ya.add(c)
            prod, err = _pedir(referencia, c, sesion, sitio)
            comun.time.sleep(comun.ESPERA)
            if prod is None:
                print(f"    variantes color={c:<6} -> {err}")
                continue
            correctas, todas = _fotos_de(prod, color)
            print(f"    variantes color={c:<6} -> "
                  f"{len(correctas)} de este color, {len(todas)} en total"
                  f"   (tipo={prod.get('productType')})"
                  + (f"   ej: {todas[0].rsplit('/', 1)[-1]}" if todas else ""))
            if correctas:
                break

        print(f"    ficha: {FICHA_WEB.format(base=base, referencia=referencia)}")
        print()

    u, cols, nom, prob = consultar(referencia, color, sesion)
    print(f"  RESULTADO FINAL: {len(u)} fotos" + (f"   |  {prob}" if prob else ""))
    for x in u:
        print(f"     {x}")
    print()


def reporte_diagnostico(excel, sesion, destino, limite=None):
    """
    Genera un archivo con TODO lo que devolvio el sitio, producto por
    producto, para poder revisar si el metodo esta funcionando.

    El metodo de cada marca se construyo mirando UN producto. Este reporte
    existe para dejar de suponer: muestra, por cada articulo del excel y
    por cada pais, que respondio la ficha, que colores publica, que
    devolvio el endpoint y con que se quedo el programa.
    """
    prod = comun.cargar(excel, marca=MARCA)
    if limite:
        prod = prod.head(limite)

    lineas = [
        "REPORTE DE DIAGNOSTICO - MICHAEL KORS",
        f"archivo : {Path(excel).name}",
        f"fecha   : {datetime.now():%Y-%m-%d %H:%M}",
        f"paises  : {', '.join(p for _, p, _ in SITIOS)}",
        f"productos: {len(prod)}",
        "=" * 78, "",
    ]

    con, sin = 0, 0
    for i, (_, p) in enumerate(prod.iterrows(), 1):
        ref, col = p["referencia"], p["cod_color"]
        lineas += [f"[{i}/{len(prod)}] {p['cod']}",
                   f"  excel    : ref={ref}  color={col}  "
                   f"nombre={p.get('nombre','')}  color_texto={p.get('color','')}"]

        for base, pais, locale in SITIOS:
            sitio = (base, pais, locale)

            urls, vistos, mapa = _de_la_ficha(ref, col, sesion, sitio)
            comun.time.sleep(comun.ESPERA)
            lineas.append(f"  {pais} ficha    : {len(urls)} fotos de este color"
                          f" | codigo real del color: {mapa.get(str(col).strip(), '-')}"
                          f" | colores: {', '.join(vistos) or '-'}")
            lineas.append(f"    url      : {FICHA_WEB.format(base=base, referencia=ref)}")
            for u in urls[:4]:
                lineas.append(f"      {u}")

            orden = candidatos(col, vistos)
            if str(col).strip() in mapa:
                orden.insert(0, mapa[str(col).strip()])
            for c in orden[:3]:
                prodj, err = _pedir(ref, c, sesion, sitio)
                comun.time.sleep(comun.ESPERA)
                if prodj is None:
                    lineas.append(f"  {pais} variantes color={c}: {err}")
                    continue
                correctas, todas = _fotos_de(prodj, col)
                lineas.append(
                    f"  {pais} variantes color={c}: tipo={prodj.get('productType')}"
                    f" | fotos_de_este_color={len(correctas)} | total={len(todas)}"
                    f" | colores={', '.join(_colores_de(prodj)) or '-'}")
                if todas:
                    lineas.append(f"      ejemplo: {todas[0].rsplit('/', 1)[-1]}")
                if correctas:
                    break

        u, cs, nom, prob = consultar(ref, col, sesion)
        lineas += [f"  RESULTADO: {len(u)} fotos"
                   + (f"  |  {prob}" if prob else ""), ""]
        con += 1 if u else 0
        sin += 0 if u else 1

    lineas += ["=" * 78,
               f"con fotos: {con}   sin fotos: {sin}   de {len(prod)}"]

    destino = Path(destino)
    destino.write_text("\n".join(lineas), encoding="utf-8")
    return destino, con, sin


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
    ap.add_argument("--reporte", action="store_true",
                    help="genera diagnostico-mk.txt con todo lo que devolvio "
                         "el sitio para cada producto del excel, sin bajar nada")
    ap.add_argument("--auditoria", action="store_true",
                    help="genera auditoria-mk.json con toda la evidencia cruda "
                         "de varios productos, para revisar el metodo")
    ap.add_argument("--muestra", type=int, default=12,
                    help="cuantos productos audita (por defecto 12)")
    args = ap.parse_args(argv)

    if args.diagnostico:
        diagnostico(args.diagnostico[0], args.diagnostico[1], nueva_sesion())
        return

    if args.reporte:
        destino = args.excel.parent / "diagnostico-mk.txt"
        print(f"\n  revisando {args.limite or 'todos los'} productos, "
              f"esto tarda...")
        d, con, sin = reporte_diagnostico(args.excel, nueva_sesion(),
                                          destino, args.limite)
        print(f"  con fotos: {con}  |  sin fotos: {sin}")
        print(f"\n  Listo: {d.resolve()}")
        return

    prod = comun.cargar(args.excel, marca=MARCA)

    if args.auditoria:
        if len(prod) == 0:
            print("El excel no tiene productos MK.")
            return
        print(f"\nAuditando {min(args.muestra, len(prod))} de {len(prod)} "
              f"productos. Tarda un rato.\n")
        destino = auditoria(prod, nueva_sesion(),
                            args.salida / "auditoria-mk.json", args.muestra)
        print(f"\n  Listo: {destino.resolve()}")
        print("  Ese archivo es el que hay que revisar.")
        return
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